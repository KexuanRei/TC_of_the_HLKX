# -*- coding: utf-8 -*-
"""
window.py —— 把横道图控制台装进一个原生窗口（pywebview + Edge WebView2）

双击「窗口.bat」运行。它做三件事：

    1. 起本地服务（server.py 的 HTTP 服务，同时是 Webhook 推送接收器）；
    2. 用 pywebview 开一个原生窗口，加载 http://127.0.0.1:<port>/app；
    3. 关窗口 = 退进程（服务是守护线程，不会留后台）。

没有 pywebview 时：先试一次 `pip install pywebview`；还不行就退回 Edge 应用窗口
（`msedge --app=…`，同样没有地址栏/标签页，是一个独立窗口）。

参数：
    --port N        指定端口（默认取 config.json）
    --debug         打开 WebView 调试
    --no-window     只起服务，不开窗口
    --selftest      开窗 4 秒后自动关闭，验证环境可用
"""
from __future__ import annotations

import argparse
import io
import json
import os
import socket
import subprocess
import sys
import threading
import time
import traceback
import urllib.error
import urllib.request

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import paths

BASE = paths.BASE
LOG_PATH = os.path.join(BASE, "logs", "window.log")
PANEL_NAME = "gantt-window"
ICON_PATH = paths.res("TM.ico")               # 窗口左上角（标题栏）图标


# ---------------------------------------------------------------- 日志

class _Tee(io.TextIOBase):
    """stdout 同时写控制台和 logs/window.log —— 控制台被藏起来时还能事后查。"""

    def __init__(self, path):
        self.fh = None
        self.orig = sys.stdout
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            self.fh = io.open(path, "a", encoding="utf-8")
        except OSError:
            self.fh = None

    def write(self, s):
        if self.orig is not None:
            try:
                self.orig.write(s)
            except (ValueError, OSError):
                pass
        if self.fh is not None:
            try:
                self.fh.write(s)
                self.fh.flush()
            except (ValueError, OSError):
                pass
        return len(s) if isinstance(s, str) else 0

    def flush(self):
        for target in (self.orig, self.fh):
            if target is not None:
                try:
                    target.flush()
                except (ValueError, OSError):
                    pass


def log(msg):
    print(time.strftime("[%H:%M:%S] ") + str(msg), flush=True)


def hide_console_if_requested():
    if os.name != "nt" or os.environ.get("GANTT_HIDE_CONSOLE") != "1":
        return
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)      # SW_HIDE
    except Exception:                                      # noqa: BLE001
        pass


def alert(msg, title="分段生产横道图"):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(msg)[:2000], title, 0x10)
    except Exception:                                      # noqa: BLE001
        pass


# ---------------------------------------------------------------- 服务

def probe_panel(port, timeout=0.8):
    try:
        req = urllib.request.Request("http://127.0.0.1:%d/api/version" % int(port),
                                     headers={"X-Panel": "1"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8", "replace"))
        if isinstance(data, dict) and data.get("name") == PANEL_NAME:
            return data
    except (urllib.error.URLError, OSError, ValueError):
        pass
    return None


def port_free(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind(("127.0.0.1", int(port)))
            return True
        except OSError:
            return False


def start_server(port=None):
    import server

    cfg_path = os.path.join(BASE, "config.json")
    cfg = server.load_config(cfg_path)
    want = int(port or cfg.get("port") or 8089)

    httpd = None
    for candidate in range(want, want + 25):
        if not port_free(candidate):
            continue
        try:
            httpd, cfg = server.create_server(port=candidate, config_path=cfg_path)
            break
        except OSError:
            continue
    if httpd is None:
        raise RuntimeError("从 %d 起找不到可用端口" % want)

    threading.Thread(target=httpd.serve_forever, name="panel-http", daemon=True).start()
    log("服务已启动：http://127.0.0.1:%d/  （控制台 /app）" % int(cfg["port"]))
    return httpd, cfg, int(cfg["port"])


def wait_http(port, timeout=8.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if probe_panel(port):
            return True
        time.sleep(0.15)
    return False


# ---------------------------------------------------------------- pywebview 桥

class Bridge:
    def __init__(self, app_url):
        self.app_url = app_url

    def open_external(self, url):
        url = str(url or "").strip()
        if not url or not (url.startswith("http://") or url.startswith("https://")):
            return False
        try:
            import webbrowser
            webbrowser.open(url)
            return True
        except Exception:                                  # noqa: BLE001
            return False

    def pick_file(self, directory=""):
        """打开系统的文件选择框，让用户挑一张本地表。返回完整路径，取消返回空串。"""
        directory = (directory or "").strip()
        if not os.path.isdir(directory):
            directory = os.path.dirname(BASE) if os.path.isdir(os.path.dirname(BASE)) else BASE
        file_types = ("Excel 表格 (*.xlsx;*.xlsm;*.xls)", "所有文件 (*.*)")

        # 首选 pywebview 自带的选择框（和窗口同进程，最顺）
        try:
            import webview
            windows = webview.windows or []
            if windows:
                result = windows[0].create_file_dialog(
                    webview.OPEN_DIALOG, directory=directory,
                    allow_multiple=False, file_types=file_types)
                if isinstance(result, (list, tuple)):
                    return result[0] if result else ""
                return result or ""
        except Exception as exc:                           # noqa: BLE001
            log("pywebview 文件选择框失败，改用 tkinter：%r" % (exc,))

        # 兜底：tkinter 的原生文件对话框（一定存在，也一定弹得出资源管理器）
        return self._tk_pick_file(directory, file_types)

    def _tk_pick_file(self, directory, file_types):
        try:
            import tkinter as tk
            from tkinter import filedialog
            types = [("Excel 表格", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")]
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            try:
                path = filedialog.askopenfilename(
                    initialdir=directory or None, title="选择本地 Excel 表", filetypes=types)
            finally:
                root.destroy()
            return path or ""
        except Exception as exc:                           # noqa: BLE001
            log("tkinter 文件选择框也失败：%r" % (exc,))
            return ""

    def version(self):
        return {"panel": "1.0", "python": sys.version.split()[0], "url": self.app_url}

    def quit(self):
        try:
            import webview
            for w in list(webview.windows or []):
                w.destroy()
        except Exception:                                  # noqa: BLE001
            pass
        return True


# ---------------------------------------------------------------- 入口

def find_edge():
    roots = [os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"),
             os.environ.get("LOCALAPPDATA")]
    for root in roots:
        if not root:
            continue
        path = os.path.join(root, r"Microsoft\Edge\Application\msedge.exe")
        if os.path.exists(path):
            return path
    return None


def open_edge_app(url, size=(1560, 960)):
    """用 Edge 的「应用窗口」打开界面；返回 Popen（关窗口后进程结束）。"""
    edge = find_edge()
    if not edge:
        return None
    profile = os.path.join(os.environ.get("TEMP") or BASE, "gantt-window-profile")
    try:
        return subprocess.Popen([edge, "--app=%s" % url, "--window-size=%d,%d" % size,
                                 "--user-data-dir=%s" % profile])
    except OSError:
        return None


def ensure_webview():
    try:
        import webview
        return webview
    except ImportError:
        pass
    if paths.FROZEN:
        log("打包版未内置 pywebview，改用 Edge 应用窗口")
        return None
    log("第一次用窗口界面：正在安装 pywebview（需要联网，几十秒）…")
    try:
        subprocess.call([sys.executable, "-m", "pip", "install", "pywebview"])
    except OSError as exc:
        log("安装失败：%r" % (exc,))
    try:
        import webview
        return webview
    except ImportError:
        return None


def run_gantt_sync():
    """定时任务入口：拉数出图后退出（不打开窗口）。"""
    paths.ensure_defaults()
    try:
        import gantt_sync
        conf = gantt_sync.load_config(gantt_sync.CONFIG, {})
        gantt_sync.build(conf)
        return 0
    except SystemExit as exc:
        log("定时取数失败：%s" % (exc.code if exc.code is not None else exc))
        return 1
    except Exception:                                      # noqa: BLE001
        traceback.print_exc()
        return 1


def main():
    ap = argparse.ArgumentParser(description="分段生产横道图 · 窗口界面")
    ap.add_argument("--port", type=int)
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--no-window", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--gantt-sync", dest="gantt_sync", action="store_true",
                    help="定时任务用：拉数出图后退出")
    args = ap.parse_args()

    sys.stdout = _Tee(LOG_PATH)
    if args.gantt_sync:
        return run_gantt_sync()
    if not args.debug and not args.no_window and not args.selftest:
        hide_console_if_requested()

    paths.ensure_defaults()
    log("启动窗口界面…")

    want = int(args.port or 8089)
    try:
        cfg = json.load(io.open(os.path.join(BASE, "config.json"), encoding="utf-8"))
        want = int(args.port or cfg.get("port") or 8089)
    except (OSError, ValueError):
        pass

    httpd = None
    if probe_panel(want):
        port = want
        log("端口 %d 上已有窗口服务，直接复用（数据共享）" % port)
    else:
        httpd, cfg, port = start_server(args.port)
        if not wait_http(port):
            raise RuntimeError("服务没起来，端口 %d 无响应" % port)

    app_url = "http://127.0.0.1:%d/app" % port

    if args.no_window:
        log("只起服务模式：%s   （Ctrl+C 退出）" % app_url)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            log("收到 Ctrl+C，退出")
        return 0

    webview = ensure_webview()
    if webview is None:
        log("没有 pywebview，改用 Edge 应用窗口")
        proc = open_edge_app(app_url)
        if proc is not None:
            log("已用 Edge 打开窗口：%s（关掉窗口即退出）" % app_url)
            try:
                proc.wait()
            except KeyboardInterrupt:
                pass
            log("Edge 窗口已关闭，退出")
            return 0
        import webbrowser
        webbrowser.open(app_url)
        log("系统浏览器也打不开的话，手动访问：%s" % app_url)
        return 1

    bridge = Bridge(app_url)
    window = webview.create_window(
        "分段生产横道图 · 控制台",
        app_url,
        width=1560, height=960, min_size=(1120, 700),
        js_api=bridge, text_select=False, background_color="#eef2f8",
    )
    log("窗口已创建，加载 %s" % app_url)

    if args.selftest:
        def closer(win):
            time.sleep(4)
            try:
                win.destroy()
            except Exception:                              # noqa: BLE001
                pass
            log("selftest：窗口已自动关闭")

        webview.start(closer, (window,), debug=False, icon=ICON_PATH if os.path.exists(ICON_PATH) else None)
        log("selftest OK")
        return 0

    webview.start(debug=bool(args.debug), icon=ICON_PATH if os.path.exists(ICON_PATH) else None)
    log("窗口已关闭，退出")
    if httpd is not None:
        try:
            httpd.shutdown()
        except Exception:                                  # noqa: BLE001
            pass
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:                                      # noqa: BLE001
        tb = traceback.format_exc()
        try:
            print(tb, flush=True)
        except (ValueError, OSError):
            pass
        try:
            os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
            with io.open(LOG_PATH, "a", encoding="utf-8") as fh:
                fh.write(tb + "\n")
        except OSError:
            pass
        alert("窗口启动失败：\n\n" + tb.strip().splitlines()[-1] +
              "\n\n详细日志：\n" + LOG_PATH)
        sys.exit(1)
