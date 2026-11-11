# -*- coding: utf-8 -*-
"""
server.py —— 横道图 Webhook 服务

    POST /webhook/gantt    JSON: {"rows":[...]} 或 [...]        → 重新生成 data.js
    POST /webhook/csv      text/csv（首行表头）                  → 重新生成 data.js
    GET  /                 查看器（viewer.html）
    GET  /app              窗口界面（app.html，由 window.py 装进原生窗口）
    GET  /data.js          当前数据
    GET  /health           健康检查
    GET  /api/status       最近一次构建信息
    *    /api/*            窗口界面的后端（见 panel.py，仅本机回环 + X-Panel: 1）

鉴权：请求头 X-Token，或 ?token=xxx。config.json 里 token 为空则不校验（仅建议本机）。
/api/* 不走 token：只认本机回环地址，且要求 X-Panel: 1 头（浏览器跨站请求会带 Origin，被拒）。
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, urlsplit

import core
import panel
import paths

BASE = paths.BASE
DEFAULT_CONFIG = {
    "host": "0.0.0.0",
    "port": 8089,
    "token": "",
    "data": "data.js",
    "viewer": "viewer.html",
    "archive": "incoming",
    "keep": 30,
    "max_mb": 32,
    "edits": "edits.json",
}

STATE = {"last": None, "error": None, "builds": 0}
LOCK = threading.Lock()
CFG = dict(DEFAULT_CONFIG)


def log(msg):
    line = "%s  %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        with io.open(os.path.join(BASE, "service.log"), "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass


def load_config(path):
    cfg = dict(DEFAULT_CONFIG)
    if path and os.path.exists(path):
        try:
            with io.open(path, encoding="utf-8") as fh:
                cfg.update({k: v for k, v in json.load(fh).items() if v is not None})
        except Exception as exc:  # noqa: BLE001
            log("读取配置失败，用默认值：%r" % (exc,))
    return cfg


def archive(payload, kind):
    """把原始 payload 存一份，方便回溯。"""
    folder = os.path.join(BASE, CFG["archive"])
    os.makedirs(folder, exist_ok=True)
    name = "%s-%s.%s" % (dt.datetime.now().strftime("%Y%m%d-%H%M%S"), kind,
                         "json" if kind else "json")
    path = os.path.join(folder, name)
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False)
        files = sorted(os.listdir(folder))
        for old in files[: max(0, len(files) - int(CFG.get("keep", 30)))]:
            os.remove(os.path.join(folder, old))
    except OSError as exc:
        log("归档失败：%r" % (exc,))
    return path


def rebuild(rows, source):
    payload = core.build_payload(rows, source)
    if not payload["blocks"]:
        raise ValueError("没有可用数据：所有行都缺船号或缺日期")
    size = core.write_js(payload, os.path.join(BASE, CFG["data"]))
    info = {
        "ok": True,
        "blocks": len(payload["blocks"]),
        "hulls": len(payload["hulls"]),
        "bytes": size,
        "gen": payload["gen"],
        "source": source,
    }
    STATE["last"] = info
    STATE["builds"] += 1
    log("构建成功：%d 段 / %d 船 / %.2f MB" % (info["blocks"], info["hulls"], size / 1048576.0))
    return info


# ---------- 进度条操作存档（本地 edits.json，跨重启保留） ----------
def _edits_path():
    return os.path.join(BASE, CFG.get("edits") or "edits.json")


def read_edits():
    """读取存档；缺失或损坏时返回空存档。"""
    path = _edits_path()
    if not os.path.exists(path):
        return {"v": 1, "e": {}}
    try:
        with io.open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict) and isinstance(data.get("e"), dict):
            return {"v": data.get("v", 1), "e": data["e"]}
    except Exception as exc:  # noqa: BLE001
        log("读取存档失败：%r" % (exc,))
    return {"v": 1, "e": {}}


def write_edits(data):
    """原子写入存档，返回文件路径。"""
    path = _edits_path()
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    os.replace(tmp, path)
    return path


def _data_path():
    return os.path.join(BASE, CFG.get("data") or "data.js")


def apply_edits_to_data(edits):
    """把进度条改动写回 data.js（成为「真数据」）。成功返回 True，失败返回 False。

    edits 结构：{ 分段索引(str): [10 个日期列(6..15)] }，与 viewer.html 的存档一致。
    写回后同时重算每艘船的 minDay/maxDay/count，保证分组排序和总览范围同步更新。
    """
    if not edits:
        return False
    path = _data_path()
    try:
        with io.open(path, encoding="utf-8") as fh:
            text = fh.read()
    except OSError:
        return False
    m = re.match(r"^\s*window\.GANTT\s*=\s*", text)
    if not m:
        return False
    try:
        payload = json.loads(text[m.end():].strip().rstrip(";"))
    except Exception:                                   # noqa: BLE001
        return False
    blocks = payload.get("blocks") or []
    hulls = payload.get("hulls") or []
    if not blocks:
        return False

    for k, ov in edits.items():
        try:
            i = int(k)
        except (TypeError, ValueError):
            continue
        if not (0 <= i < len(blocks)) or not isinstance(ov, list):
            continue
        b = blocks[i]
        for j in range(10):
            if j < len(ov) and ov[j]:
                b[6 + j] = ov[j]

    # 重算每艘船的 minDay / maxDay / count（与 core.build_payload 同一口径）
    for h in hulls:
        h[2] = 0
        h[3] = 0
        h[4] = 0
    for b in blocks:
        hi = b[0]
        if not (0 <= hi < len(hulls)):
            continue
        h = hulls[hi]
        lo = min([d for d in (b[6], b[8], b[12], b[14]) if d], default=0)
        hi_v = max([d for d in (b[7], b[9], b[11], b[13], b[15]) if d], default=0)
        if lo and (h[2] == 0 or lo < h[2]):
            h[2] = lo
        if hi_v > h[3]:
            h[3] = hi_v
        h[4] += 1

    payload["gen"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    data = "window.GANTT=" + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n"
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=d, suffix=".tmp")
    try:
        with io.open(fd, "w", encoding="utf-8") as fh:
            fh.write(data)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return True


class Handler(BaseHTTPRequestHandler):
    server_version = "GanttWebhook/1.0"

    # ---------- 工具 ----------
    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        raw = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Token")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            try:
                self.wfile.write(raw)
            except (BrokenPipeError, ConnectionResetError):
                pass

    def _auth(self):
        token = str(CFG.get("token") or "")
        if not token:
            return True
        given = self.headers.get("X-Token") or ""
        if not given:
            qs = self.path.split("?", 1)[1] if "?" in self.path else ""
            for part in qs.split("&"):
                if part.startswith("token="):
                    given = part[6:]
        return given == token

    def _body(self):
        length = int(self.headers.get("Content-Length") or 0)
        limit = int(CFG.get("max_mb", 32)) * 1048576
        if length > limit:
            raise ValueError("请求体过大：%d 字节（上限 %d MB）" % (length, CFG.get("max_mb")))
        return self.rfile.read(length) if length else b""

    def _file(self, name, ctype, extra=None):
        path = os.path.join(BASE, name)
        if not os.path.exists(path):
            path = paths.res(name)          # 打包后静态文件在 _MEIPASS
        if not os.path.exists(path):
            self._send(404, {"ok": False, "error": "not found: " + name})
            return
        with open(path, "rb") as fh:
            self._send(200, fh.read(), ctype, extra=extra)

    # ---------- 窗口界面 API ----------
    def _panel_auth(self):
        """/api/* 只给本机窗口用：回环地址 + X-Panel 头 + 同源 Origin。"""
        host = self.client_address[0] if self.client_address else ""
        if host not in ("127.0.0.1", "::1", "::ffff:127.0.0.1"):
            return False
        if self.headers.get("X-Panel") != "1":
            return False
        origin = self.headers.get("Origin") or ""
        if origin and not re.match(r"^https?://(127\.0\.0\.1|localhost)(:\d+)?$", origin):
            return False
        return True

    def _api(self, method):
        split = urlsplit(self.path)
        query = {k: v[-1] for k, v in parse_qs(split.query).items()}
        if not self._panel_auth():
            self._send(403, {"ok": False,
                             "error": "窗口界面接口只允许本机调用（需要 X-Panel 头与回环地址）"})
            return
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else b""
        try:
            code, payload = panel.api(method, split.path, query, body)
        except Exception as exc:              # noqa: BLE001
            log("API %s %s 出错：%r" % (method, split.path, exc))
            code, payload = 500, {"ok": False, "error": repr(exc)}
        self._send(code, payload, extra={"Cache-Control": "no-store"})

    def _api_export(self):
        """把当前横道图导出成 Excel（二进制下载）。"""
        if not self._panel_auth():
            self._send(403, {"ok": False,
                             "error": "窗口界面接口只允许本机调用（需要 X-Panel 头与回环地址）"})
            return
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else b""
        try:
            data = json.loads(body.decode("utf-8-sig")) if body else {}
            content, filename = panel.export_xlsx(data)
            os.makedirs(panel.DATE_DIR, exist_ok=True)
            with open(os.path.join(panel.DATE_DIR, filename), "wb") as fh:
                fh.write(content)            # 同时留一份在 DATE，避免下载被拦截
        except Exception as exc:              # noqa: BLE001
            log("导出 Excel 出错：%r" % (exc,))
            self._send(500, {"ok": False, "error": "%s: %s" % (type(exc).__name__, exc)})
            return
        self._send(200, content,
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   extra={"Content-Disposition":
                          "attachment; filename*=UTF-8''%s" % quote(filename),
                          "Cache-Control": "no-store"})

    def _save_edits(self):
        """保存进度条操作存档（浏览器 POST /edits），并写回 data.js。"""
        try:
            raw = self._body()
            data = json.loads(raw.decode("utf-8-sig"))
        except Exception as exc:  # noqa: BLE001
            self._send(400, {"ok": False, "error": "解析失败：%s" % exc})
            return
        if not isinstance(data, dict) or not isinstance(data.get("e"), dict):
            self._send(400, {"ok": False, "error": '格式需为 {"e": {...}}'})
            return
        try:
            with LOCK:
                edits = data["e"]
                baked = apply_edits_to_data(edits)
                if baked:
                    # 已写回真数据：清空存档，避免下次加载时重复叠加
                    write_edits({"v": 2, "e": {}})
                    path = _edits_path()
                else:
                    # data.js 不可写时兜底：仍存进 edits.json，改动不丢
                    path = write_edits({"v": data.get("v", 1), "e": edits})
            self._send(200, {"ok": True, "path": os.path.basename(path), "baked": baked})
        except Exception as exc:  # noqa: BLE001
            log("保存存档失败：%r" % (exc,))
            self._send(500, {"ok": False, "error": str(exc)})

    def _restore_original(self):
        """把 data.js 恢复成最近一次从 Excel 生成的原始数据（POST /restore）。"""
        try:
            with LOCK:
                data_path = _data_path()
                orig_path = data_path + ".orig"
                if not os.path.exists(orig_path):
                    self._send(404, {"ok": False,
                                     "error": "没有原始备份：请先从 Excel / 在线表重新生成一次 data.js"})
                    return
                shutil.copyfile(orig_path, data_path)
                write_edits({"v": 2, "e": {}})
            self._send(200, {"ok": True})
        except Exception as exc:  # noqa: BLE001
            log("恢复原始失败：%r" % (exc,))
            self._send(500, {"ok": False, "error": str(exc)})

    # ---------- 路由 ----------
    def do_OPTIONS(self):
        self._send(204, b"")

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path.startswith("/api/"):
            self._api("GET")
            return
        try:
            if path in ("/", "/index.html", "/viewer.html"):
                self._file(CFG["viewer"], "text/html; charset=utf-8")
            elif path in ("/app", "/app.html", "/panel"):
                self._file("app.html", "text/html; charset=utf-8")
            elif path == "/data.js":
                self._file(CFG["data"], "application/javascript; charset=utf-8",
                           extra={"Cache-Control": "no-store"})
            elif path == "/edits.js":
                self._send(200, "window.GANTT_EDITS = %s;" %
                           json.dumps(read_edits(), ensure_ascii=False),
                           "application/javascript; charset=utf-8",
                           extra={"Cache-Control": "no-store"})
            elif path == "/health":
                stale = None
                data_path = os.path.join(BASE, CFG["data"])
                if os.path.exists(data_path):
                    stale = int(dt.datetime.now().timestamp() - os.path.getmtime(data_path))
                self._send(200, {"ok": True, "builds": STATE["builds"],
                                 "last": STATE["last"], "data_age_sec": stale})
            elif path == "/api/status":
                self._send(200, {"ok": True, "builds": STATE["builds"],
                                 "last": STATE["last"], "error": STATE["error"]})
            else:
                self._send(404, {"ok": False, "error": "unknown path"})
        except Exception as exc:  # noqa: BLE001
            log("GET %s 出错：%r" % (path, exc))
            self._send(500, {"ok": False, "error": str(exc)})

    def do_POST(self):
        path = self.path.split("?", 1)[0]
        if path == "/api/export":
            self._api_export()
            return
        if path.startswith("/api/"):
            self._api("POST")
            return
        if path == "/edits":
            self._save_edits()
            return
        if path == "/restore":
            self._restore_original()
            return
        if path not in ("/webhook/gantt", "/webhook/csv"):
            self._send(404, {"ok": False, "error": "unknown path"})
            return
        if not self._auth():
            self._send(401, {"ok": False, "error": "bad token"})
            return

        try:
            raw = self._body()
            text = raw.decode("utf-8-sig")
        except Exception as exc:  # noqa: BLE001
            self._send(400, {"ok": False, "error": str(exc)})
            return

        try:
            if path == "/webhook/csv":
                rows = core.parse_csv(text)
                source = "webhook/csv"
            else:
                body = json.loads(text)
                rows = body if isinstance(body, list) else body.get("rows") or body.get("data")
                if not isinstance(rows, list):
                    raise ValueError('JSON 需要是数组，或 {"rows":[...]}')
                source = "webhook/json"
        except Exception as exc:  # noqa: BLE001
            STATE["error"] = str(exc)
            log("解析失败：%r" % (exc,))
            self._send(400, {"ok": False, "error": "解析失败：%s" % exc})
            return

        with LOCK:
            try:
                archive(rows, "payload")
                info = rebuild(rows, source)
                self._send(200, info)
            except Exception as exc:  # noqa: BLE001
                STATE["error"] = str(exc)
                log("构建失败：%r\n%s" % (exc, traceback.format_exc()))
                self._send(422, {"ok": False, "error": str(exc)})

    def log_message(self, fmt, *args):  # 静默默认访问日志
        pass


def create_server(host=None, port=None, config_path=None):
    """建好（还没开始 serve 的）HTTP 服务。窗口界面 window.py 也走这里。"""
    global CFG
    paths.ensure_defaults()
    CFG = load_config(config_path or os.path.join(BASE, "config.json"))
    if host:
        CFG["host"] = host
    if port:
        CFG["port"] = int(port)
    panel.configure(CFG)
    httpd = ThreadingHTTPServer((CFG["host"], int(CFG["port"])), Handler)
    httpd.daemon_threads = True
    return httpd, CFG


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.path.join(BASE, "config.json"))
    ap.add_argument("--host")
    ap.add_argument("--port", type=int)
    ap.add_argument("--selftest", action="store_true", help="用示例数据跑一遍构建，不起服务")
    args = ap.parse_args()

    if args.selftest:
        panel.configure(load_config(args.config))
        sample = os.path.join(BASE, "sample_payload.json")
        with io.open(sample, encoding="utf-8") as fh:
            body = json.load(fh)
        rows = body if isinstance(body, list) else body.get("rows")
        info = rebuild(rows, "selftest")
        print(json.dumps(info, ensure_ascii=False))
        return 0

    srv, cfg = create_server(args.host, args.port, args.config)
    log("横道图 Webhook 已启动：http://%s:%s/" % (cfg["host"], cfg["port"]))
    log("  查看器  http://127.0.0.1:%s/" % cfg["port"])
    log("  窗口界面 http://127.0.0.1:%s/app" % cfg["port"])
    log("  推送    POST http://127.0.0.1:%s/webhook/gantt" % cfg["port"])
    if not cfg.get("token"):
        log("  注意：未设置 token，任何能访问到本机的人都能推送数据")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        log("收到 Ctrl+C，退出")
    finally:
        srv.server_close()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        sys.exit(1)
