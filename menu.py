# -*- coding: utf-8 -*-
"""
menu.py —— 横道图工具（中文菜单）

双击 横道图.bat 进入；所有功能都在这里，直接调用 Python 模块，不经过其他 bat。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

CONFIG = os.path.join(BASE, "gantt_config.json")
PY = sys.executable


def clear():
    os.system("cls" if os.name == "nt" else "clear")


def pause():
    input("\n按回车返回菜单…")


def read_cfg():
    if not os.path.exists(CONFIG):
        return {}
    with open(CONFIG, encoding="utf-8-sig") as fh:
        return json.load(fh)


def save_cfg(cfg):
    with open(CONFIG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)


def cmd_sync():
    """主流程：粘贴一个 webhook → 拉数 → 出图。"""
    cfg = read_cfg()
    cur = cfg.get("webhook", "")
    print("当前 webhook：" + (cur[:72] + "…" if len(cur) > 72 else cur or "（未设置）"))
    url = input("粘贴 webhook（直接回车 = 沿用当前）：").strip() or cur
    if not url:
        print("没有可用的 webhook")
        return

    print("当前表名：" + (cfg.get("sheet") or "（自动取第一张）"))
    sheet = input("表名（回车 = 不改）：").strip()

    cfg["webhook"] = url
    if sheet:
        cfg["sheet"] = sheet
    save_cfg(cfg)
    print("已保存到 gantt_config.json\n")

    import gantt_sync
    conf = gantt_sync.load_config(gantt_sync.CONFIG, {})
    gantt_sync.build(conf)

    if input("\n现在打开横道图？[Y/n] ").strip().lower() in ("", "y", "yes"):
        cmd_open()


def cmd_sheets():
    import gantt_sync
    cfg = gantt_sync.load_config(gantt_sync.CONFIG, {})
    print("工作表：")
    for s in gantt_sync.list_sheets(cfg):
        print("  %-16s %6d 行 × %3d 列" % (s["name"], s["rows"], s["cols"]))


def cmd_peek():
    import gantt_sync
    cfg = gantt_sync.load_config(gantt_sync.CONFIG, {})
    sheets = gantt_sync.list_sheets(cfg)
    sheet = gantt_sync.pick_sheet(cfg, sheets)
    res = gantt_sync.call(cfg, {"action": "data", "sheet": sheet["name"],
                                "row_start": 1, "row_count": 6})
    rows = (res.get("rows") if isinstance(res, dict) else res) or []
    for i, r in enumerate(rows):
        print("第%d行: %s" % (i + 1, [("" if v is None else v) for v in r[:14]]))


def cmd_offline():
    subprocess.run([PY, os.path.join(BASE, "build.py")], cwd=BASE)


def cmd_open():
    path = os.path.join(BASE, "viewer.html")
    if os.name == "nt":
        os.startfile(path)  # noqa: S606
    else:
        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", path])


def cmd_server():
    print("启动 Webhook 服务（Ctrl+C 停止）…")
    subprocess.run([PY, os.path.join(BASE, "server.py")], cwd=BASE)


def cmd_schedule():
    minutes = input("每多少分钟跑一次？[默认 120] ").strip() or "120"
    task = "Gantt_Sync"
    cmd = '"%s" "%s"' % (PY, os.path.join(BASE, "gantt_sync.py"))
    subprocess.run(["schtasks", "/Create", "/F", "/SC", "MINUTE", "/MO", minutes,
                    "/TN", task, "/TR", cmd], check=False)
    print("已创建计划任务 %s，每 %s 分钟" % (task, minutes))


def cmd_toggle(enable):
    subprocess.run(["schtasks", "/Change", "/TN", "Gantt_Sync",
                    "/ENABLE" if enable else "/DISABLE"], check=False)
    print("已" + ("恢复" if enable else "暂停") + "定时任务")


def cmd_token():
    tok = input("粘贴 AirScript 令牌（输入时不回显请用菜单外的 setx）：").strip()
    if not tok:
        return
    os.environ["WPS_AIRSCRIPT_TOKEN"] = tok
    subprocess.run(["setx", "WPS_AIRSCRIPT_TOKEN", tok], check=False,
                   stdout=subprocess.DEVNULL)
    print("已写入用户环境变量 WPS_AIRSCRIPT_TOKEN（新开的窗口才生效）")


def cmd_webhook():
    url = input("粘贴 webhook 地址：").strip()
    if not url:
        return
    cfg = read_cfg()
    cfg["webhook"] = url
    with open(CONFIG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    print("已写入 gantt_config.json")


def cmd_log():
    path = os.path.join(BASE, "gantt.log")
    if not os.path.exists(path):
        print("还没有日志")
        return
    with open(path, encoding="utf-8", errors="replace") as fh:
        lines = fh.readlines()[-30:]
    print("".join(lines))


def cmd_astoken():
    tok = os.environ.get("WPS_AIRSCRIPT_TOKEN")
    print("令牌：" + ("已设置（%d 位）" % len(tok) if tok else "未设置"))


def cmd_window():
    """窗口化界面：原生窗口 + 控制台（菜单 1~13 都能在窗口里点）。"""
    subprocess.Popen([PY, os.path.join(BASE, "window.py")], cwd=BASE)
    print("已打开窗口界面（关窗口即退出）。")


MENU = [
    ("1", "粘贴 webhook → 生成横道图", cmd_sync),
    ("2", "查看工作表（确认表名）", cmd_sheets),
    ("3", "查看前几行（确认列位置）", cmd_peek),
    ("4", "用本地 Excel 生成（离线）", cmd_offline),
    ("5", "打开横道图", cmd_open),
    ("6", "启动 Webhook 服务（常驻）", cmd_server),
    ("7", "安装/重建定时任务", cmd_schedule),
    ("8", "暂停定时任务", lambda: cmd_toggle(False)),
    ("9", "恢复定时任务", lambda: cmd_toggle(True)),
    ("10", "设置 AirScript 令牌", cmd_token),
    ("11", "设置 webhook", cmd_webhook),
    ("12", "查看日志", cmd_log),
    ("13", "检查令牌状态", cmd_astoken),
    ("14", "打开窗口界面（推荐日常用）", cmd_window),
]


def main():
    while True:
        clear()
        cfg = read_cfg()
        print("=" * 62)
        print("  分段生产横道图工具")
        print("=" * 62)
        wh = cfg.get("webhook", "")
        print("  webhook : " + (wh[:58] + "…" if len(wh) > 58 else wh or "（未设置）"))
        print("  数据表   : " + (cfg.get("sheet") or "（自动取第一张）"))
        print("-" * 62)
        for key, label, _ in MENU:
            print("  %-3s %s" % (key, label))
        print("  0   退出")
        print("-" * 62)

        choice = input("请选择: ").strip()
        if choice == "0":
            return
        fn = next((f for k, _, f in MENU if k == choice), None)
        if not fn:
            continue
        print()
        try:
            fn()
        except SystemExit as exc:
            print("错误：%s" % exc)
        except Exception as exc:  # noqa: BLE001
            print("出错：%r" % (exc,))
        pause()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
