# -*- coding: utf-8 -*-
"""paths.py —— 统一路径（源码运行 / PyInstaller 打包后都能用）

打包成 EXE 后：
  · BASE = EXE 所在目录（可写：data.js / gantt_config.json / logs / DATE）
  · RES  = PyInstaller 解包目录 _MEIPASS（只读：app.html / viewer.html / TM.ico / 默认配置）
"""
from __future__ import annotations

import json
import os
import shutil
import sys

FROZEN = bool(getattr(sys, "frozen", False))
BASE = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))
RES = getattr(sys, "_MEIPASS", BASE)


def res(*parts):
    """只读资源：优先 _MEIPASS，找不到再退回 BASE。"""
    p = os.path.join(RES, *parts)
    if os.path.exists(p):
        return p
    return os.path.join(BASE, *parts)


def out(*parts):
    """可写文件：始终在 BASE。"""
    return os.path.join(BASE, *parts)


def load_env():
    """从注册表补齐缺失的环境变量（Windows）。

    用 setx 写用户环境变量后，已经在跑的资源管理器 / 已打开的窗口不会刷新环境块，
    从桌面双击 EXE 拿到的还是旧环境（看不到刚设的 WPS_AIRSCRIPT_TOKEN）。
    这里在启动时把注册表里的用户 / 系统变量补进 os.environ（只补缺失的，不覆盖）。
    """
    if os.name != "nt":
        return
    try:
        import winreg
    except ImportError:
        return
    keys = (
        (winreg.HKEY_CURRENT_USER, r"Environment"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
    )
    for hive, path in keys:
        try:
            with winreg.OpenKey(hive, path) as h:
                count = winreg.QueryInfoKey(h)[1]
                for i in range(count):
                    try:
                        name, value, _kind = winreg.EnumValue(h, i)
                    except OSError:
                        continue
                    if not name or not isinstance(value, str):
                        continue
                    if os.environ.get(name):
                        continue               # 已有值（哪怕是设的）一律不动
                    try:
                        os.environ[name] = os.path.expandvars(value)
                    except (ValueError, TypeError):
                        os.environ[name] = value
        except OSError:
            continue


load_env()


def ensure_defaults(names=("config.json", "gantt_config.json", "data.js")):
    """首次运行：把打包时带的默认文件复制到 BASE（已存在就不动）。"""
    for name in names:
        dst = out(name)
        if os.path.exists(dst):
            continue
        src = res(name)
        if os.path.exists(src) and os.path.abspath(src) != os.path.abspath(dst):
            try:
                shutil.copy2(src, dst)
            except OSError:
                pass


def load_json(path, default=None):
    """读一个 JSON 文件；缺失或损坏返回 default。"""
    try:
        with open(path, encoding="utf-8-sig") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def edits_reset_enabled():
    """出图重建 data.js 时是否清空进度条存档（gantt_config.json，缺省 True）。"""
    cfg = load_json(out("gantt_config.json"), {}) or {}
    return bool(cfg.get("reset_edits_on_rebuild", True))


def reset_edits_archive():
    """出图重建 data.js 后，按开关清空进度条操作存档；返回是否清过。"""
    if not edits_reset_enabled():
        return False
    svc = load_json(out("config.json"), {}) or {}
    name = svc.get("edits") or "edits.json"
    path = out(name)
    try:
        if os.path.exists(path):
            os.remove(path)
        return True
    except OSError:
        return False
