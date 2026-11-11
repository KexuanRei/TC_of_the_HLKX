# -*- coding: utf-8 -*-
"""
vdesktop.py —— 把下载专用 Edge 放到另一个 Windows 虚拟桌面

用「切换桌面 → 启动 Edge → 切回来」的方式，让下载过程发生在另一个虚拟桌面上，
当前桌面不受打扰。只用公开/稳定能力：
  · 注册表读当前虚拟桌面数量（HKCU ... Explorer\\VirtualDesktops\\VirtualDesktopIDs）
  · user32 keybd_event 发送 Win+Ctrl+D / Win+Ctrl+Right / Win+Ctrl+Left

注意：切换桌面是瞬间动作（约 0.6 秒），真正耗时的下载过程发生在另一个桌面。
"""
from __future__ import annotations

import ctypes
import time
import winreg

VK_LWIN = 0x5B
VK_LCONTROL = 0xA2
VK_RIGHT = 0x27
VK_LEFT = 0x25
VK_D = 0x44
KEYEVENTF_KEYUP = 0x0002

_user32 = ctypes.windll.user32


def _key(vk: int, up: bool = False) -> None:
    _user32.keybd_event(vk, 0, KEYEVENTF_KEYUP if up else 0, 0)


def _combo(keys) -> None:
    for vk in keys:
        _key(vk, False)
    for vk in reversed(keys):
        _key(vk, True)


def desktop_count() -> int:
    """当前有多少个虚拟桌面（0 或读不到时按 1 处理）。"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops",
        )
        try:
            data, _ = winreg.QueryValueEx(key, "VirtualDesktopIDs")
        finally:
            winreg.CloseKey(key)
        if isinstance(data, bytes) and data:
            return max(1, len(data) // 16)
    except OSError:
        pass
    return 1


def switch_away() -> None:
    """切到「另一个」虚拟桌面；若只有一个桌面则先新建一个（Win+Ctrl+D）。"""
    if desktop_count() <= 1:
        _combo([VK_LWIN, VK_LCONTROL, VK_D])      # 新建虚拟桌面并切过去
    else:
        _combo([VK_LWIN, VK_LCONTROL, VK_RIGHT])  # 切到右侧相邻桌面
    time.sleep(0.6)


def switch_back() -> None:
    """切回上一个虚拟桌面（与 switch_away 对称）。"""
    _combo([VK_LWIN, VK_LCONTROL, VK_LEFT])
    time.sleep(0.6)
