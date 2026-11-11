# -*- coding: utf-8 -*-
"""net.py —— 统一 HTTP 请求（绕开失效的系统代理，最后兜底 curl.exe）

背景：Windows 上 urllib 会读注册表里的 WinINET 代理（Internet 选项 → 局域网设置）。
代理软件（如 127.0.0.1:21081 这类本地客户端）退出后这个设置往往还留着，
于是 urllib 去连那个死端口，报 URLError / ConnectionAbortedError(10053)，
而浏览器和 curl 走直连一切正常——看起来像"网站挂了"，其实是代理设置残留。

策略：直连 → 系统/环境变量代理 → curl.exe（Windows 自带）。
4xx/5xx 直接抛 HTTPError，交给调用方判断（比如令牌失效）。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
import urllib.error
import urllib.request

DEFAULT_TIMEOUT = 60
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 gantt/1.1"


def _openers():
    """直连放最前面：大多数情况下系统代理是残留设置，直连才是对的。"""
    yield urllib.request.build_opener(urllib.request.ProxyHandler({}))
    yield urllib.request.build_opener()


def _curl_exe():
    exe = shutil.which("curl")
    if exe:
        return exe
    guess = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "curl.exe")
    return guess if os.path.exists(guess) else None


def _curl_request(url, data, headers, timeout):
    exe = _curl_exe()
    if not exe:
        raise OSError("找不到 curl.exe")
    tmp = []
    try:
        fd, body_path = tempfile.mkstemp(prefix="net-body-", suffix=".bin")
        os.close(fd)
        fd, out_path = tempfile.mkstemp(prefix="net-out-", suffix=".bin")
        os.close(fd)
        tmp = [body_path, out_path]
        cmd = [exe, "-sS", "--max-time", str(int(timeout)), "-o", out_path, "-w", "%{http_code}"]
        for key, val in (headers or {}).items():
            cmd += ["-H", "%s: %s" % (key, val)]
        if data is not None:
            with open(body_path, "wb") as fh:
                fh.write(data)
            cmd += ["--data-binary", "@" + body_path]
        cmd.append(url)
        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000        # CREATE_NO_WINDOW：不闪黑框
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=timeout + 20, **kwargs)
        code = int((proc.stdout or b"0").decode("ascii", "replace").strip() or "0")
        with open(out_path, "rb") as fh:
            raw = fh.read()
        if code >= 400:
            detail = (proc.stderr or b"").decode("utf-8", "replace").strip()[:200]
            raise urllib.error.HTTPError(url, code, detail or "HTTP %d" % code, {}, None)
        if code == 0:
            detail = (proc.stderr or b"").decode("utf-8", "replace").strip()[:300]
            raise OSError("curl 请求失败：%s" % (detail or "无输出"))
        return raw
    finally:
        for path in tmp:
            try:
                os.remove(path)
            except OSError:
                pass


def request(url, data=None, headers=None, timeout=DEFAULT_TIMEOUT, retries=3):
    """返回响应体 bytes；网络失败会依次尝试直连 / 系统代理 / curl。"""
    hdr = {"User-Agent": UA}
    hdr.update(headers or {})
    method = "POST" if data is not None else "GET"
    last = None
    for attempt in range(max(1, int(retries))):
        for opener in _openers():
            try:
                req = urllib.request.Request(url, data=data, headers=hdr, method=method)
                with opener.open(req, timeout=timeout) as resp:
                    return resp.read()
            except urllib.error.HTTPError:
                raise
            except Exception as exc:                    # noqa: BLE001
                last = exc
        try:
            return _curl_request(url, data, hdr, timeout)
        except urllib.error.HTTPError:
            raise
        except Exception as exc:                        # noqa: BLE001
            last = exc
        if attempt < int(retries) - 1:
            time.sleep(2 * (attempt + 1))
    raise urllib.error.URLError(last if last is not None else "网络请求失败")
