# -*- coding: utf-8 -*-
"""
gantt_sync.py —— 从 WPS/kdocs AirScript webhook 拉数据 → 生成横道图 data.js

协议与 wps_sync 一致：POST {"Context":{"argv":{...}}} + 头 AirScript-Token
    action=meta  → {sheets:[{name,rows,cols}]}
    action=data  → {rows:[[...]], next_row}

用法：
    python gantt_sync.py --check            # 列出工作表
    python gantt_sync.py --peek             # 看前 3 行数据，确认列位置
    python gantt_sync.py                    # 拉全量并生成 data.js
    python gantt_sync.py --sheet 临时 --webhook https://... --token xxx
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request

import core
import net
import paths
from build import COL, HEADER_ROWS      # 列映射与表头行数，和 xlsx 路径保持同一套

BASE = paths.BASE
CONFIG = paths.out("gantt_config.json")
LOG = paths.out("gantt.log")
ENV_RE = re.compile(r"\$\{ENV:([A-Za-z_][A-Za-z0-9_]*)\}")

# build.py 的列名 → core.py 认识的规范名
KEY_MAP = {
    "hull": "hull", "model": "model", "seg": "seg", "type": "type", "ton": "ton",
    "area": "area", "shop": "shop", "dept": "dept", "state": "state",
    "pb_a": "pba", "pb_b": "pbb", "tt_a": "tta", "tt_b": "ttb",
    "jg_a": "jga", "sn_a": "sna", "sn_b": "snb", "zx_a": "zxa", "zx_b": "zxb",
}


def log(msg):
    line = "%s  %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line)
    try:
        with io.open(LOG, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass


def load_config(path, override):
    cfg = {}
    if os.path.exists(path):
        with io.open(path, encoding="utf-8-sig") as fh:
            cfg = json.load(fh)

    def expand(v):
        if isinstance(v, str):
            return ENV_RE.sub(lambda m: os.environ.get(m.group(1), ""), v)
        if isinstance(v, dict):
            return {k: expand(x) for k, x in v.items()}
        if isinstance(v, list):
            return [expand(x) for x in v]
        return v

    cfg = expand(cfg)
    cfg.update({k: v for k, v in (override or {}).items() if v})
    cfg.setdefault("sheet", "")
    cfg.setdefault("rows_per_call", 500)
    cfg.setdefault("max_rows", 200000)
    cfg.setdefault("timeout", 120)
    cfg.setdefault("header_rows", HEADER_ROWS)
    cfg.setdefault("out", "data.js")
    if not cfg.get("webhook"):
        raise SystemExit("没有 webhook：请在 gantt_config.json 里填，或用 --webhook 传")
    if not cfg.get("token"):
        raise SystemExit("没有 token：在窗口「设置」里粘贴 AirScript 令牌，"
                         "或设置环境变量 WPS_AIRSCRIPT_TOKEN，或菜单 9 写入")
    return cfg


def call(cfg, argv, retries=3):
    body = json.dumps({"Context": {"argv": argv}}, ensure_ascii=False).encode("utf-8")
    last = None
    # net.request 内部会依次试「直连 → 系统代理 → curl.exe」，这里只管重试和错误翻译，
    # 免得刷新一下就报 10053（系统里残留的 127.0.0.1 代理设置导致的）。
    for i in range(retries):
        try:
            raw = net.request(cfg["webhook"], data=body,
                              headers={"Content-Type": "application/json; charset=utf-8",
                                       "AirScript-Token": cfg["token"]},
                              timeout=cfg["timeout"], retries=1)
            payload = json.loads(raw.decode("utf-8", "replace"))
            break
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                raise SystemExit("HTTP %d：令牌无效或过期，或与文件不属于同一账号" % exc.code)
            last = exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last = exc
        if i < retries - 1:
            time.sleep(5 * (i + 1))
    else:
        reason = getattr(last, "reason", None) or last
        raise SystemExit("调用失败（网络不通）：%r" % (reason,))

    if payload.get("error"):
        raise SystemExit("脚本报错：%s" % payload["error"])
    result = (payload.get("data") or {}).get("result")
    # 脚本可能返回对象，也可能被运行时序列化成 JSON 字符串（甚至二次编码），这里逐层解开
    for _ in range(3):
        if not isinstance(result, str):
            break
        text = result.strip()
        if text in ("", "null", "undefined", "[Undefined]"):
            return None
        try:
            result = json.loads(text)
        except ValueError:
            break
    return result


def list_sheets(cfg):
    meta = call(cfg, {"action": "meta"})
    if not isinstance(meta, dict) or "sheets" not in meta:
        raise SystemExit("meta 返回异常：%s" % str(meta)[:200])
    return meta["sheets"]


def pick_sheet(cfg, sheets):
    want = (cfg.get("sheet") or "").strip()
    if want:
        for s in sheets:
            if s["name"] == want:
                return s
        names = ", ".join(s["name"] for s in sheets)
        raise SystemExit("配置里的表 %r 不存在，现有：%s" % (want, names))
    return sheets[0]


def fetch_rows(cfg, sheet):
    total = min(int(sheet["rows"]), int(cfg["max_rows"]))
    cols = int(sheet["cols"])
    step = max(10, int(cfg["rows_per_call"]))
    out = []
    start = 1
    while start <= total:
        n = min(step, total - start + 1)
        res = call(cfg, {"action": "data", "sheet": sheet["name"],
                         "row_start": start, "row_count": n})
        chunk = (res.get("rows") if isinstance(res, dict) else res) or []
        if not chunk:
            break
        out.extend(chunk)
        start += len(chunk)
        print("\r  已拉取 %d / %d 行" % (len(out), total), end="", flush=True)
    print()
    return out, cols


def to_dicts(rows, cols):
    """按列索引转成 core 认识的字典；表头行跳过。"""
    out = []
    for r in rows[int(HEADER_ROWS):]:
        if not r:
            continue
        d = {}
        for name, idx in COL.items():
            d[KEY_MAP.get(name, name)] = r[idx] if idx < len(r) else None
        # 转序日期为空时用 AZ「生产部转序」兜底
        if not d.get("zxb") and len(r) > 51:
            d["zxb"] = r[51]
        if not d.get("hull"):
            continue
        out.append(d)
    return out


def build(cfg, sheets=None):
    sheets = sheets or list_sheets(cfg)
    sheet = pick_sheet(cfg, sheets)
    log("表 %s：%d 行 × %d 列" % (sheet["name"], sheet["rows"], sheet["cols"]))

    rows, cols = fetch_rows(cfg, sheet)
    dicts = to_dicts(rows, cols)
    log("有效数据行 %d" % len(dicts))

    payload = core.build_payload(dicts, source="wps:%s" % sheet["name"])
    if not payload["blocks"]:
        raise SystemExit("没有生成任何分段：检查列映射或日期列是否为空")
    size = core.write_js(payload, os.path.join(BASE, cfg["out"]))
    log("已生成 %s：%d 段 / %d 船 / %.2f MB"
        % (cfg["out"], len(payload["blocks"]), len(payload["hulls"]), size / 1048576.0))
    return payload


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="列出工作表")
    ap.add_argument("--peek", action="store_true", help="看前 3 行数据")
    ap.add_argument("--sheet")
    ap.add_argument("--webhook")
    ap.add_argument("--token")
    ap.add_argument("--open", action="store_true", help="生成后打开查看器")
    ap.add_argument("--config", default=CONFIG)
    args = ap.parse_args()

    cfg = load_config(args.config, {"sheet": args.sheet, "webhook": args.webhook, "token": args.token})
    if args.check:
        for s in list_sheets(cfg):
            print("  %-14s %6d 行 × %3d 列" % (s["name"], s["rows"], s["cols"]))
        return 0

    if args.peek:
        sheets = list_sheets(cfg)
        sheet = pick_sheet(cfg, sheets)
        res = call(cfg, {"action": "data", "sheet": sheet["name"], "row_start": 1, "row_count": 6})
        rows = (res.get("rows") if isinstance(res, dict) else res) or []
        for i, r in enumerate(rows):
            print("第%d行:" % (i + 1), [("" if v is None else v) for v in r[:12]])
        return 0

    build(cfg)
    if args.open:
        import os as _os
        _os.startfile(os.path.join(BASE, "viewer.html"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
