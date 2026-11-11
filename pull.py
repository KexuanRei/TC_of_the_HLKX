# -*- coding: utf-8 -*-
"""
pull.py —— 从线上 URL 拉数据，直接生成 data.js（推荐：本机跑，不依赖 AirScript）

用法：
    python pull.py --url "https://xxx/webhook/abc"
    python pull.py --url "https://xxx" --header "Authorization: Bearer xxx"
    python pull.py --url "https://xxx" --interval 300      # 每 5 分钟拉一次
    python pull.py --url "https://xxx" --dry-run           # 只拉不写，看看返回什么

数据源支持：
    JSON 数组                [{...}, {...}]
    JSON 包一层              {"rows":[...]} / {"data":[...]}
    CSV / TSV 文本           第一行是表头

字段名支持中文表头（船号/新段号/上胎/下胎/转序日期…），也支持英文短名，见 core.py 的 ALIASES。
"""
from __future__ import annotations

import argparse
import io
import json
import sys
import time
import traceback
import urllib.error
import urllib.request

import core
import net


def fetch(url, headers, timeout=30):
    raw = net.request(url, headers=headers or {"User-Agent": "gantt-pull/1.0"},
                      timeout=timeout, retries=2)
    ctype = ""
    text = raw.decode("utf-8-sig", errors="replace")
    return text, ctype


def to_rows(text, ctype):
    txt = text.strip()
    if not txt:
        raise ValueError("返回内容为空")
    if txt[0] in "[{":
        body = json.loads(txt)
        if isinstance(body, list):
            return body
        for key in ("rows", "data", "items", "list", "records"):
            if isinstance(body.get(key), list):
                return body[key]
        raise ValueError("JSON 里找不到数组（试过 rows/data/items/list/records）")
    return core.parse_csv(txt)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True, help="数据源地址（webhook / API）")
    ap.add_argument("--header", action="append", default=[], help='额外请求头，如 "Authorization: Bearer x"')
    ap.add_argument("--out", default="data.js")
    ap.add_argument("--interval", type=int, default=0, help=">0 时循环拉取，单位秒")
    ap.add_argument("--dry-run", action="store_true", help="只打印返回摘要，不生成文件")
    args = ap.parse_args()

    headers = {}
    for h in args.header:
        if ":" in h:
            k, v = h.split(":", 1)
            headers[k.strip()] = v.strip()

    def once():
        text, ctype = fetch(args.url, headers)
        rows = to_rows(text, ctype)
        print("拉到 %d 行，Content-Type=%s" % (len(rows), ctype or "-"))
        if args.dry_run:
            for r in rows[:3]:
                print("  样例:", json.dumps(r, ensure_ascii=False)[:300])
            return
        payload = core.build_payload(rows, "pull:%s" % args.url[:60])
        if not payload["blocks"]:
            raise ValueError("没有可用数据：缺船号或缺日期")
        size = core.write_js(payload, args.out)
        print("已生成 %s：%d 段 / %d 船 / %.2f MB  (%s)"
              % (args.out, len(payload["blocks"]), len(payload["hulls"]), size / 1048576.0, payload["gen"]))

    if args.interval <= 0:
        once()
        return 0

    while True:
        try:
            once()
        except urllib.error.HTTPError as exc:
            print("HTTP %s：%s" % (exc.code, exc.reason))
        except Exception as exc:  # noqa: BLE001
            print("拉取失败：%r" % (exc,))
            traceback.print_exc()
        time.sleep(args.interval)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("已停止")
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        sys.exit(1)
