# -*- coding: utf-8 -*-
"""
build.py —— 把《数据更新中.xlsx》的「临时」表转成横道图数据（data.js）

用法：
    python build.py                    # 用默认路径
    python build.py --src "D:\\xx.xlsx"

输出（都写在程序根目录 / EXE 旁边）：
    data.js        浏览器可直接 <script src> 加载（file:// 也能用，不需要起服务器）
    build_report.txt  本次生成的数据统计

数据单位说明：
    日期一律转成「相对 2020-01-01 的天数」，JS 侧只做整数运算，避免时区/夏令时问题。
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import os
import sys
import traceback

from openpyxl import load_workbook

try:
    import paths
except Exception:                                   # noqa: BLE001
    paths = None

DEFAULT_SRC = r"D:\DSP6\数据更新中_在建.xlsx"
SHEET = "Sheet1"

if paths is not None:
    OUT_DIR = paths.BASE
else:
    OUT_DIR = os.path.dirname(os.path.abspath(__file__))

OUT_JS = os.path.join(OUT_DIR, "data.js")
OUT_REPORT = os.path.join(OUT_DIR, "build_report.txt")

EPOCH = dt.date(2020, 1, 1)

# 「临时」表的列索引（0 基），依据实际表头第 4 行
COL = {
    "hull": 4,        # 船号
    "model": 1,       # 船型
    "block": 5,       # 总段
    "seg": 8,         # 新段号
    "type": 12,       # 分段类型
    "ton": 14,        # 修正吨
    "area": 52,       # 施工区域
    "shop": 55,       # 分段车间
    "dept": 56,       # 施工部门
    "pb_a": 15,       # 拼板开工
    "pb_b": 16,       # 拼板完工
    "tt_a": 18,       # 分段上胎
    "tt_b": 19,       # 分段下胎
    "jg_a": 21,       # 结构完工
    "sn_a": 37,       # 舾装开工
    "sn_b": 41,       # 舾装销项完工
    "zx_a": 47,       # 结构成品化
    "zx_b": 50,       # 转序日期（有值 = 该分段已完工，工期在这天截断）
    "zx_b2": 51,      # 生产部转序（AZ 列，转序日期为空时兜底）
    "state": 66,      # 「船体状态」（完成 / 在建）实际在第 67 列（0 基 66）；旧映射 74 已空
}

HEADER_ROWS = 4  # 第 1 行空，2/3/4 是三层表头


def as_day(v):
    """把单元格转成天数；非法返回 0。"""
    if isinstance(v, dt.datetime):
        return (v.date() - EPOCH).days
    if isinstance(v, dt.date):
        return (v - EPOCH).days
    return 0


def clean_str(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("/", "-", "—", "None") else s


def run(src, sheet):
    """读 Excel → 写 data.js。返回统计 dict；出错抛 RuntimeError。"""
    # 控制台是 GBK 时，直接 print 中文会抛 UnicodeEncodeError，这里兜底
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

    if not os.path.exists(src):
        raise RuntimeError("找不到源文件：%s" % src)

    report = []

    def log(s=""):
        report.append(str(s))
        print(s)

    log("源文件：%s" % src)
    log("读取时间：%s" % dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    wb = load_workbook(src, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        wb.close()
        raise RuntimeError("没有工作表 %s，现有：%r" % (sheet, wb.sheetnames))
    ws = wb[sheet]

    rows = ws.iter_rows(values_only=True)
    for _ in range(HEADER_ROWS):
        next(rows, None)

    hull_index = {}
    hulls = []          # [name, model, first_day, last_day, count]
    type_index = {}
    types = []
    area_index = {}
    areas = []
    shop_index = {}
    shops = []

    blocks = []
    finished_idx = set()      # BW 列标"完成"的行
    total = 0
    skipped_no_date = 0
    phase_hits = {"pb": 0, "tt": 0, "jg": 0, "sn": 0, "zx": 0}

    for row in rows:
        if row is None:
            continue
        total += 1
        n = len(row)

        def g(key):
            i = COL[key]
            return row[i] if i < n else None

        hull = clean_str(g("hull"))
        if not hull:
            continue

        # 先登记船号：即使该船还没有任何阶段日期（未开工），也要出现在总览里，
        # 否则会漏掉整艘还没开始生产的船。
        if hull not in hull_index:
            hull_index[hull] = len(hulls)
            hulls.append([hull, clean_str(g("model")), 0, 0, 0])

        pb_a, pb_b = as_day(g("pb_a")), as_day(g("pb_b"))
        tt_a, tt_b = as_day(g("tt_a")), as_day(g("tt_b"))
        jg_a = as_day(g("jg_a"))
        sn_a, sn_b = as_day(g("sn_a")), as_day(g("sn_b"))
        zx_a = as_day(g("zx_a"))
        zx_b = as_day(g("zx_b")) or as_day(g("zx_b2"))   # 优先「转序日期」，空则用 AZ 列

        if not (pb_a or pb_b or tt_a or tt_b or jg_a or sn_a or zx_a):
            skipped_no_date += 1
            continue

        # 不补齐终点：终点为空表示"还没完工"，交给查看器画到今天
        # 结构 = 下胎(或上胎) -> 结构完工
        jg_start = tt_b or tt_a
        if jg_start and jg_a and jg_a < jg_start:
            jg_a = jg_start

        for k, a, b in (("pb", pb_a, pb_b), ("tt", tt_a, tt_b),
                        ("jg", jg_start, jg_a), ("sn", sn_a, sn_b),
                        ("zx", zx_a, zx_b)):
            if a and b:
                phase_hits[k] += 1

        def intern(val, index, pool):
            if not val:
                return -1
            if val not in index:
                index[val] = len(pool)
                pool.append(val)
            return index[val]

        type_i = intern(clean_str(g("type")), type_index, types)
        area_i = intern(clean_str(g("area")), area_index, areas)
        shop_i = intern(clean_str(g("shop")), shop_index, shops)

        ton = g("ton")
        try:
            ton = round(float(ton), 1)
        except (TypeError, ValueError):
            ton = 0

        blocks.append([
            hull_index[hull],
            clean_str(g("seg")) or clean_str(g("block")),
            type_i,
            ton,
            area_i,
            shop_i,
            pb_a, pb_b, tt_a, tt_b, jg_start, jg_a, sn_a, sn_b, zx_a, zx_b,
        ])
        if "完成" in clean_str(g("state")):          # BW 列「船体状态」
            finished_idx.add(len(blocks) - 1)

        h = hulls[hull_index[hull]]
        lo = min([d for d in (pb_a, tt_a, sn_a, zx_a) if d], default=0)
        hi = max([d for d in (pb_b, tt_b, jg_a, sn_b, zx_b) if d], default=0)
        if lo and (h[2] == 0 or lo < h[2]):
            h[2] = lo
        if hi > h[3]:
            h[3] = hi
        h[4] += 1

    wb.close()

    # ---- BW 列标"完成"但没有转序日期的分段：按顺序向后借用下一个分段的转序日期 ----
    by_ship = {}
    for bi, b in enumerate(blocks):
        by_ship.setdefault(b[0], []).append(bi)

    filled = 0
    for hi, idxs in by_ship.items():
        idxs.sort(key=lambda i: (blocks[i][8] or blocks[i][14] or 10 ** 9, str(blocks[i][1])))
        ship_max = max([blocks[i][15] for i in idxs if blocks[i][15]] or [0])
        nxt = 0
        for i in reversed(idxs):
            b = blocks[i]
            if b[15]:                      # 自己有转序日期
                nxt = b[15]
            elif i in finished_idx and (b[8] or b[14]):   # 完成、有开工日期、但缺转序 → 借用后一道的
                use = nxt or ship_max
                if use and use >= (b[8] or 0):
                    b[15] = use
                    nxt = use
                    filled += 1

    today = (dt.date.today() - EPOCH).days
    payload = {
        "gen": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "today": today,
        "epoch": EPOCH.isoformat(),
        "hulls": hulls,
        "types": types,
        "areas": areas,
        "shops": shops,
        "blocks": blocks,
    }

    js = "window.GANTT=" + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n"
    with io.open(OUT_JS, "w", encoding="utf-8") as fh:
        fh.write(js)

    # 同时写一份「原始备份」，供进度条改动写回 data.js 后仍能「恢复原始」
    try:
        with io.open(OUT_JS + ".orig", "w", encoding="utf-8") as fh:
            fh.write(js)
    except OSError:
        pass

    if paths is not None:
        try:
            paths.reset_edits_archive()
        except Exception:           # noqa: BLE001
            pass

    size_mb = os.path.getsize(OUT_JS) / 1048576.0

    log("")
    log("===== 统计 =====")
    log("扫描行数           : %d" % total)
    log("无任何日期跳过     : %d" % skipped_no_date)
    log("生成分段行         : %d" % len(blocks))
    log("船数               : %d" % len(hulls))
    log("分段类型           : %d" % len(types))
    log("施工区域           : %d" % len(areas))
    log("分段车间           : %d" % len(shops))
    log("各阶段有完整起止行 : %s" % phase_hits)
    log("输出文件           : %s (%.2f MB)" % (OUT_JS, size_mb))
    log("")
    log("船号 Top10（按分段数）：")
    for h in sorted(hulls, key=lambda x: -x[4])[:10]:
        log("   %-30s %4d 段   %s" % (h[0], h[4], h[1]))

    with io.open(OUT_REPORT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(report))

    return {
        "src": src, "sheet": sheet,
        "scanned": total, "skipped": skipped_no_date,
        "blocks": len(blocks), "hulls": len(hulls),
        "types": len(types), "areas": len(areas), "shops": len(shops),
        "phase_hits": phase_hits, "filled": filled,
        "out": OUT_JS, "size": os.path.getsize(OUT_JS),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--sheet", default=SHEET)
    args = ap.parse_args()
    try:
        run(args.src, args.sheet)
    except RuntimeError as exc:
        print(exc)
        return 2
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        sys.exit(1)
