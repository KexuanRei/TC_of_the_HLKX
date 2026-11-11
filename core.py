# -*- coding: utf-8 -*-
"""
core.py —— 数据归一化 + 生成 data.js

Webhook 和 Excel 两条路共用同一套结构：
    rows(字典列表) -> blocks/hulls -> window.GANTT = {...}

字段可以用中文表头，也可以用英文短名，见 ALIASES。
日期支持：datetime / date / Excel 序列号 / "2026-10-20" / "2026/10/20" / "2026年10月20日"。
"""
from __future__ import annotations

import datetime as dt
import io
import json
import os
import tempfile

try:
    import paths
except Exception:           # noqa: BLE001
    paths = None

EPOCH = dt.date(2020, 1, 1)
DONE_WORDS = ("完成", "完工", "已完工", "已交", "已转")

ALIASES = {
    "hull": ["hull", "船号", "船体号", "ship", "hullno", "船名"],
    "model": ["model", "船型", "船型名"],
    "seg": ["seg", "新段号", "段号", "分段", "block", "blockno", "分段号"],
    "type": ["type", "分段类型", "段型"],
    "ton": ["ton", "修正吨", "吨位", "weight"],
    "area": ["area", "施工区域", "区域"],
    "shop": ["shop", "分段车间", "车间"],
    "dept": ["dept", "施工部门", "部门"],
    "pba": ["pba", "拼板开工"],
    "pbb": ["pbb", "拼板完工"],
    "tta": ["tta", "上胎", "分段上胎"],
    "ttb": ["ttb", "下胎", "分段下胎"],
    "jga": ["jga", "结构完工"],
    "sna": ["sna", "舾装开工"],
    "snb": ["snb", "舾装完工", "销项完工", "舾装销项"],
    "zxa": ["zxa", "成品化", "结构成品化"],
    "zxb": ["zxb", "转序日期", "转序", "生产部转序"],
    "state": ["state", "船体状态", "状态"],
}

_LOOKUP = {}
for _canon, _names in ALIASES.items():
    for _n in _names:
        _LOOKUP[_n.strip().lower()] = _canon


def pick(row, canon):
    """按别名取值（大小写、空格不敏感）。"""
    for k, v in row.items():
        if _LOOKUP.get(str(k).strip().lower()) == canon:
            return v
    return None


def parse_day(v) -> int:
    """任何日期表示 → 相对 2020-01-01 的天数；无法识别返回 0。"""
    if v is None:
        return 0
    if isinstance(v, dt.datetime):
        return (v.date() - EPOCH).days
    if isinstance(v, dt.date):
        return (v - EPOCH).days
    if isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        # Excel 序列号（1900 系统）
        if v and 20000 < v < 80000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=int(v)) - EPOCH).days
        return 0
    s = str(v).strip()
    if not s or s in ("/", "-", "—", "无", "None", "nan"):
        return 0
    s = (s.replace("年", "-").replace("月", "-").replace("日", "")
          .replace("/", "-").replace(".", "-").strip(" -"))
    head = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%y-%m-%d", "%Y-%m", "%Y%m%d"):
        try:
            return (dt.datetime.strptime(head, fmt).date() - EPOCH).days
        except ValueError:
            continue
    return 0


def _clean(v):
    s = str(v or "").strip()
    return "" if s in ("/", "-", "—", "None", "nan") else s


def _intern(val, index, pool):
    if not val:
        return -1
    if val not in index:
        index[val] = len(pool)
        pool.append(val)
    return index[val]


def build_payload(rows, source="webhook"):
    """rows: 字典列表 → 可直接写进 data.js 的 payload。"""
    blocks = []
    hulls = []
    hull_index = {}
    types, type_index = [], {}
    areas, area_index = [], {}
    shops, shop_index = [], {}
    finished = set()

    for r in rows:
        if not isinstance(r, dict):
            continue
        hull = str(pick(r, "hull") or "").strip()
        if not hull:
            continue

        # 先登记船号：即使该船还没有任何阶段日期（未开工 / 无工期），也要出现在总览里
        if hull not in hull_index:
            hull_index[hull] = len(hulls)
            hulls.append([hull, str(pick(r, "model") or ""), 0, 0, 0])

        pba, pbb = parse_day(pick(r, "pba")), parse_day(pick(r, "pbb"))
        tta, ttb = parse_day(pick(r, "tta")), parse_day(pick(r, "ttb"))
        jga = parse_day(pick(r, "jga"))
        sna, snb = parse_day(pick(r, "sna")), parse_day(pick(r, "snb"))
        zxa, zxb = parse_day(pick(r, "zxa")), parse_day(pick(r, "zxb"))

        if not any((pba, pbb, tta, ttb, jga, sna, snb, zxa, zxb)):
            continue

        jg_start = ttb or tta
        if jg_start and jga and jga < jg_start:
            jga = jg_start

        ton = pick(r, "ton")
        try:
            ton = round(float(ton), 1)
        except (TypeError, ValueError):
            ton = 0

        blocks.append([
            hull_index[hull],
            str(pick(r, "seg") or "").strip(),
            _intern(str(pick(r, "type") or "").strip(), type_index, types),
            ton,
            _intern(_clean(pick(r, "area")), area_index, areas),
            _intern(_clean(pick(r, "shop")), shop_index, shops),
            pba, pbb, tta, ttb, jg_start, jga, sna, snb, zxa, zxb,
        ])

        state = str(pick(r, "state") or "")
        if any(w in state for w in DONE_WORDS):
            finished.add(len(blocks) - 1)

        h = hulls[hull_index[hull]]
        lo = min([d for d in (pba, tta, sna, zxa) if d], default=0)
        hi = max([d for d in (pbb, ttb, jga, snb, zxb) if d], default=0)
        if lo and (h[2] == 0 or lo < h[2]):
            h[2] = lo
        if hi > h[3]:
            h[3] = hi
        h[4] += 1

    # “完成”但缺转序日期的分段：按工序顺序借用下一道的转序日期；整船都没有就用该船最晚转序日
    by_ship = {}
    for bi, b in enumerate(blocks):
        by_ship.setdefault(b[0], []).append(bi)

    for _hi, idxs in by_ship.items():
        idxs.sort(key=lambda i: (blocks[i][8] or blocks[i][14] or 10 ** 9, str(blocks[i][1])))
        ship_max = max([blocks[i][15] for i in idxs if blocks[i][15]] or [0])
        nxt = 0
        for i in reversed(idxs):
            b = blocks[i]
            if b[15]:
                nxt = b[15]
            elif i in finished and (b[8] or b[14]):
                use = nxt or ship_max
                if use and use >= (b[8] or 0):
                    b[15] = use
                    nxt = use

    return {
        "gen": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "today": (dt.date.today() - EPOCH).days,
        "epoch": EPOCH.isoformat(),
        "source": source,
        "hulls": hulls,
        "types": types,
        "areas": areas,
        "shops": shops,
        "blocks": blocks,
    }


def write_js(payload, path):
    """原子写：先写临时文件再 rename，浏览器不会读到半个文件。"""
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
    # 同时写一份「原始备份」（data.js.orig），供进度条改动写回后仍能「恢复原始」
    try:
        with io.open(path + ".orig", "w", encoding="utf-8") as fh:
            fh.write(data)
    except OSError:
        pass
    if paths is not None:
        try:
            paths.reset_edits_archive()
        except Exception:           # noqa: BLE001
            pass
    return len(data)


def parse_csv(text):
    """极简 CSV → 字典列表（支持逗号/制表符，带引号）。"""
    import csv
    sample = text[:4096]
    delim = "\t" if sample.count("\t") > sample.count(",") else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    return [dict(row) for row in reader]
