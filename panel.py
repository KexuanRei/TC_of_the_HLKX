# -*- coding: utf-8 -*-
"""
panel.py —— 窗口界面的后端 API（webhook / 取数出图 / 定时任务 / 令牌 / 日志）

server.py 把 /api/* 转到这里；页面是 app.html，由 window.py 装进 pywebview 原生窗口。
菜单里的 13 项功能在这里各有一个入口，一一对应：

    1  粘贴 webhook → 生成横道图   POST /api/jobs {"kind":"sync"}       （先存 webhook 再拉）
    2  查看工作表                  POST /api/jobs {"kind":"sheets"}
    3  查看前几行                  POST /api/jobs {"kind":"peek"}
    4  用本地 Excel 生成（离线）    POST /api/jobs {"kind":"offline"}
    5  打开横道图                  窗口里的「横道图」页 / 浏览器单独打开
    6  启动 Webhook 服务（常驻）   窗口本身就在跑服务（见 /api/state 的 server）
    7  安装/重建定时任务           POST /api/tasks {"action":"install"}
    8  暂停定时任务                POST /api/tasks {"action":"pause"}
    9  恢复定时任务                POST /api/tasks {"action":"resume"}
    10 设置 AirScript 令牌         POST /api/token
    11 设置 webhook                POST /api/webhook
    12 查看日志                    GET  /api/log
    13 检查令牌状态                POST /api/jobs {"kind":"check"}（真连一次 meta）

约定
----
· 耗时动作（拉数、出图、工作表、预览、验令牌）都开**后台任务**，前端轮询日志；
  同一时刻只跑一个（_RUN_LOCK），避免两个构建同时写 data.js。
· 只允许本机回环调用（server.py 里检查），并要求请求头 X-Panel: 1。
"""
from __future__ import annotations

import datetime as dt
import io
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback

import paths

BASE = paths.BASE
CONFIG = paths.out("gantt_config.json")               # 取数配置（webhook / 表名 / 列映射参数）
SVC_CONFIG = paths.out("config.json")                 # Webhook 服务配置（host / port / token）
TASK_NAME = "Gantt_Sync"
PANEL_VERSION = "1.2"
DATE_DIR = paths.out("DATE")                          # 在线表下载/落盘目录（程序根目录下）

FIELD_CN = {
    "hull": "船号", "model": "船型", "block": "总段", "seg": "分段号",
    "type": "分段类型", "ton": "修正吨", "area": "施工区域", "shop": "分段车间",
    "dept": "施工部门", "pb_a": "拼板开工", "pb_b": "拼板完工",
    "tt_a": "分段上胎", "tt_b": "分段下胎", "jg_a": "结构完工",
    "sn_a": "舾装开工", "sn_b": "舾装销项完工", "zx_a": "结构成品化",
    "zx_b": "转序日期", "zx_b2": "生产部转序（兜底）", "state": "船体状态",
}

_CFG = {}
_JOBS = {}
_JOBS_LOCK = threading.Lock()
_RUN_LOCK = threading.Lock()
_SEQ = 0
_DATA_CACHE = {"key": None, "payload": None}

JOB_LABEL = {
    "sync": "拉取在线表并出图",
    "sheets": "查看工作表",
    "peek": "预览前几行",
    "offline": "用本地 Excel 出图",
    "check": "检查令牌",
    "link": "下载在线表并出图",
}


def configure(cfg):
    global _CFG
    _CFG = dict(cfg or {})


# ---------------------------------------------------------------- 工具

def _now():
    return dt.datetime.now().strftime("%H:%M:%S")


def _read_json(path, default=None):
    try:
        with io.open(path, encoding="utf-8-sig") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def _write_json(path, data):
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _day_str(day):
    if not day:
        return ""
    return (dt.date(2020, 1, 1) + dt.timedelta(days=int(day))).isoformat()


def _short(text, n=58):
    text = str(text or "")
    return text[:n] + "…" if len(text) > n else text


# ---------------------------------------------------------------- 配置

def read_cfg():
    """gantt_config.json 原样读（token 里的 ${ENV:...} 占位不能展开后写回）。"""
    return _read_json(CONFIG, {}) or {}


def save_cfg(cfg):
    _write_json(CONFIG, cfg)


def svc_cfg():
    cfg = {"host": "127.0.0.1", "port": 8089, "token": "", "data": "data.js",
           "viewer": "viewer.html", "archive": "incoming", "keep": 30, "max_mb": 32}
    cfg.update({k: v for k, v in (_read_json(SVC_CONFIG, {}) or {}).items() if v is not None})
    return cfg


def token_state():
    env = os.environ.get("WPS_AIRSCRIPT_TOKEN") or ""
    raw = read_cfg().get("token") or ""
    return {"env_set": bool(env.strip()), "env_len": len(env.strip()),
            "cfg_token": raw, "cfg_is_env": "${ENV:" in raw,
            "effective": bool(env.strip()) or bool(raw and "${ENV:" not in raw)}


# ---------------------------------------------------------------- 任务引擎

class Job:
    def __init__(self, jid, kind, args):
        self.id = jid
        self.kind = kind
        self.args = args or {}
        self.status = "running"
        self.lines = []
        self.result = None
        self.error = None
        self.started = dt.datetime.now().strftime("%H:%M:%S")
        self.ended = None
        self.lock = threading.Lock()

    def log(self, text):
        text = str(text if text is not None else "").rstrip()
        if not text:
            return
        with self.lock:
            self.lines.append(text)
            if len(self.lines) > 4000:
                del self.lines[:len(self.lines) - 4000]

    def dump(self, since=0):
        with self.lock:
            lines = self.lines[max(0, int(since)):]
            total = len(self.lines)
        return {"ok": True, "id": self.id, "kind": self.kind,
                "label": JOB_LABEL.get(self.kind, self.kind), "status": self.status,
                "lines": lines, "since": since, "total": total,
                "result": self.result, "error": self.error,
                "started": self.started, "ended": self.ended,
                "done": self.status != "running"}


class _Sink(io.TextIOBase):
    """任务里 print 的内容同时进任务日志（\\r 当换行，进度不会糊成一行）。"""

    def __init__(self, job, orig):
        self.job = job
        self.orig = orig
        self.buf = ""

    def write(self, s):
        if not isinstance(s, str):
            s = str(s)
        self.buf += s
        while True:
            idx = [i for i in (self.buf.find("\n"), self.buf.find("\r")) if i >= 0]
            if not idx:
                break
            i = min(idx)
            line, self.buf = self.buf[:i], self.buf[i + 1:]
            self.job.log(line)
        if self.orig is not None:
            try:
                self.orig.write(s)
            except (ValueError, OSError):
                pass
        return len(s)

    def flush(self):
        if self.buf.strip():
            self.job.log(self.buf)
            self.buf = ""
        if self.orig is not None:
            try:
                self.orig.flush()
            except (ValueError, OSError):
                pass

    def isatty(self):
        return False


def _run_job(job, fn):
    sink = _Sink(job, sys.stdout)
    old = sys.stdout
    with _RUN_LOCK:
        sys.stdout = sink
        try:
            job.log("▶ %s 开始（%s）" % (JOB_LABEL.get(job.kind, job.kind), job.started))
            fn(job)
            job.status = "done"
            job.log("✔ 完成 %s" % _now())
        except SystemExit as exc:
            msg = str(exc.code if exc.code is not None else exc) or "任务中止"
            job.status = "error"
            job.error = msg
            job.log("✘ " + msg)
        except Exception as exc:                       # noqa: BLE001
            job.status = "error"
            job.error = "%s: %s" % (type(exc).__name__, exc)
            job.log("✘ " + job.error)
            job.log(traceback.format_exc().strip().splitlines()[-1])
        finally:
            sink.flush()
            sys.stdout = old
            job.ended = _now()


def start_job(kind, args):
    global _SEQ
    fn = {"sync": _job_sync, "sheets": _job_sheets, "peek": _job_peek,
          "offline": _job_offline, "check": _job_check, "link": _job_link}.get(kind)
    if fn is None:
        raise SystemExit("未知任务类型：%s" % kind)
    with _JOBS_LOCK:
        if len(_JOBS) > 30:
            for jid in sorted(_JOBS, key=lambda k: _JOBS[k].started)[:10]:
                if _JOBS[jid].status != "running":
                    _JOBS.pop(jid, None)
        _SEQ += 1
        job = Job("j%d" % _SEQ, kind, args)
        _JOBS[job.id] = job
    threading.Thread(target=_run_job, args=(job, fn), daemon=True,
                     name="job-%s" % job.id).start()
    return job


def get_job(jid, since=0):
    job = _JOBS.get(jid)
    if job is None:
        raise SystemExit("没有这个任务：%s" % jid)
    return job.dump(since)


def busy_job():
    for job in _JOBS.values():
        if job.status == "running":
            return job.id
    return None


# ---------------------------------------------------------------- 任务实现

def _load_config(webhook="", sheet="", save=False):
    """取一份 gantt_sync 用的配置；save=True 时把 webhook/表名写回 gantt_config.json。"""
    import gantt_sync
    import build
    if save:
        cfg = read_cfg()
        if webhook:
            cfg["webhook"] = webhook
        if sheet:
            cfg["sheet"] = sheet
        save_cfg(cfg)
    conf = gantt_sync.load_config(gantt_sync.CONFIG,
                                  {"webhook": webhook or None, "sheet": sheet or None})
    conf["header_rows"] = build.HEADER_ROWS       # 列位置按 build.py 的固定映射
    return conf


def _job_sync(job):
    import gantt_sync

    webhook = (job.args.get("webhook") or "").strip()
    sheet = (job.args.get("sheet") or "").strip()
    cur = read_cfg()
    if webhook and webhook != (cur.get("webhook") or ""):
        job.log("已更新 webhook：%s" % _short(webhook, 96))
    if sheet and sheet != (cur.get("sheet") or ""):
        job.log("已更新表名：%s" % sheet)
    conf = _load_config(webhook=webhook, sheet=sheet, save=True)
    job.log("数据表：%s" % (conf.get("sheet") or "自动取第一张"))
    payload = gantt_sync.build(conf)
    job.result = _payload_info(payload, conf.get("sheet") or "自动")


def _job_sheets(job):
    import gantt_sync

    conf = _load_config(webhook=(job.args.get("webhook") or "").strip(),
                        sheet=(job.args.get("sheet") or "").strip())
    sheets = gantt_sync.list_sheets(conf)
    job.log("共 %d 张工作表" % len(sheets))
    for s in sheets:
        job.log("  %-16s %6d 行 × %3d 列" % (s.get("name", "?"), s.get("rows", 0), s.get("cols", 0)))
    job.result = {"kind": "sheets",
                  "sheets": [{"name": s.get("name"), "rows": s.get("rows"),
                              "cols": s.get("cols")} for s in sheets]}


def _job_peek(job):
    import gantt_sync
    import build

    conf = _load_config(webhook=(job.args.get("webhook") or "").strip(),
                        sheet=(job.args.get("sheet") or "").strip())
    sheets = gantt_sync.list_sheets(conf)
    sheet = gantt_sync.pick_sheet(conf, sheets)
    job.log("工作表：%s（%d 行 × %d 列）" % (sheet["name"], sheet.get("rows", 0), sheet.get("cols", 0)))
    res = gantt_sync.call(conf, {"action": "data", "sheet": sheet["name"],
                                 "row_start": 1, "row_count": 8})
    rows = (res.get("rows") if isinstance(res, dict) else res) or []
    job.result = {"kind": "peek", "sheet": sheet["name"], "header_rows": build.HEADER_ROWS,
                  "rows": [["" if v is None else str(v) for v in r] for r in rows[:8]]}


def _job_check(job):
    import gantt_sync

    conf = _load_config(webhook=(job.args.get("webhook") or "").strip(),
                        sheet=(job.args.get("sheet") or "").strip())
    tok = conf.get("token") or ""
    job.log("令牌：%s" % (("已设置（%d 位）" % len(tok)) if tok else "未设置"))
    job.log("webhook：%s" % _short(conf.get("webhook"), 96))
    sheets = gantt_sync.list_sheets(conf)
    job.log("✔ 令牌可用，读到 %d 张工作表" % len(sheets))
    job.result = {"kind": "check", "ok_token": True, "sheets": len(sheets),
                  "token_len": len(tok)}


def _job_offline(job):
    import build

    src = (job.args.get("src") or "").strip() or build.DEFAULT_SRC
    sheet = (job.args.get("sheet") or "").strip() or build.SHEET
    if not os.path.exists(src):
        raise SystemExit("找不到 Excel：%s" % src)
    job.log("源文件：%s" % src)
    job.log("工作表：%s（表头 %d 行，列映射固定）" % (sheet, build.HEADER_ROWS))
    try:
        build.run(src, sheet)
    except RuntimeError as exc:
        raise SystemExit(str(exc))
    info = data_info()
    if not info.get("exists"):
        raise SystemExit("出图结束，但没看到 data.js")
    job.result = {"kind": "offline", "blocks": info.get("blocks"), "hulls": info.get("hulls"),
                  "gen": info.get("gen"), "range": info.get("range"),
                  "size": info.get("size"), "src": src, "sheet": sheet}


# ---------------------------------------------------------------- 在线表链接 → DATE → 出图

DATE_COL_KEYS = ("pb_a", "pb_b", "tt_a", "tt_b", "jg_a",
                 "sn_a", "sn_b", "zx_a", "zx_b", "zx_b2")


def _is_file_url(url):
    low = url.lower().split("?", 1)[0]
    return low.endswith((".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"))


def _is_webhook(url):
    return ("/sync_task" in url) or ("api/v3/ide/file" in url)


def _to_date(v):
    """把「可能是日期」的值转成 date；不是日期就原样返回。"""
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        if 20000 < v < 80000:                    # Excel 1900 序列号
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        return v
    s = str(v).strip()
    if not s or s in ("/", "-", "—", "无", "None", "nan"):
        return v
    t = (s.replace("年", "-").replace("月", "-").replace("日", "")
          .replace("/", "-").replace(".", "-").strip(" -"))
    head = t.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%y-%m-%d", "%Y-%m", "%Y%m%d"):
        try:
            return dt.datetime.strptime(head, fmt).date()
        except ValueError:
            continue
    return v


def _download_file(job, url):
    """直接下载一个 .xlsx/.xlsm 文件到 DATE；返回本地路径。"""
    import urllib.error
    import urllib.parse

    import net

    os.makedirs(DATE_DIR, exist_ok=True)
    try:
        raw = net.request(url, headers={"User-Agent": "Mozilla/5.0 gantt/1.1"},
                          timeout=120, retries=2)
    except urllib.error.HTTPError as exc:
        raise SystemExit("下载失败：HTTP %d（链接失效或需要登录）。可改用 AirScript webhook。" % exc.code)
    except (urllib.error.URLError, OSError) as exc:
        raise SystemExit("下载失败（连不上）：%s" % (getattr(exc, "reason", exc),))

    head = raw[:512]
    if (b"<!DOCTYPE" in head or b"<html" in head or b"<HTML" in head):
        raise SystemExit("下载到的是网页而不是表格（这个链接可能需要登录）。请改用 AirScript webhook。")
    if len(raw) < 4096:
        raise SystemExit("下载到的文件太小（%d 字节），可能不是有效表格。" % len(raw))

    bare = url.split("?")[0].lower()
    ext = ".xlsx"
    for e in (".xlsm", ".xlsb", ".xls"):
        if bare.endswith(e):
            ext = e
            break
    if bare.endswith(".csv"):
        raise SystemExit("不支持 CSV，请提供 .xlsx / .xlsm 下载链接，或改用 AirScript webhook。")
    if ext not in (".xlsx", ".xlsm"):
        raise SystemExit("只能读取 .xlsx / .xlsm，下载到的是 %s。" % ext)

    base = urllib.parse.unquote(url.rsplit("/", 1)[-1].split("?")[0])
    name = base if base.lower().endswith((".xlsx", ".xlsm")) else ""
    if not name:
        name = "在线表_%s%s" % (dt.datetime.now().strftime("%Y%m%d_%H%M%S"), ext)

    path = os.path.join(DATE_DIR, name)
    with open(path, "wb") as fh:
        fh.write(raw)
    job.log("已下载到 DATE：%s（%.1f KB）" % (path, len(raw) / 1024.0))
    return path


def _fetch_rows_via_webhook(job, webhook, sheet):
    """用 AirScript webhook 拉整表原始行（保持原列宽），返回 (rows, cols, sheet_name)。"""
    import gantt_sync
    conf = _load_config(webhook=webhook, sheet=sheet, save=False)
    job.log("数据表：%s" % (conf.get("sheet") or "自动取第一张"))
    sheets = gantt_sync.list_sheets(conf)
    s = gantt_sync.pick_sheet(conf, sheets)
    job.log("在线表 %s：%d 行 × %d 列" % (s["name"], s.get("rows", 0), s.get("cols", 0)))
    rows, cols = gantt_sync.fetch_rows(conf, s)
    job.log("已拉取 %d 行" % len(rows))
    return rows, cols, s["name"]


def _rows_to_xlsx(job, rows, cols):
    """把原始行（含表头，保持原列宽）写成 xlsx 落盘 DATE，日期列转成真正的日期单元格。"""
    import build
    from openpyxl import Workbook

    os.makedirs(DATE_DIR, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(DATE_DIR, "在线表_%s.xlsx" % ts)
    date_idx = set(build.COL[k] for k in DATE_COL_KEYS)

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Sheet1")
    for r in rows:
        if not r:
            continue
        ws.append([_to_date(v) if i in date_idx else v for i, v in enumerate(r)])
    wb.save(path)
    job.log("在线表已落盘：%s（%d 行 × %d 列）" % (path, len(rows), cols))
    return path


def _job_link(job):
    """在线表链接 → 下载/拉取到 DATE → 从 DATE 读 Excel 出图。"""
    url = (job.args.get("url") or "").strip()
    sheet = (job.args.get("sheet") or "").strip()
    if not url:
        raise SystemExit("先粘贴在线表链接")
    os.makedirs(DATE_DIR, exist_ok=True)
    job.log("在线表链接：%s" % _short(url, 120))

    if _is_file_url(url):
        job.log("识别为文件下载链接")
        xlsx_path = _download_file(job, url)
        used_sheet = sheet or "Sheet1"
    elif _is_webhook(url):
        # 直接粘贴了 AirScript webhook：沿用原 webhook 拉取路径
        cfg = read_cfg()
        if url != (cfg.get("webhook") or ""):
            cfg["webhook"] = url
            if sheet:
                cfg["sheet"] = sheet
            save_cfg(cfg)
            job.log("已把该地址存为 webhook")
        rows, cols, _ = _fetch_rows_via_webhook(job, url, sheet)
        if not rows:
            raise SystemExit("在线表没拉到任何行。")
        xlsx_path = _rows_to_xlsx(job, rows, cols)
        used_sheet = "Sheet1"
    else:
        # kdocs/WPS 分享链接：用已登录 Edge 直接下载成 xlsx（不再回退到 webhook）
        job.log("识别为 kdocs/WPS 分享链接，用已登录 Edge 直接下载")
        import kdocs_download
        info = kdocs_download.download_kdocs(url, DATE_DIR, sheet_name=sheet, progress=job.log)
        xlsx_path = info["path"]
        used_sheet = info["sheet"]

    job.log("从 DATE 读取 Excel 并出图：%s" % os.path.basename(xlsx_path))
    import build
    try:
        build.run(xlsx_path, used_sheet)
    except RuntimeError as exc:
        raise SystemExit(str(exc))
    info = data_info()
    if not info.get("exists"):
        raise SystemExit("出图结束，但没看到 data.js")
    job.result = {"kind": "link", "blocks": info.get("blocks"), "hulls": info.get("hulls"),
                  "gen": info.get("gen"), "range": info.get("range"),
                  "size": info.get("size"), "url": url,
                  "saved": xlsx_path, "file": os.path.basename(xlsx_path)}


def _payload_info(payload, sheet=""):
    days = []
    for b in payload.get("blocks", []):
        days.extend([d for d in b[6:16] if d])
    return {"kind": "sync", "sheet": sheet, "gen": payload.get("gen"),
            "blocks": len(payload.get("blocks") or []),
            "hulls": len(payload.get("hulls") or []),
            "types": len(payload.get("types") or []),
            "range": [_day_str(min(days)), _day_str(max(days))] if days else [],
            "size": os.path.getsize(os.path.join(BASE, "data.js"))
            if os.path.exists(os.path.join(BASE, "data.js")) else 0}


# ---------------------------------------------------------------- 数据 / 状态

def data_info():
    path = os.path.join(BASE, "data.js")
    if not os.path.exists(path):
        return {"exists": False, "file": "data.js"}
    mtime = os.path.getmtime(path)
    size = os.path.getsize(path)
    key = (mtime, size)
    payload = _DATA_CACHE["payload"] if _DATA_CACHE["key"] == key else None
    if payload is None:
        try:
            with io.open(path, encoding="utf-8", errors="replace") as fh:
                text = fh.read()
            payload = json.loads(text.split("=", 1)[1].rstrip().rstrip(";"))
            _DATA_CACHE["key"] = key
            _DATA_CACHE["payload"] = payload
        except (OSError, ValueError, IndexError):
            payload = None
    if not isinstance(payload, dict):
        return {"exists": True, "file": "data.js", "size": size,
                "mtime": dt.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"),
                "error": "data.js 解析失败，重新出一次图"}
    days = []
    for b in payload.get("blocks") or []:
        days.extend([d for d in b[6:16] if d])
    return {"exists": True, "file": "data.js", "size": size,
            "mtime": dt.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "age_sec": int(dt.datetime.now().timestamp() - mtime),
            "gen": payload.get("gen"), "source": payload.get("source"),
            "blocks": len(payload.get("blocks") or []),
            "hulls": len(payload.get("hulls") or []),
            "range": [_day_str(min(days)), _day_str(max(days))] if days else []}


def excel_candidates():
    """离线出图可用的 Excel：build.py 的默认值 + 上一层的 *.xlsx。"""
    import build
    out = []
    seen = set()

    def add(path, why):
        path = os.path.abspath(path)
        if path in seen or not os.path.exists(path):
            return
        seen.add(path)
        try:
            size = os.path.getsize(path)
            mtime = dt.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
        except OSError:
            return
        if size <= 4096:                       # 空的「新建 Excel」不算
            return
        out.append({"path": path, "name": os.path.basename(path), "size": size,
                    "mtime": mtime, "why": why})

    add(build.DEFAULT_SRC, "build.py 默认")
    for folder in (os.path.dirname(BASE), BASE):
        if not os.path.isdir(folder):
            continue
        for fn in sorted(os.listdir(folder)):
            if fn.lower().endswith((".xlsx", ".xlsm")) and not fn.startswith("~$"):
                add(os.path.join(folder, fn), "找到的表格")
    out.sort(key=lambda x: (x["why"] != "build.py 默认", x["mtime"]), reverse=True)
    return out


def _run_quiet(cmd, timeout=30):
    """跑一个系统命令，返回 (returncode, 文本)。schtasks 输出是本机编码。"""
    kw = {}
    if os.name == "nt":
        kw["creationflags"] = 0x08000000          # CREATE_NO_WINDOW：不闪黑框
    try:
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                              timeout=timeout, **kw)
    except (OSError, subprocess.TimeoutExpired) as exc:
        return 1, "%r" % (exc,)
    raw = proc.stdout or b""
    for enc in ("mbcs", "gbk", "utf-8"):
        try:
            return proc.returncode, raw.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return proc.returncode, raw.decode("utf-8", "replace")


def _parse_kv(text):
    out = {}
    for line in (text or "").splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip()
        value = " ".join(value.split())            # 固定宽度输出里有多余空格
        if key and key not in out:
            out[key] = value
    return out


def task_status():
    """计划任务 Gantt_Sync 的状态（没安装时 installed=False）。"""
    code, text = _run_quiet(["schtasks", "/query", "/tn", TASK_NAME, "/fo", "LIST", "/v"])
    if code != 0 or not text.strip():
        return {"installed": False, "name": TASK_NAME, "state": "missing",
                "label": "未安装", "raw": _short(text.strip(), 300)}
    kv = _parse_kv(text)
    # schtasks 中文环境里状态叫「模式 / 计划任务状态」，英文叫 Status；都要认
    status = (kv.get("模式") or kv.get("计划任务状态") or kv.get("状态")
              or kv.get("Status") or "").strip()
    disabled = ("禁用" in status) or status.lower().startswith("disabled")
    running = ("运行" in status) or status.lower().startswith("running")
    state = "disabled" if disabled else ("running" if running else "ready")
    return {
        "installed": True, "name": TASK_NAME, "state": state,
        "label": {"disabled": "已暂停", "running": "正在运行", "ready": "就绪"}[state],
        "status_raw": status,
        "next": kv.get("下次运行时间") or kv.get("Next Run Time") or "",
        "last": kv.get("上次运行时间") or kv.get("Last Run Time") or "",
        "result": kv.get("上次结果") or kv.get("Last Result") or "",
        "command": kv.get("要运行的任务") or kv.get("Task To Run") or "",
        "schedule": kv.get("计划类型") or kv.get("Schedule Type") or "",
        "repeat": kv.get("重复") or kv.get("重复: 每隔") or kv.get("Repeat: Every") or "",
    }


def task_action(data):
    action = (data.get("action") or "").strip()
    if action == "install":
        try:
            minutes = max(5, int(data.get("minutes") or 120))
        except (TypeError, ValueError):
            minutes = 120
        if paths.FROZEN:
            cmd = '"%s" --gantt-sync' % sys.executable
        else:
            cmd = '"%s" "%s"' % (sys.executable, os.path.join(BASE, "gantt_sync.py"))
        code, text = _run_quiet(["schtasks", "/Create", "/F", "/SC", "MINUTE",
                                 "/MO", str(minutes), "/TN", TASK_NAME, "/TR", cmd])
        if code != 0:
            raise SystemExit("创建计划任务失败（退出码 %d）：%s" % (code, _short(text, 300)))
        time.sleep(0.5)                                # 让 schtasks 落盘
        return {"ok": True, "action": action, "minutes": minutes,
                "output": _short(text, 300), "task": task_status()}
    if action in ("pause", "resume"):
        if not task_status().get("installed"):
            raise SystemExit("还没安装计划任务，先点「安装 / 重建」")
        flag = "/ENABLE" if action == "resume" else "/DISABLE"
        code, text = _run_quiet(["schtasks", "/Change", "/TN", TASK_NAME, flag])
        if code != 0:
            raise SystemExit("%s 计划任务失败（退出码 %d）：%s"
                             % ("恢复" if action == "resume" else "暂停", code, _short(text, 300)))
        time.sleep(0.5)                                # /Change 之后状态要等一下才刷出来
        return {"ok": True, "action": action, "output": _short(text, 300),
                "task": task_status()}
    if action == "run":
        if not task_status().get("installed"):
            raise SystemExit("还没安装计划任务，先点「安装 / 重建」")
        code, text = _run_quiet(["schtasks", "/Run", "/TN", TASK_NAME])
        if code != 0:
            raise SystemExit("立即运行失败（退出码 %d）：%s" % (code, _short(text, 300)))
        return {"ok": True, "action": action, "output": _short(text, 300),
                "task": task_status()}
    if action == "delete":
        code, text = _run_quiet(["schtasks", "/Delete", "/F", "/TN", TASK_NAME])
        if code != 0:
            raise SystemExit("删除计划任务失败（退出码 %d）：%s" % (code, _short(text, 300)))
        return {"ok": True, "action": action, "output": _short(text, 300),
                "task": task_status()}
    raise SystemExit("未知操作：%s" % action)


def state():
    cfg = read_cfg()
    svc = svc_cfg()
    # 实际监听的 host/port 以运行时的 CFG 为准（端口被占时 window.py 会往后挪）
    svc["host"] = _CFG.get("host") or svc.get("host")
    svc["port"] = _CFG.get("port") or svc.get("port")
    webhook = cfg.get("webhook") or ""
    return {
        "ok": True, "panel": PANEL_VERSION, "base": BASE, "python": sys.version.split()[0],
        "now": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cfg": {"webhook": webhook, "webhook_short": _short(webhook),
                "sheet": cfg.get("sheet") or "", "header_rows": cfg.get("header_rows", 4),
                "rows_per_call": cfg.get("rows_per_call", 500),
                "max_rows": cfg.get("max_rows", 200000), "timeout": cfg.get("timeout", 120),
                "link_url": cfg.get("link_url") or "",
                "reset_edits_on_rebuild": cfg.get("reset_edits_on_rebuild", True)},
        "token": token_state(),
        "data": data_info(),
        "server": {"host": svc.get("host"), "port": svc.get("port"),
                   "url": "http://127.0.0.1:%s/" % svc.get("port"),
                   "push_url": "http://127.0.0.1:%s/webhook/gantt" % svc.get("port"),
                   "push_token": bool(svc.get("token"))},
        "task": task_status(),
        "excel": excel_candidates(),
        "columns": _column_map(),
        "job": busy_job(),
    }


def _column_map():
    """固定列映射（build.py 的 COL），给「确认列位置」用。"""
    try:
        import build
    except Exception:                                  # noqa: BLE001  （没装 openpyxl 也能看界面）
        return []
    rows = []
    for key, idx in build.COL.items():
        rows.append({"key": key, "cn": FIELD_CN.get(key, key), "col": idx + 1})
    rows.sort(key=lambda r: r["col"])
    return rows


# ---------------------------------------------------------------- 配置写入 / 令牌

def set_webhook(data):
    cfg = read_cfg()
    webhook = str(data.get("webhook") or "").strip()
    sheet = str(data.get("sheet") or "").strip()
    if not webhook:
        raise SystemExit("webhook 不能为空")
    if not webhook.lower().startswith(("http://", "https://")):
        raise SystemExit("webhook 看起来不对，应该是 https://…/sync_task")
    cfg["webhook"] = webhook
    cfg["sheet"] = sheet
    save_cfg(cfg)
    return {"ok": True, "state": state()}


def set_link_url(data):
    cfg = read_cfg()
    link = str(data.get("link_url") or "").strip()
    if link and not link.lower().startswith(("http://", "https://")):
        raise SystemExit("默认链接看起来不对，应该是 https://…")
    cfg["link_url"] = link
    save_cfg(cfg)
    return {"ok": True, "link_url": link, "state": state()}


def set_token(data):
    token = str(data.get("token") or "").strip()
    persist = bool(data.get("persist"))
    if token:
        os.environ["WPS_AIRSCRIPT_TOKEN"] = token
    else:
        os.environ.pop("WPS_AIRSCRIPT_TOKEN", None)
    # 同时写进 gantt_config.json：不依赖环境变量也不会丢（打包版换机 / 重启后仍可用）
    if token or str(read_cfg().get("token") or ""):
        cfg = read_cfg()
        cfg["token"] = token
        save_cfg(cfg)
    persisted, error = False, None
    if persist and token:
        try:
            subprocess.run(["setx", "WPS_AIRSCRIPT_TOKEN", token], check=False,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=20)
            persisted = True
        except Exception as exc:                       # noqa: BLE001
            error = "写入用户环境变量失败：%r" % (exc,)
    return {"ok": True, "persisted": persisted, "error": error,
            "token": token_state(), "state": state()}


# ---------------------------------------------------------------- 日志 / 目录

LOG_FILES = {"gantt": "gantt.log", "service": "service.log", "window": os.path.join("logs", "window.log")}


def read_log(name="gantt", tail=400):
    fn = LOG_FILES.get(name or "gantt", LOG_FILES["gantt"])
    path = os.path.join(BASE, fn)
    if not os.path.exists(path):
        return {"ok": True, "lines": [], "path": path, "missing": True}
    try:
        with io.open(path, encoding="utf-8", errors="replace") as fh:
            lines = fh.read().splitlines()
    except OSError as exc:
        raise SystemExit("读日志失败：%r" % (exc,))
    n = max(1, min(int(tail or 400), 5000))
    return {"ok": True, "lines": lines[-n:], "path": path, "size": os.path.getsize(path),
            "mtime": dt.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")}


def reveal(data):
    target = (data.get("target") or "base").strip()
    paths = {"base": BASE, "data": os.path.join(BASE, "data.js"),
             "config": CONFIG, "svc_config": SVC_CONFIG,
             "logs": BASE, "incoming": os.path.join(BASE, svc_cfg().get("archive") or "incoming")}
    path = paths.get(target) or BASE
    if not os.path.exists(path):
        raise SystemExit("还没有：%s" % path)
    try:
        os.startfile(path)                             # noqa: S606
    except OSError as exc:
        raise SystemExit("打不开：%r" % (exc,))
    return {"ok": True, "path": path}


# ---------------------------------------------------------------- 导出 Excel（样式与当前横道图一致）

def export_xlsx(data):
    """把查看器算好的当前视图写成一个 Excel（左 6 列 + 按天的条形图）。返回 (bytes, filename)。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    phase_name = data.get("phaseName") or data.get("phase") or ""
    scale = data.get("scale") or "month"
    days = data.get("days") or []
    months = data.get("months") or []
    rows = data.get("rows") or []
    legend = data.get("legend") or []
    title = data.get("title") or "分段生产横道图"
    rng = data.get("range") or []

    n_days = len(days)
    N_LEFT = 6
    left_cols = data.get("leftCols") or ["分段 / 船", "类型", "吨位", "车间", "开始", "工期"]
    total_cols = N_LEFT + n_days
    scale_label = {"month": "月视图", "week": "周视图", "day": "日视图"}.get(scale, scale)

    wb = Workbook()
    ws = wb.active
    ws.title = "横道图"

    hdr_fill = PatternFill("solid", fgColor="F4F8FE")
    month_fill = PatternFill("solid", fgColor="FBFDFF")
    weekend_fill = PatternFill("solid", fgColor="F7F9FD")
    today_fill = PatternFill("solid", fgColor="E5484D")
    today_col_fill = PatternFill("solid", fgColor="FDECED")
    thin = Side(style="thin", color="E9EEF6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    gray = Font(color="8B97AB", size=9)

    # 标题 + 副标题
    rng_text = ("%s ~ %s" % (rng[0], rng[1])) if len(rng) >= 2 else ""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    t = ws.cell(1, 1, "%s · %s · %s" % (title, phase_name, scale_label))
    t.font = Font(bold=True, size=14)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    s2 = ws.cell(2, 1, "数据时段　%s" % rng_text)
    s2.font = gray
    s2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    HEADER_ROW = 3
    for i, name in enumerate(left_cols):
        c = ws.cell(HEADER_ROW, i + 1, name)
        c.font = Font(bold=True, size=10, color="55627A")
        c.fill = hdr_fill
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    # 月份表头（按天合并）
    col = N_LEFT + 1
    for m in months:
        span = max(1, int(m.get("span") or 1))
        text = str(m.get("text") or "")
        if span > 1:
            ws.merge_cells(start_row=HEADER_ROW, start_column=col,
                           end_row=HEADER_ROW, end_column=col + span - 1)
        c = ws.cell(HEADER_ROW, col, text)
        c.font = Font(bold=True, size=9, color="33415C" if "年" in text else "4F5D75")
        c.fill = month_fill
        c.border = border
        c.alignment = Alignment(horizontal="left", vertical="center")
        col += span
    ws.row_dimensions[HEADER_ROW].height = 20

    # 日号行（仅日视图，和查看器一致）
    show_days = (scale == "day")
    day_row = HEADER_ROW + 1 if show_days else HEADER_ROW
    if show_days:
        for i, d in enumerate(days):
            c = ws.cell(HEADER_ROW + 1, N_LEFT + 1 + i, d.get("num"))
            is_today = bool(d.get("today"))
            c.font = Font(size=8, bold=is_today, color="FFFFFF" if is_today else "8B97AB")
            if is_today:
                c.fill = today_fill
            elif d.get("wd") in (0, 6):
                c.fill = weekend_fill
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center")

    # 数据行
    today_cols = [i for i, d in enumerate(days) if d.get("today")]
    r = day_row + 1
    for row in rows:
        cells = row.get("cells") or []
        for i in range(N_LEFT):
            c = ws.cell(r, i + 1, cells[i] if i < len(cells) else "")
            c.font = Font(size=10)
            c.border = border
            c.alignment = Alignment(horizontal="center" if i >= 4 else "left", vertical="center")
        # 今天列：跨数据行铺浅红底（模拟今天的红色虚线），条会盖在它上面
        for ti in today_cols:
            ws.cell(r, N_LEFT + 1 + ti).fill = today_col_fill
        bar = row.get("bar")
        if bar:
            c0 = max(0, int(bar.get("c0") or 0))
            c1 = min(n_days - 1, int(bar.get("c1") if bar.get("c1") is not None else c0))
            color = (bar.get("color") or "#4d8bff").lstrip("#")
            fill = PatternFill("solid", fgColor=color)
            for k in range(c0, c1 + 1):
                cell = ws.cell(r, N_LEFT + 1 + k)
                cell.fill = fill
                cell.border = border
            label = bar.get("label") or ""
            if label and (c1 - c0 + 1) >= 3:
                lc = ws.cell(r, N_LEFT + 1 + c0, label)
                lc.font = Font(size=8, bold=True, color="FFFFFF")
                lc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1

    # 图例行
    lr = r + 1
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=N_LEFT)
    ws.cell(lr, 1, "图例：").font = gray
    ci = 0
    for lg in legend:
        cc = ws.cell(lr, N_LEFT + 1 + ci * 3)
        cc.value = " %s " % (lg.get("name") or "")
        cc.fill = PatternFill("solid", fgColor=(lg.get("color") or "888888").lstrip("#"))
        cc.font = Font(size=8, bold=True, color="FFFFFF")
        cc.alignment = Alignment(horizontal="center", vertical="center")
        ci += 1

    # 列宽 / 冻结
    for i, w in enumerate([20, 10, 7, 11, 8, 8]):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    for i in range(n_days):
        ws.column_dimensions[get_column_letter(N_LEFT + 1 + i)].width = 3.0
    ws.freeze_panes = "%s%d" % (get_column_letter(N_LEFT + 1), day_row + 1)

    buf = io.BytesIO()
    wb.save(buf)
    fname = "横道图_%s.xlsx" % dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return buf.getvalue(), fname


# ---------------------------------------------------------------- 路由

def api(method, path, query=None, body=b""):
    query = query or {}
    try:
        if method == "GET":
            if path == "/api/state":
                return 200, state()
            if path == "/api/version":
                return 200, {"ok": True, "name": "gantt-window", "panel": PANEL_VERSION,
                             "python": sys.version.split()[0]}
            if path == "/api/log":
                return 200, read_log(query.get("name", "gantt"), query.get("tail", 400))
            if path == "/api/tasks":
                return 200, {"ok": True, "task": task_status()}
            if path.startswith("/api/jobs/"):
                return 200, get_job(path.rsplit("/", 1)[-1], query.get("since", 0))
            return 404, {"ok": False, "error": "未知接口：%s" % path}

        raw = body.decode("utf-8-sig") if isinstance(body, (bytes, bytearray)) else (body or "")
        data = json.loads(raw) if raw.strip() else {}
        if path == "/api/jobs":
            kind = str(data.get("kind") or "").strip()
            job = start_job(kind, data.get("args") or {})
            return 200, {"ok": True, "id": job.id, "label": JOB_LABEL.get(kind, kind)}
        if path == "/api/webhook":
            return 200, set_webhook(data)
        if path == "/api/link-url":
            return 200, set_link_url(data)
        if path == "/api/edits-reset":
            enabled = bool(data.get("enabled"))
            cfg = read_cfg()
            cfg["reset_edits_on_rebuild"] = enabled
            save_cfg(cfg)
            return 200, {"ok": True, "reset_edits_on_rebuild": enabled, "state": state()}
        if path == "/api/token":
            return 200, set_token(data)
        if path == "/api/tasks":
            return 200, task_action(data)
        if path == "/api/reveal":
            return 200, reveal(data)
        if path == "/api/kdocs-open":
            import kdocs_download
            url = str(data.get("url") or "").strip()
            if not url:
                raise SystemExit("先粘贴在线表链接")
            opened = kdocs_download.open_kdocs_login(url)
            return 200, {"ok": True, "opened": opened,
                         "message": "已打开下载专用 Edge，请在其中登录 kdocs；"
                                    "登录后关闭该窗口，再点「下载并生成横道图」"}
        return 404, {"ok": False, "error": "未知接口：%s" % path}
    except SystemExit as exc:
        msg = str(exc.code if exc.code is not None else exc) or "操作失败"
        return 400, {"ok": False, "error": msg}
    except ValueError as exc:
        return 400, {"ok": False, "error": "请求格式不对：%s" % exc}
    except Exception as exc:                           # noqa: BLE001
        return 500, {"ok": False, "error": "%s: %s" % (type(exc).__name__, exc)}
