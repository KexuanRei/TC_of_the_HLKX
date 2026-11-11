/**
 * airscript_sync_task.js —— 横道图 AirScript 取数脚本（sync_task 协议）
 * =====================================================================
 * 作用：让本机程序通过 webhook + 脚本令牌，从在线表格里分页读取整张表。
 *
 * 协议（与 gantt_sync.py 对应）：
 *   action=meta  → { sheets: [ { name, rows, cols }, ... ] }
 *   action=data  → { rows: [ [ ... ], ... ], next_row: N }
 *
 * 用法：
 *   1. 打开在线表 → 工具 → 开发工具 → AirScript；
 *   2. 打开现有的「横道图脚本」（或新建一个），全选删除，把本文件全部内容粘进去；
 *   3. 保存（Ctrl+S），点「运行」自测，日志里应看到 [横道图] 之类信息且不报错；
 *   4. 回到本机程序点「下载并生成横道图」。
 *
 * 注意：
 *   · 本脚本不含任何 Node.js 专用对象（process / require / module 等），
 *     在 WPS AirScript 里可以直接跑；之前报 "process is not defined" 就是残留了 Node 代码。
 *   · 脚本通过顶层 return 返回结果，WPS 会把它放到 data.result 里。
 */

/* ---------------- 小工具 ---------------- */

function toInt(v, def) {
  var n = parseInt(v, 10);
  return isNaN(n) ? def : n;
}

function isArray(v) {
  return Object.prototype.toString.call(v) === '[object Array]';
}

// 列号 → Excel 列字母：1→A, 26→Z, 27→AA, 80→CB
function colLetter(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

// Value2 可能是 2 维数组，也可能是单行/单列的 1 维数组，统一成 2 维
function to2d(v) {
  if (v === undefined || v === null) return [];
  if (!isArray(v)) return [[v]];
  if (v.length === 0) return [];
  if (isArray(v[0])) return v;
  return [v];
}

function sheetsColl() {
  return Application.Worksheets || Application.Sheets;
}

function getSheet(name) {
  var coll = sheetsColl();
  if (name) {
    try { return coll.Item(name); } catch (e) { /* 名字不对就用活动表 */ }
  }
  return Application.ActiveSheet;
}

/**
 * 取一张表的绝对范围：最后一行的行号、最后一列的列号。
 * 用 UsedRange 的 Row/Column 是为了兼容「第 1 行是空行」的情况：
 * 返回的是绝对行号，程序从第 1 行开始读，表头（前 4 行）自然不会错位。
 */
function sheetSize(ws) {
  var ur = ws.UsedRange;
  var lastRow = ur.Row + ur.Rows.Count - 1;
  var lastCol = ur.Column + ur.Columns.Count - 1;
  if (lastRow < 1) lastRow = 1;
  if (lastCol < 1) lastCol = 1;
  return { rows: lastRow, cols: lastCol };
}

/* ---------------- 两个动作 ---------------- */

function doMeta() {
  var coll = sheetsColl();
  var n = coll.Count;
  var out = [];
  for (var i = 1; i <= n; i++) {
    var ws = coll.Item(i);
    var name = ws.Name;
    try {
      var sz = sheetSize(ws);
      out.push({ name: name, rows: sz.rows, cols: sz.cols });
    } catch (e) {
      out.push({ name: name, rows: 0, cols: 0 });
    }
  }
  return { sheets: out };
}

function doData(argv) {
  var ws = getSheet(argv.sheet || '');
  var sz = sheetSize(ws);

  var start = toInt(argv.row_start, 1);
  if (start < 1) start = 1;
  var count = toInt(argv.row_count, 500);
  if (count < 1) count = 500;

  var end = Math.min(sz.rows, start + count - 1);
  if (end < start) return { rows: [], next_row: sz.rows + 1 };

  var a1 = 'A' + start + ':' + colLetter(sz.cols) + end;
  var vals = ws.Range(a1).Value2;
  return { rows: to2d(vals), next_row: end + 1 };
}

/* ---------------- 入口 ---------------- */

function handle() {
  var argv = {};
  try {
    if (typeof Context !== 'undefined' && Context && Context.argv) {
      argv = Context.argv;
      if (typeof argv === 'string') argv = JSON.parse(argv);
    }
  } catch (e) {
    argv = {};
  }

  var action = argv.action || 'meta';
  if (action === 'meta') return doMeta();
  if (action === 'data') return doData(argv);
  return { error: 'unknown action: ' + action };
}

// 顶层 return：WPS 会把返回值塞进 data.result，本机程序再从里面取 JSON。
// 这里显式 JSON.stringify 成字符串，避免个别运行时把对象转成 "[object Object]"。
return JSON.stringify(handle());
