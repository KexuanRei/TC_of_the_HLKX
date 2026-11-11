/**
 * wps_gantt.js —— WPS AirScript 脚本：从 webhook 取数 → 推给本机横道图服务
 *
 * 在 WPS 里：工具 → 开发工具 → AirScript → 新建脚本 → 全选粘贴 → 改下面 CONFIG → 运行
 *
 * 两种取数方式（二选一）：
 *   1) 填 SOURCE_URL：直接拉线上 webhook 的 JSON
 *   2) 留空：读当前表格（第一行当表头，A1 开始的连续区域）
 *
 * 推给谁：本机跑的 server.py（python server.py），默认 http://127.0.0.1:8089
 */
var CONFIG = {
  SOURCE_URL: '',                                  // 例：'https://xxx.com/webhook/abc'
  SOURCE_HEADERS: {},                              // 例：{ 'Authorization': 'Bearer xxx' }
  GANTT_ENDPOINT: 'http://127.0.0.1:8089/webhook/gantt',
  TOKEN: 'change-me-please',                       // 与 config.json 里的 token 一致
  SHEET_NAME: '',                                  // 留空 = 当前活动表
  RANGE: 'A1'                                      // 从这个单元格开始取连续区域
};

/* ---------- 取数：webhook ---------- */
async function fetchRows() {
  var url = CONFIG.SOURCE_URL;
  if (!url) return null;

  var opt = { method: 'GET', headers: CONFIG.SOURCE_HEADERS || {} };
  var resp = await fetch(url, opt);     // AirScript 若不支持 fetch 会在这里报错
  var text = await resp.text();         // await 对 Promise 和普通值都成立
  var body = JSON.parse(text);

  if (Array.isArray(body)) return body;
  var keys = ['rows', 'data', 'items', 'list', 'records'];
  for (var i = 0; i < keys.length; i++) {
    if (Array.isArray(body[keys[i]])) return body[keys[i]];
  }
  throw new Error('返回里没有数组，试过 rows/data/items/list/records');
}

/* ---------- 取数：当前表格 ---------- */
function sheetRows() {
  var app = Application;
  var sheet = CONFIG.SHEET_NAME ? app.Worksheets.Item(CONFIG.SHEET_NAME) : app.ActiveSheet;
  var vals = sheet.Range(CONFIG.RANGE).CurrentRegion.Value2;   // 二维数组
  if (!vals || vals.length < 2) throw new Error('表格没有数据（至少要有表头和一行）');

  var head = vals[0];
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var row = {};
    var empty = true;
    for (var c = 0; c < head.length; c++) {
      var key = String(head[c] || '').replace(/\s+/g, '');
      if (!key) continue;
      var v = vals[r][c];
      row[key] = v;
      if (v !== null && v !== undefined && v !== '') empty = false;
    }
    if (!empty) out.push(row);
  }
  return out;
}

/* ---------- 推送 ---------- */
async function push(rows) {
  var body = JSON.stringify({ rows: rows });
  var resp = await fetch(CONFIG.GANTT_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-Token': CONFIG.TOKEN },
    body: body
  });

  var text = await resp.text();
  var info;
  try { info = JSON.parse(text); } catch (e) { info = { ok: false, raw: text }; }
  return info;
}

/* ---------- 入口 ---------- */
async function main() {
  try {
    var rows = await fetchRows();
    var from = 'webhook';
    if (!rows) { rows = sheetRows(); from = 'sheet'; }

    if (!rows || !rows.length) throw new Error('没取到数据');
    console.log('来源 ' + from + '，共 ' + rows.length + ' 行');

    var info = await push(rows);
    if (!info.ok) throw new Error('服务端返回失败：' + JSON.stringify(info));

    console.log('生成成功：' + info.blocks + ' 段 / ' + info.hulls + ' 船，' +
                (info.bytes / 1048576).toFixed(2) + ' MB');
    console.log('打开 http://127.0.0.1:8089/ 并 Ctrl+F5 刷新');
    return info;
  } catch (e) {
    console.log('失败：' + e.message);
    console.log('排查：①server.py 是否在跑 ②token 是否一致 ③AirScript 是否允许联网');
    return null;
  }
}

/* Node 下直接跑（方便本机测试）：node wps_gantt.js --source=... --endpoint=... */
if (typeof process !== 'undefined' && process.argv) {
  process.argv.slice(2).forEach(function (a) {
    var m = /^--([^=]+)=(.*)$/.exec(a);
    if (!m) return;
    if (m[1] === 'source') CONFIG.SOURCE_URL = m[2];
    if (m[1] === 'endpoint') CONFIG.GANTT_ENDPOINT = m[2];
    if (m[1] === 'token') CONFIG.TOKEN = m[2];
  });
  main();
} else {
  main();   // AirScript 里直接执行
}
