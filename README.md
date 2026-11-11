# 分段生产横道图工具

把工作区里的 **《数据更新中_在建.xlsx》** 直接变成可交互的横道图。Excel 更新后重跑一次即可，不用改代码。

```
数据更新中_在建.xlsx ──build.py──▶ data.js ──viewer.html──▶ 浏览器里的横道图
```

## 窗口界面（推荐日常用）

双击 **`窗口.bat`** —— 弹出一个独立窗口（原生窗口，没有地址栏/标签页），菜单里的 13 项都能在窗口里点：

| 页面 | 对应菜单 |
|---|---|
| **横道图** | 5 打开横道图（内嵌，点船钻取 / 阶段切换 / 搜索缩放都在里面） |
| **取数与出图** | 1 粘贴 webhook → 拉取出图；2 查看工作表；3 预览前几行；4 本地 Excel 离线出图；11 设置 webhook；以及固定列映射对照表 |
| **定时任务** | 7 安装/重建；8 暂停；9 恢复；外加「立即运行 / 删除」和状态（下次/上次运行） |
| **日志** | 12 查看日志（gantt.log / service.log / window.log） |
| **设置** | 10 设置 AirScript 令牌；13 检查令牌（真连一次）；Webhook 服务地址、目录入口 |

- 底部是**任务面板**：拉取/出图的进度实时滚动，完成后给出分段数、船数、时段、data.js 大小。
- 窗口壳是 **pywebview + Edge WebView2**；第一次运行若没装 pywebview 会先 `pip install`，装不上就退回 Edge 应用窗口（同样是独立窗口）。
- 窗口本身就在跑 Webhook 服务（= 菜单 6），关窗口即停；要常驻后台还是用 `run_server.bat`。
- 参数：`窗口.bat --debug` 开调试；`python window.py --no-window` 只起服务；`--selftest` 开窗 4 秒自检。
- 控制台接口只认本机（`X-Panel` 头 + 回环地址），局域网浏览器只能看 `/` 图表。

### 在线表链接直出图（推荐）

「取数与出图」页顶部新增 **在线表链接**：粘贴 kdocs/WPS 分享链接（或 AirScript webhook、`.xlsx` 下载链接）→ 点 **下载并生成横道图**：

1. 程序用已登录的 **Edge** 直接把在线表下载到程序根目录的 **`DATE`** 文件夹（不再回退到 AirScript webhook；直接粘贴 `.xlsx` 下载链接或 webhook 仍走各自原有路径）；
2. 再从 DATE 里刚下载的 Excel 读出横道图（出图逻辑不变）。

首次使用会弹出下载专用 Edge 窗口，请在其中登录 kdocs（登录态保留在程序根目录 `kdocs_profile`）。此功能依赖 `playwright`：`pip install playwright`（channel 用系统 Edge，无需 `playwright install` 下载浏览器）。

### 导出 Excel（与当前横道图样式一致）

横道图右上角新增 **导出 Excel** 按钮：把当前视图（阶段 / 月周日缩放 / 船或搜索 / 时间轴范围）按原样式输出成 Excel——左 6 列（分段·船 / 类型 / 吨位 / 车间 / 开始 / 工期）+ 按天的彩色条形图，含月份表头、周末/今天底色、冻结窗格和图例。文件同时存到 `DATE` 文件夹。独立用浏览器直接打开 `viewer.html`（file://）时没有本地服务，导出会提示需要窗口界面。

## 一键使用

双击 **`run.bat`**（或命令行 `python build.py`）。

- 自动读 Excel → 生成 `data.js` → 打开 `viewer.html`
- 只读 Excel，**不回写、不改你的源文件**
- 生成的页面不需要服务器，双击就能开，断网也能用

改完 Excel 直接再跑一次，页面刷新即可。

## 换数据源

源文件和工作表名写在 `build.py` 顶部两行：

```python
DEFAULT_SRC = r"D:\DSP6\数据更新中_在建.xlsx"
SHEET = "Sheet1"
```

- 文件名或工作表名变了 → 改这两行，之后 `run.bat` 一直用新的；
- 只想临时换一次 → 命令行传参，不改代码：

```bat
python build.py --src "D:\DSP6\别的表.xlsx" --sheet Sheet1
```

> 列号不用动：这张表的列位置和 `COL` 完全一致（船号第 5 列、新段号第 9 列、拼板开工第 16 列、转序日期第 51 列、船体状态第 75 列…）。
> 换一张列序不同的表，才需要改 `COL`。

## 数据来源与列映射

只读 **`Sheet1`** 表（在建）：三层表头（第 2/3/4 行），数据从第 5 行开始，共 19,998 行。

| 用途 | 列 | 说明 |
|---|---|---|
| 船号 | 5 船号 | 33 艘，分组依据 |
| 分段号 | 9 新段号（空则取 6 总段） | 行标识 |
| 类型 | 13 分段类型 | 223 种 |
| 吨位 | 15 修正吨 | |
| 车间 | 56 分段车间 / 53 施工区域 | 14 个车间、50 个区域 |
| 拼板 | 16 拼板开工 → 17 拼板完工 | |
| 在胎 | 19 分段上胎 → 20 分段下胎 | |
| 结构 | 20 下胎 → 22 结构完工 | 下胎缺失时用上胎 |
| 舾装 | 38 开工 → 42 销项完工 | |
| 转序 | 48 结构成品化 → **51 转序日期** | 有转序日期 = 该分段已完工，工期就到这天；空时用 52 生产部转序兜底 |
| 完工判定 | **75 BW「船体状态」** | `完成` / `在建`。标"完成"但缺转序日期的，按工序顺序借用下一个分段的转序日期，整船都没有则用该船最晚转序日 |

只有起点没有终点的，按"当天开始当天结束"补成 1 天；一个日期都没有的行会被丢弃（这张在建表里是 12,246 行）。

列号要改的话，只动 `build.py` 顶部的 `COL = {...}`。

## 三种视图

**① 船队总览**（默认）
33 艘船，每行一条汇总条 = 该船所有分段在所选阶段的最早开始 → 最晚结束。时间轴右端 = **数据里最晚日期 + 3 个月**（数据到 10 月就铺到次年 1 月底，到 11 月就铺到次年 2 月底），所以末尾那几个月的空档也看得到。

**② 点船钻取**
点任意船号行 → 只看这艘船的全部段（约 200–560 行）。进去后：
- 时间轴只铺这艘船自己的周期，从它最早那天开始，不显示前面的空白
- 按集装箱宽度自适应，船短也不会只占左边半截
- 行按该阶段的开工日排序，画出来是**阶梯状**，不是 Excel 原顺序那种乱堆

左上「← 全部船」返回。

**④ 未填完工日期**
某一阶段只填了开始、没填结束，**并且这个分段还没转序（没有转序日期）** → 条一直画到**今天**，右端带斜纹，工期列带 `*`，悬停显示"未完工（画到今天）"。\n\n已经转序的分段即使某个阶段缺结束日期，也**不会**画到今天——它已经完工了。

在 Excel 里把结束日期补上、重跑 `run.bat`，它就会变成正常的一段。

**⑤ 完全没有日期**
某一阶段连开始都没填 → 开始列显示 `—`、工期列显示 `待定`（灰字），不画条；鼠标悬停显示"还没有具体时间 / 工期"。补上日期后重跑即可。

**③ 阶段切换**
右上角下拉：在胎 / 拼板 / 结构 / 舾装 / 转序 / 全流程。整张图按该阶段重画——用来回答"这个月谁在胎、谁该转序"。

其它：`月/周/日` 三档缩放、分段号/船号搜索、定位今天、悬停看该分段所有节点日期。

推荐把常用视图存成书签：

```
viewer.html?hull=HL-C21K-6              # 直接进某艘船
viewer.html?phase=zx                    # 直接看转序
viewer.html?hull=HL-C21K-6&scale=day    # 某艘船的日视图
```

## 性能

7,199 行数据，页面 DOM 里始终只有 30 来行——行虚拟化 + 节点复用，滚动不掉帧。数据以 `data.js`（0.4 MB）形式内联，`file://` 直接加载，不需要起服务。

## 已知限制 / 下一步

- **只读**：不能在图里拖拽改期回写 Excel。要编辑的话得另做一套写回逻辑。
- **没有依赖线**：分段之间没有强制的工序依赖数据，所以只画条不画箭头。
- **工序映射是推断的**：例如"结构 = 下胎→结构完工"、"转序 = 成品化→转序日期"。如果现场口径不同，改 `viewer.html` 里 `PHASES` 的 `s/e` 索引即可。
- **没做船级线表视图**：线表台账里有开工/铺底/下水/交船四个节点，可以再加一个"船台计划"页签（那份 `船号.xlsx` 目前不在工作区）。
- 文件被 Excel 独占打开时读不了，先关掉再跑。

## 连不上 / 报 10053 怎么办

`URLError(ConnectionAbortedError(10053, ...))`、`调用失败（网络不通）` 这一类，绝大多数是
**系统里残留的代理设置**：Internet 选项 → 局域网设置里勾着 `127.0.0.1:21081`（Clash / v2ray 之类），
代理软件退出了但设置还在，于是 Python 的 `urllib` 去连那个死端口。

程序已经处理过：`net.py` 里的请求顺序是 **直连 → 系统/环境变量代理 → `curl.exe`**，
所以上面这种情况会自动走通。想彻底干净的话，把系统代理关掉即可。

## 文件说明

| 文件 | 作用 |
|---|---|
| `build.py` | **读 Excel → 生成 `data.js`**（`COL` 列映射、`DEFAULT_SRC`、`SHEET` 都在顶部） |
| `core.py` | 归一化 + 写 `data.js`（本地 Excel / webhook / pull 三条路共用） |
| `viewer.html` | 查看器（唯一需要维护的前端文件） |
| `data.js` | 生成物，别手改 |
| `build_report.txt` | 上一次生成的统计（行数、船数、各阶段完整起止行数） |
| `run.bat` | 一键：生成 + 打开 |
| `横道图.bat` + `menu.py` | 中文菜单（14 项，含「打开窗口界面」） |
| `窗口.bat` + `window.py` | **窗口界面入口**：起服务 + pywebview 原生窗口 |
| `app.html` | 窗口界面的控制台页面 |
| `panel.py` | 窗口界面的后端 API（取数/出图/定时任务/令牌/日志） |
| `gantt_sync.py` | 从 WPS/kdocs webhook 拉数出图（菜单 1，见第七节说明） |
| `pull.py` | 从任意 URL 拉数出图（可 `--interval` 定时） |
| `server.py` + `run_server.bat` | Webhook 服务：接收 WPS 推送、托管查看器 |
| `wps_gantt.js` | 贴进 WPS AirScript 的取数/推送脚本 |
| `airscript_sync_task.js` | **贴进 WPS AirScript 的 sync_task 取数脚本**（meta/data 分页协议，替代报 `process is not defined` 的旧脚本） |
| `config.json` | 服务配置（端口 / 令牌 / 归档） |
| `gantt_config.json` | 同步配置（webhook / 表名 / 分页） |
| `preview-overview.png` / `preview-ship.png` | 效果截图 |

---

## 六、接入方式（自动化）

三条路，选一条就行：

**① 手工（最简单）**
改 Excel → 双击 `run.bat` → 浏览器 Ctrl+F5。

**② 本机定时拉（推荐）**
```
python pull.py --url "https://你的webhook地址" --interval 300
```
每 5 分钟拉一次、自动重生成 `data.js`，浏览器开着就会看到最新数据。
支持 JSON 数组 / `{"rows":[...]}` / CSV，字段用中文表头或英文短名都行。

**③ WPS AirScript 推送**
把 `wps_gantt.js` 全选粘贴到 WPS 的 AirScript，改顶部 `CONFIG`（SOURCE_URL / GANTT_ENDPOINT / TOKEN）即可。
脚本只做两件事：取数、POST 给 `server.py`，所有归一化逻辑都在 Python 里，改口径不用碰 JS。

服务端：
```
python server.py            # 或双击 run_server.bat
```
- `POST /webhook/gantt` JSON  `POST /webhook/csv` 文本
- `GET /` 查看器  `GET /health` 健康检查
- 鉴权：`X-Token` 请求头，值见 `config.json`（默认 `change-me-please`，**上线前务必改**）
- 每次推送都会在 `incoming/` 留档，保留最近 30 份

> AirScript 跑在 WPS 的运行时里：如果它能访问本机（同一台电脑），直连 `127.0.0.1:8089` 没问题；
> 如果报 `fetch is not defined` 或连不上，说明这个版本不允许联网，改用方式 ②。

---

## 七、WPS 在线表直连（菜单）

双击 **`横道图.bat`** 进入菜单（和 `WPS同步工具.bat` 一个风格）。

> ⚠️ 如果双击后窗口一闪就退：这台机器的 PATH 里 `python` 是 Windows 商店的空壳（退出码 9009）。
> 临时办法是在 `gantt\ship` 目录下执行 `python menu.py`。

```
  1  粘贴 webhook → 生成横道图     8  暂停定时任务
  2  查看工作表（确认表名）        9  恢复定时任务
  3  查看前几行（确认列位置）      10 设置 AirScript 令牌
  4  用本地 Excel 生成（离线）     11 设置 webhook
  5  打开横道图                   12 查看日志
  6  启动 Webhook 服务（常驻）     13 检查令牌状态
  7  安装/重建定时任务             14 打开窗口界面（推荐）
```

**第一次用**：菜单 10 粘贴令牌 → 菜单 2 看表名 → 菜单 3 确认第 5 行开始是数据 → 菜单 1 生成。
这些操作在「窗口界面」里更直观（见开头）。

> ⚠️ 菜单 1 只认和 `build.py` 同一套的 `COL` 列号（79 列「临时」表）。
> 当前线上那张表若列数/列序不同（例如 24 列的「汇总」），会报"没有生成任何分段"——
> 这种情况请直接用 `run.bat` 读本地 Excel。

协议与 `wps_sync` 完全一致：`POST {"Context":{"argv":{...}}}` + 头 `AirScript-Token`，
`action=meta` 取表清单、`action=data` 分页取行（默认每页 500 行）。

> ⚠️ 这套协议要求在线表里部署的是 **`airscript_sync_task.js`**（就是实现 `action=meta` / `action=data` 的那份）。
> 如果脚本里残留了 Node 代码，会报 `process is not defined`；把 `airscript_sync_task.js` 全文替换进去即可。

- 令牌默认直接存在 `gantt_config.json` 的 `token` 字段（也可写 `${ENV:WPS_AIRSCRIPT_TOKEN}` 用环境变量）；
  在窗口「设置」里粘贴一次就会同时写进配置 + 进程环境，勾了"写入用户环境变量"还会 `setx`。180 天过期后重设即可；
- 打包版启动时 `paths.load_env()` 会从注册表补一遍缺失的用户/系统环境变量，
  这样用 `setx` 设完变量、不解散资源管理器直接双击 EXE 也能读到（否则拿到的是旧环境块）；
- 表名留空 = 自动用第一张表；写错会提示现有表名；
- 列映射与 `build.py` 共用一套（`COL` / `HEADER_ROWS`），改口径只改一处。

---

## 八、打包成免安装 EXE

已用 PyInstaller 打成单文件 **`dist\分段生产横道图.exe`**：双击即用，目标机不用装 Python。

- 打包版的可写文件（`data.js` / `gantt_config.json` / `logs` / `DATE`）放在 **EXE 旁边**；
  只读资源（`app.html` / `viewer.html` / `TM.ico`…）在 PyInstaller 的 `_MEIPASS` 里，由 `paths.py` 区分。
- 首次运行会把内置的 `config.json` / `gantt_config.json` / `data.js` 复制到 EXE 旁边（已存在不覆盖）。
- 窗口优先用 WebView2（pywebview）；没有时自动退回 Edge 应用窗口。

改了代码后重新打包：

```bat
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "分段生产横道图" ^
  --icon TM.ico ^
  --add-data "app.html;." --add-data "viewer.html;." --add-data "data.js;." ^
  --add-data "config.json;." --add-data "gantt_config.json;." --add-data "TM.ico;." ^
  --add-data "wps_gantt.js;." --add-data "airscript_sync_task.js;." --add-data "README.md;." ^
  --hidden-import server --hidden-import panel --hidden-import core --hidden-import build ^
  --hidden-import gantt_sync --hidden-import paths --hidden-import webview ^
  --collect-submodules openpyxl window.py
```

需要 Python 3.12 + `pip install pyinstaller openpyxl pywebview`。
