# -*- coding: utf-8 -*-
"""
kdocs_download.py —— 用已登录的 Edge 直接下载 kdocs/WPS 分享链接成 .xlsx

移植自「统一任务工作台」的 table_c_download.py / online_format.py，只保留“下载”
这一件事：通过可见的“导出为 Excel”菜单 + 真实浏览器 Download 事件拿到原始 xlsx，
再原子落盘到程序根目录的 DATE 目录。下载之后仍由 build.run() 读 Excel 出图，那一段
逻辑不变。

依赖：playwright（channel='msedge' 直接用系统 Edge，无需 `playwright install` 下载浏览器）。
下载用的 Edge 窗口会放到另一个 Windows 虚拟桌面（Win+Tab 可看到），不遮挡当前桌面；
首次使用请先点「打开链接（登录）」弹出可见窗口登录 kdocs（独立 profile：BASE/kdocs_profile），
登录态会被保留，之后无需重复登录。

用法：
    from kdocs_download import download_kdocs
    info = download_kdocs(url, out_dir, sheet_name="", progress=print)
    print(info)  # {'path': ..., 'sheet': ..., 'size': ...}
"""
from __future__ import annotations

import datetime as dt
import os
import re
import tempfile
import time
import urllib.parse
from pathlib import Path

try:
    import paths
except Exception:  # noqa: BLE001
    paths = None

try:
    import vdesktop
except Exception:  # noqa: BLE001
    vdesktop = None

if paths is not None:
    BASE = paths.BASE
else:
    BASE = os.path.dirname(os.path.abspath(__file__))

DEFAULT_PROFILE = os.path.join(BASE, "kdocs_profile")

# 与统一任务工作台保持一致：优先识别明确的 Excel 导出项，再识别泛化的“导出/下载”。
EXCEL_LABELS = (
    '导出为 Excel(.xlsx)', '导出为Excel(.xlsx)',
    '导出为 Excel (.xlsx)', '导出为Excel（.xlsx）',
    'Excel (.xlsx)', 'Excel（.xlsx）', 'Excel文件(.xlsx)',
    'Microsoft Excel (.xlsx)', 'Excel', '.xlsx', 'XLSX',
    '导出为Excel', '导出为Excel文件', '另存为Excel文件', '下载为Excel',
)
EXPORT_LABELS = ('导出为', '下载为', '导出', '下载到本地', '下载', '另存为')

CONTROL_SELECTOR = 'button,[role="button"],[role="menuitem"],[title],[aria-label],[data-tooltip]'
INVENTORY_SELECTOR = ('button,[role="button"],[role="menuitem"],[role="option"],[title],[aria-label],'
                      '[data-tooltip],[onclick],[tabindex]:not([tabindex="-1"])')
DIALOG_SELECTOR = '[role="dialog"],dialog[open],[aria-modal="true"],.ant-modal,.el-dialog,.dialog'

FREEZE_NOTICE = ('当前工作表的冻结线在显示区域之外，无法滚动浏览。若要浏览全部内容，'
                 '请临时取消冻结窗格，此操作不会影响其他人。')


def _label(text):
    text = re.sub(r'\s+', ' ', text or '').strip()
    text = re.sub(r'\s*\([A-Za-z]\)\s*$', '', text)
    return re.sub(r'(?:\.\.\.|…|：|:)\s*$', '', text).strip().casefold()


class _Ui:
    """精简版 OnlineFormatter：只保留下载所需的控件定位 / 点击 / 对话框能力。"""

    def __init__(self, page, progress=lambda message: None, diagnostic_dir=None):
        self.page = page
        self.progress = progress
        self.diagnostic_dir = diagnostic_dir

    def _frames(self):
        for frame in self.page.frames:
            host = urllib.parse.urlparse(frame.url).hostname
            if host is None or host in ('kdocs.cn', 'wps.cn') or host.endswith(('.kdocs.cn', '.wps.cn')):
                yield frame

    def _pause(self):
        self.page.wait_for_timeout(100)

    def _controls(self, labels, scope=None, *, role=None, menu_only=False):
        wanted = {_label(value) for value in labels}
        roots = [(None, scope)] if scope is not None else [(frame, frame) for frame in self._frames()]
        found = []
        seen = set()
        for frame, root in roots:
            locator = root.locator(CONTROL_SELECTOR)
            try:
                matches = locator.evaluate_all('''(elements, wanted) => {
                    const normalize = text => (text||'').replace(/\\s+/g,' ').trim()
                        .replace(/\\s*\\([A-Za-z]\\)\\s*$/,'').replace(/(?:\\.\\.\\.|…|：|:)\\s*$/,'').trim().toLowerCase();
                    return elements.flatMap((el,index) => {
                        const control = el.closest('button,[role="button"],[role="menuitem"]') || el;
                        const labels = [el.title,el.getAttribute('aria-label'),el.getAttribute('data-tooltip')];
                        if (el.tagName === 'BUTTON' || ['button','menuitem'].includes(el.getAttribute('role')))
                            labels.push(el.innerText);
                        if(!labels.some(label=>wanted.includes(normalize(label))) || !el.getClientRects().length)
                            return [];
                        const style=getComputedStyle(el);
                        if(style.visibility==='hidden'||style.display==='none')return [];
                        let path = [], current = control;
                        while (current && current.nodeType === 1) {
                            let index=1, prev=current.previousElementSibling;
                            while(prev){if(prev.tagName===current.tagName)index++;prev=prev.previousElementSibling;}
                            path.unshift(current.tagName.toLowerCase()+':nth-of-type('+index+')');
                            current=current.parentElement;
                        }
                        return [{index, path:path.join(' > '), role:control.getAttribute('role'),
                                popup:control.getAttribute('aria-haspopup'),
                                disabled:control.disabled || control.getAttribute('aria-disabled')==='true'}];
                    });
                }''', list(wanted))
            except Exception:  # noqa: BLE001
                continue
            for info in matches:
                item = locator.nth(info['index'])
                if not item.is_visible():
                    continue
                if role and info['role'] != role:
                    continue
                if menu_only and info['popup'] not in ('true', 'menu'):
                    continue
                key = (id(frame) if frame is not None else id(root), info['path'])
                if key in seen:
                    continue
                seen.add(key)
                found.append((item, info))
        return found

    def _unique_control(self, labels, scope=None, *, optional=False, role=None, menu_only=False):
        found = self._controls(labels, scope, role=role, menu_only=menu_only)
        if not found and optional:
            return None
        if len(found) != 1:
            exc = ValueError('无法唯一定位“%s”控件；当前网页界面需适配后才能继续' % labels[0])
            exc.diagnostic_context = {'wanted': list(labels), 'matched_controls': len(found),
                                      'visible_controls': self._control_inventory()}
            raise exc
        item, info = found[0]
        if info['disabled'] or not item.is_enabled():
            exc = ValueError('“%s”控件不可用，请确认拥有下载/导出权限' % labels[0])
            exc.diagnostic_context = {'wanted': list(labels), 'visible_controls': self._control_inventory()}
            raise exc
        return item

    def _click(self, item, label):
        try:
            item.click(timeout=3000)
        except Exception as exc:  # noqa: BLE001
            raise ValueError('无法操作“%s”控件，已停止' % label) from exc
        self._pause()

    def _visible_dialogs(self):
        dialogs = []
        for frame in self._frames():
            for item in frame.locator(DIALOG_SELECTOR).all():
                if item.is_visible():
                    nested = item.evaluate('''el => {
                        let p=el.parentElement;
                        while(p){if(p.matches('[role="dialog"],dialog[open],[aria-modal="true"],.ant-modal,.el-dialog,.dialog')
                            && p.getClientRects().length)return true;p=p.parentElement;}return false;
                    }''')
                    if not nested:
                        dialogs.append(item)
        return dialogs

    def _control_inventory(self):
        items = []
        for frame in self._frames():
            try:
                info = frame.locator(INVENTORY_SELECTOR).evaluate_all('''(elements) => {
                    const out=[]; const seen=new Set();
                    for(let i=0;i<elements.length && out.length<60;i++){
                        const el=elements[i];
                        if(el.closest('[role="grid"],[role="gridcell"],td,th'))continue;
                        const r=el.getBoundingClientRect();
                        if(!r.width||!r.height)continue;
                        const control=el.closest('button,[role="button"],[role="menuitem"]')||el;
                        const labels=[el.title,el.getAttribute('aria-label'),el.getAttribute('data-tooltip')];
                        const interactive=el.tagName==='BUTTON'||['button','menuitem','option'].includes(el.getAttribute('role'))
                            ||el.hasAttribute('onclick')||el.hasAttribute('tabindex');
                        if(interactive)labels.push(el.innerText);
                        const text=labels.filter(Boolean).join(' ').replace(/\\s+/g,' ').trim().slice(0,60);
                        if(!text)continue;
                        const key=text+'|'+(control.getAttribute('role')||'');
                        if(seen.has(key))continue; seen.add(key);
                        out.push({tag:control.tagName,role:control.getAttribute('role'),label:text});
                    }
                    return out;
                }''')
                items.extend(info)
            except Exception:  # noqa: BLE001
                continue
        return items[:60]

    def _capture_ui_dump(self):
        if self.diagnostic_dir is None:
            return None
        try:
            folder = Path(self.diagnostic_dir)
            folder.mkdir(exist_ok=True)
            stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            path = folder / f'在线表界面_{stamp}.png'
            self.page.screenshot(path=str(path))
            return str(path)
        except Exception:  # noqa: BLE001
            return None


def _main_menu(ui):
    for label in ('文件', 'File', '文件菜单', '主菜单', '菜单', 'Menu'):
        control = _control(ui, label)
        if control is not None:
            return control
    # 紧凑 WPS 皮肤用无文字汉堡按钮代替“文件”。
    found = []
    for frame in ui._frames():
        paths_ = frame.locator('[class*="menu" i],[class*="hamburger" i],'
                               '[aria-haspopup="menu"]').evaluate_all(r'''elements => {
            const targets = new Set();
            for (const el of elements) {
                const node = el.closest('button,[role="button"],[aria-haspopup="menu"]') || el;
                if (node.closest('[role="grid"],[role="gridcell"],td,th,[role="dialog"],[role="menu"]')) continue;
                const box = node.getBoundingClientRect(), style = getComputedStyle(node);
                if (style.visibility !== 'visible' || style.display === 'none' ||
                    box.x < 0 || box.x > 100 || box.y < 0 || box.y > 100 ||
                    box.width < 12 || box.width > 64 || box.height < 12 || box.height > 64) continue;
                if (!['', '☰', '≡'].includes((node.innerText || '').trim())) continue;
                targets.add(node);
            }
            return [...targets].filter(node => ![...targets].some(other => other !== node && other.contains(node)))
                .map(node => {
                    const parts = [];
                    for (let el = node; el && el.nodeType === 1; el = el.parentElement) {
                        let i = 1;
                        for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling)
                            if (prev.tagName === el.tagName) i++;
                        parts.unshift(el.tagName.toLowerCase() + ':nth-of-type(' + i + ')');
                    }
                    return parts.join(' > ');
                });
        }''')
        found.extend(frame.locator(path) for path in paths_)
    if len(found) > 1:
        raise ValueError('在线表左上角文件菜单不唯一，已停止下载')
    if found and not found[0].is_enabled():
        raise ValueError('在线表左上角文件菜单不可用，请检查权限')
    return found[0] if found else None


def _control(ui, label, scope=None):
    item = ui._unique_control((label,), scope=scope, optional=True)
    if item is not None:
        return item
    found = []
    for root in ([scope] if scope is not None else ui._frames()):
        for candidate in root.get_by_text(label, exact=True).all():
            if candidate.is_visible() and candidate.evaluate('''el =>
                !el.closest('[role="grid"],[role="gridcell"],td,th') &&
                !!el.closest('button,[role="button"],[role="menu"],[role="menuitem"],'
                    + '[class*="menu" i],[role="dialog"]')'''):
                found.append(candidate)
    if len(found) > 1:
        raise ValueError('在线表下载菜单中“%s”不唯一，已停止下载' % label)
    if found:
        if not found[0].is_enabled():
            raise ValueError('在线表的“%s”不可用，请检查下载权限' % label)
        return found[0]
    return None


def _expand_export(ui, control, label):
    # WPS 的“导出为”是悬停子菜单；点击可能反而收起它。
    control.hover(timeout=3000)
    ui.page.wait_for_timeout(250)
    if any(_control(ui, excel) is not None for excel in EXCEL_LABELS):
        return
    ui._click(control, label)


def _dismiss_freeze_notice(ui):
    for frame in ui._frames():
        notice = frame.get_by_text(FREEZE_NOTICE, exact=True)
        ignore = frame.get_by_role('button', name='忽略', exact=True)
        if notice.count() == 1 and notice.is_visible() and ignore.count() == 1 and ignore.is_visible():
            ignore.click(timeout=3000)
            notice.wait_for(state='hidden', timeout=3000)
            return True
    return False


def request_download(page, progress, *, timeout=240, source_label='在线表', diagnostic_dir=None):
    """通过正常导出菜单下载，只接受真实的浏览器 Download 事件。返回 Playwright Download。"""
    ui = _Ui(page, progress, diagnostic_dir=diagnostic_dir)
    downloads = []

    def listener(download):
        downloads.append(download)

    page.on('download', listener)
    try:
        opened = exported = selected = confirmed = False
        deadline = time.monotonic() + timeout
        while not downloads:
            if time.monotonic() >= deadline:
                exc = ValueError(
                    '%s下载超时：未收到 Excel 文件。请先点「打开链接（登录）」登录 kdocs、'
                    '确认拥有下载权限且能展开“导出为 Excel”菜单后重试' % source_label)
                exc.diagnostic_context = {'visible_controls': ui._control_inventory(),
                                          'screenshot': ui._capture_ui_dump(),
                                          'opened': opened, 'exported': exported,
                                          'selected': selected, 'confirmed': confirmed}
                raise exc
            _dismiss_freeze_notice(ui)
            control = None
            expand = False
            if not selected:
                for label in EXCEL_LABELS:
                    control = _control(ui, label)
                    if control is not None:
                        selected = True
                        break
            if control is None and not exported and not selected:
                for label in EXPORT_LABELS:
                    control = _control(ui, label)
                    if control is not None:
                        exported = True
                        expand = label in ('导出为', '下载为', '导出', '另存为')
                        break
            if control is None and not opened and not exported and not selected:
                control = _main_menu(ui)
                if control is not None:
                    opened = True
                    label = '文件菜单'
            if control is None and (selected or exported) and not confirmed:
                dialogs = ui._visible_dialogs()
                if len(dialogs) == 1:
                    for label in ('下载', '导出', '确定'):
                        control = _control(ui, label, scope=dialogs[0])
                        if control is not None:
                            confirmed = True
                            break
            if control is not None:
                progress('正在下载%s：%s…' % (source_label, label))
                if expand:
                    _expand_export(ui, control, label)
                else:
                    try:
                        ui._click(control, label)
                    except ValueError:
                        if not _dismiss_freeze_notice(ui):
                            raise
                        opened = exported = selected = confirmed = False
            else:
                page.wait_for_timeout(250)
        return downloads[0]
    finally:
        page.remove_listener('download', listener)


def save_download(download, out_dir, *, progress=lambda message: None):
    """校验下载的 xlsx 后原子落盘，返回 {'path', 'sheet', 'size'}。"""
    from openpyxl import load_workbook
    folder = Path(out_dir)
    folder.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    target = folder / ('在线表_%s.xlsx' % stamp)
    fd, filename = tempfile.mkstemp(prefix='.kdocs_download_', suffix='.xlsx', dir=folder)
    os.close(fd)
    temporary = Path(filename)
    try:
        failure = download.failure()
        if failure:
            raise ValueError('在线表下载失败：' + str(failure))
        if Path(download.suggested_filename).suffix.lower() != '.xlsx':
            raise ValueError('下载结果不是 .xlsx 文件，请在“导出为”菜单里选择 Excel 格式')
        progress('正在保存并校验下载的表格…')
        download.save_as(str(temporary))
        sheets = []
        try:
            book = load_workbook(str(temporary), read_only=True)
            try:
                sheets = list(book.sheetnames)
            finally:
                book.close()
        except Exception as exc:  # noqa: BLE001
            raise ValueError('下载的不是有效的 Excel 工作簿：' + str(exc)) from exc
        if not sheets:
            raise ValueError('下载的 Excel 里没有任何工作表')
        os.replace(temporary, target)
        return dict(path=str(target), sheet=sheets[0], sheets=sheets,
                    size=target.stat().st_size)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


def download_kdocs(url, out_dir, *, sheet_name='', progress=print, timeout=240,
                   profile_dir=None, should_stop=None):
    """打开 kdocs/WPS 分享链接，用已登录 Edge 下载成 xlsx 落到 out_dir。

    返回 {'path', 'sheet', 'sheets', 'size'}；sheet 为解析出的工作表名。
    """
    try:
        from playwright.sync_api import sync_playwright
    except ModuleNotFoundError as exc:
        raise ValueError('缺少 playwright 组件，请先运行：pip install playwright（使用系统 Edge，'
                         '无需 playwright install 下载浏览器）') from exc

    url = (url or '').strip()
    if not url.lower().startswith(('http://', 'https://')):
        raise ValueError('在线表链接无效')

    profile = Path(profile_dir or DEFAULT_PROFILE)
    profile.parent.mkdir(parents=True, exist_ok=True)

    pw = sync_playwright().start()
    context = None
    switched = False
    back = False
    try:
        # 把下载专用 Edge 放到另一个虚拟桌面：先切过去启动，再立刻切回来。
        if vdesktop is not None:
            try:
                vdesktop.switch_away()
                switched = True
            except Exception:  # noqa: BLE001
                switched = False
        try:
            context = pw.chromium.launch_persistent_context(
                str(profile), channel='msedge', headless=False,
                accept_downloads=True, no_viewport=True)
        except Exception as exc:  # noqa: BLE001
            raise ValueError('无法启动 Microsoft Edge，请确认已安装 Edge，'
                             '并关闭之前打开的下载专用 Edge 后重试') from exc

        page = context.new_page()
        if switched:
            try:
                vdesktop.switch_back()
                back = True
            except Exception:  # noqa: BLE001
                back = False
        progress('正在打开在线表：%s' % url[:80])
        page.goto(url, wait_until='domcontentloaded', timeout=60000)
        progress('下载窗口已放到另一个虚拟桌面；程序正在等待“导出为 Excel”菜单…')
        diagnostic_dir = Path(profile).parent / 'error_reports'
        download = request_download(page, progress, timeout=timeout,
                                    source_label='在线表', diagnostic_dir=str(diagnostic_dir))
        info = save_download(download, out_dir, progress=progress)
        progress('在线表已下载：%s（%.1f KB）' % (info['path'], info['size'] / 1024.0))
        if sheet_name and sheet_name.strip():
            if sheet_name.strip() not in info['sheets']:
                raise ValueError('下载的在线表里没有工作表“%s”，现有：%r'
                                 % (sheet_name.strip(), info['sheets']))
            info['sheet'] = sheet_name.strip()
        return info
    finally:
        # 若中途出错（例如 Edge 启动失败），也要切回原桌面。
        if switched and not back:
            try:
                vdesktop.switch_back()
            except Exception:  # noqa: BLE001
                pass
        try:
            if context is not None:
                context.close()
        except Exception:  # noqa: BLE001
            pass
        try:
            pw.stop()
        except Exception:  # noqa: BLE001
            pass


def _find_edge():
    import shutil
    exe = shutil.which('msedge') or shutil.which('msedge.exe')
    if exe:
        return exe
    for root in (os.environ.get('ProgramFiles(x86)'), os.environ.get('ProgramFiles'),
                 os.environ.get('LOCALAPPDATA')):
        if not root:
            continue
        path = os.path.join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe')
        if os.path.exists(path):
            return path
    return None


def open_kdocs_login(url, *, profile_dir=None):
    """用独立 Edge 应用窗口打开在线表（与下载共用同一 profile），方便首次登录。

    与 download_kdocs 共用同一 user-data-dir，登录态互通；下载前请先关闭这个窗口。
    """
    import subprocess

    url = (url or '').strip()
    if not url.lower().startswith(('http://', 'https://')):
        raise ValueError('在线表链接无效')
    edge = _find_edge()
    if not edge:
        raise ValueError('未找到 Microsoft Edge，请先安装 Edge')
    profile = Path(profile_dir or DEFAULT_PROFILE)
    profile.parent.mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen([edge, '--app=' + url, '--user-data-dir=' + str(profile)])
    return proc is not None
