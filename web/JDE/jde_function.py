#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JDE 页面操作：登录、R55GL1005 导出、利润表（R55GZ001）导出、日记帐分录录入。

本模块复用房产税流程已验证的 JDE 登录、报表下载与分录录入实现；
利润表导出按 R55GZ001 / TKHXLS001 的批版本处理流程实现（页签按候选文字匹配，
处理选项 Century / Fiscal Year / Period Number 按会计期间自动填写）。
"""

import functools
import re
import time
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from utils.helpers import ensure_dir

MENU_IFRAME = 'iframe[name="e1menuAppIframe"]'
# 日记帐分录录入前需要选中的分录方案（JDE #gtab0_1 下拉值）。
JOURNAL_SCHEME_OPTION = "8d5afde055f149ea973fb07587dafa28"
# 粘贴前必须存在的列及顺序（TSV 按该顺序整块粘贴）。
JOURNAL_COLUMNS = ("帐号", "借项金额", "贷项金额", "备注")

OUTPUT_LINK_NAME = "可单击图标 查看输出"
JOB_FILTER_NAME = "作业"
JOB_NAME = "R55GL1005_TKHXLS001"
JOB_DONE_KEYWORDS = ("完成", "Done")
JOB_FAILED_KEYWORDS = ("错误", "失败", "Error")
JOB_STATUS_POLL_SECONDS = 120
JOB_STATUS_POLL_INTERVAL_SECONDS = 5

# 利润表（Profit & Loss (Yearly)）默认参数，可由配置覆盖。
INCOME_STATEMENT_MENU = "中国政府报表"
# 页签在不同环境带公司前缀，按候选顺序依次匹配。
INCOME_STATEMENT_MENU_CANDIDATES = ("中国政府报表", "广州中国政府报表", "太古汇中国政府报表")
INCOME_STATEMENT_VERSION = "TKHXLS001"
INCOME_STATEMENT_JOB = "R55GZ001"
# 处理选项字段：Century（年份前两位，如 2026→20）、Fiscal Year（年份后两位，如 26）、
# Period Number（月）。
INCOME_STATEMENT_OPTION_FIELDS = ("E3", "E4", "E28")


# 重试
def retry_on_failure(max_retries=3, delay=5):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception:
                    if attempt < max_retries - 1:
                        logger.warning(f"{func.__name__} 执行失败，准备重试")
                        logger.info(f"等待 {delay} 秒后进行第 {attempt + 2} 次尝试...")
                        time.sleep(delay)
                    else:
                        logger.info(f"{func.__name__} 最终执行失败，已重试 {max_retries} 次")
                        raise
            return None
        return wrapper
    return decorator


# 登录 JDE
@retry_on_failure(max_retries=3, delay=5)
def jde_login(in_page, jde_url, jde_environment, jde_role, jde_user, jde_pass):
    in_page.goto(jde_url)

    element_exists = in_page.locator("#drop_mainmenu")
    if element_exists.is_visible(timeout=5000):
        logger.info("登录成功")
    else:
        element_exists = in_page.get_by_role("textbox", name="Environment")
        if not element_exists.is_visible(timeout=2000):
            in_page.get_by_role("link", name="Details").click()
            in_page.wait_for_timeout(500)

        if jde_environment:
            in_page.get_by_role("textbox", name="Environment").fill(jde_environment)
            in_page.wait_for_timeout(500)
        if jde_role:
            in_page.get_by_role("textbox", name="Role").fill(jde_role)
            in_page.wait_for_timeout(500)

        in_page.get_by_role("textbox", name="User ID").fill(jde_user)
        in_page.wait_for_timeout(500)
        in_page.get_by_role("textbox", name="Password").fill(jde_pass)
        in_page.wait_for_timeout(500)

        in_page.get_by_role("button", name="Sign In").click()
        in_page.wait_for_timeout(3000)

        element_exists = in_page.locator("#drop_mainmenu")
        element_exists.wait_for(state="visible", timeout=20000)
        if element_exists.is_visible():
            logger.info("登录成功")
        else:
            raise Exception("登录失败，未进入JDE界面，请检查！")


def _enter_total_ledger(frame, timeout):
    """先进入所有页面，再进入总账页签（页签在 iframe 内点击）。"""
    select_page = frame.get_by_role("button", name="选择页面")
    select_page.wait_for(state="visible", timeout=timeout)
    select_page.click()
    all_pages_grid = frame.get_by_role("gridcell", name="所有页面")
    all_pages_grid.wait_for(state="visible", timeout=timeout)
    all_pages_grid.click()

    all_pages = frame.get_by_text("所有页面").first
    all_pages.click()
    if all_pages.is_visible():
        select_page.click()

    tab = frame.locator("#tabBar").get_by_text("总帐")
    tab.wait_for(state="visible", timeout=timeout)
    tab.click()


def _click_menu_link(in_page, name, timeout):
    """在主文档和当前 iframe 中查找可见菜单项并点击。"""
    deadline = time.monotonic() + timeout / 1000
    while True:
        for frame in in_page.frames:
            links = frame.get_by_role("link", name=name)
            for index in range(links.count()):
                link = links.nth(index)
                if link.is_visible():
                    logger.info(f"在页面/frame中找到菜单项并点击：{name}")
                    link.click(timeout=timeout)
                    return

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise PlaywrightTimeoutError(
                f"等待 {timeout} 毫秒后，主文档和 iframe 中均未找到可见菜单项：{name}"
            )
        in_page.wait_for_timeout(min(500, max(1, int(remaining * 1000))))


def _click_checkbox(frame, name, timeout):
    """JDE 网格复选框由页面脚本维护状态，只做点击。"""
    checkbox = frame.get_by_role("checkbox", name=name).first
    checkbox.wait_for(state="visible", timeout=timeout)
    checkbox.click(timeout=timeout)
    logger.info(f"已点击勾选框：{name}")


def _set_clipboard_text(text):
    """把整块 TSV 写入系统剪贴板，供 JDE 网格 Ctrl+V 粘贴。"""
    try:
        import win32clipboard
    except ImportError as exc:
        raise RuntimeError(f"缺少 pywin32，无法写入剪贴板：{exc}") from exc
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def _focus_journal_grid(frame, in_page, timeout):
    """把焦点移到分录网格首格，确认不了就报错，绝不盲目粘贴。"""
    focused_selector = "tr[id] input.JSTextfield:focus"
    for selector in ("tr[id^='G0_'] input.JSTextfield", "#G0_1 input.JSTextfield"):
        cells = frame.locator(selector)
        if cells.count() == 0:
            continue
        cells.first.click(timeout=timeout)
        in_page.wait_for_timeout(300)
        if frame.locator(focused_selector).count() > 0:
            logger.info(f"焦点已移到分录网格：{selector}")
            return
    for _ in range(12):
        in_page.keyboard.press("Tab")
        in_page.wait_for_timeout(200)
        if frame.locator(focused_selector).count() > 0:
            logger.info("已用 Tab 把焦点移入分录网格")
            return
    raise RuntimeError("未能把焦点移到分录网格，已跳过粘贴，避免写入日期栏")


def _normalize_label(text):
    return re.sub(r"\s+", "", str(text or ""))


def _role_header_labels(frame):
    labels = []
    try:
        headers = frame.get_by_role("columnheader")
        for index in range(headers.count()):
            item = headers.nth(index)
            try:
                if not item.is_visible():
                    continue
                label = _normalize_label(item.inner_text())
            except Exception:
                continue
            if label:
                labels.append(label)
    except Exception:
        return []
    return labels


def _th_header_labels(frame):
    labels = []
    try:
        headers = frame.locator("th")
        for index in range(headers.count()):
            item = headers.nth(index)
            try:
                if not item.is_visible():
                    continue
                label = _normalize_label(item.inner_text())
            except Exception:
                continue
            if label:
                labels.append(label)
    except Exception:
        return []
    return labels


def _position_header_labels(frame):
    """按表头文字与网格首行的相对位置读取列顺序。"""
    inputs = frame.locator("#G0_1 input.JSTextfield")
    if inputs.count() == 0:
        inputs = frame.locator("tr[id^='G0_'] input.JSTextfield")
    if inputs.count() == 0:
        return []
    try:
        anchor = inputs.first.bounding_box()
    except Exception:
        return []
    if not anchor:
        return []
    boxes = []
    for name in JOURNAL_COLUMNS:
        locator = frame.get_by_text(name)
        found = []
        for index in range(min(locator.count(), 20)):
            item = locator.nth(index)
            try:
                if not item.is_visible():
                    continue
                box = item.bounding_box()
            except Exception:
                continue
            if not box or box["width"] <= 0 or box["height"] <= 0:
                continue
            if box["y"] + box["height"] <= anchor["y"] + 2:
                found.append(box)
        if not found:
            return []
        found.sort(key=lambda box: box["y"] + box["height"])
        boxes.append(found[-1])
    tops = [box["y"] for box in boxes]
    if max(tops) - min(tops) > 40:
        return []
    return [
        name for name, _ in sorted(zip(JOURNAL_COLUMNS, boxes), key=lambda pair: pair[1]["x"])
    ]


def _journal_header_labels(frame):
    for labels in (_role_header_labels(frame), _th_header_labels(frame),
                   _position_header_labels(frame)):
        if labels and all(name in labels for name in JOURNAL_COLUMNS):
            return labels
    return []


def _verify_journal_columns(frame, timeout):
    """粘贴前校验网格列顺序，不通过则中止。"""
    wait_timeout = min(timeout, 5000)
    try:
        frame.locator("#G0_1 input.JSTextfield").first.wait_for(
            state="visible", timeout=wait_timeout
        )
    except Exception:
        try:
            frame.locator("tr[id^='G0_'] input.JSTextfield").first.wait_for(
                state="visible", timeout=wait_timeout
            )
        except Exception:
            pass
    labels = _journal_header_labels(frame)
    if not labels:
        raise RuntimeError("无法读取日记帐分录网格表头，未执行粘贴")
    positions = {}
    for index, label in enumerate(labels):
        positions.setdefault(label, index)
    missing = [name for name in JOURNAL_COLUMNS if name not in positions]
    if missing:
        raise RuntimeError(
            f"分录网格缺少列：{'、'.join(missing)}；当前表头：{' / '.join(labels)}"
        )
    order = [name for name in sorted(positions, key=positions.get)]
    if tuple(order) != JOURNAL_COLUMNS:
        raise RuntimeError(
            f"分录网格列顺序不符：期望 {'、'.join(JOURNAL_COLUMNS)}，"
            f"实际 {'、'.join(order)}；当前表头：{' / '.join(labels)}"
        )
    logger.info(f"分录网格列顺序校验通过：{' / '.join(labels)}")


def _click_cancel(frame, timeout):
    try:
        button = frame.get_by_role("button", name="取消 (Ctrl+Alt+L)").first
        if button.count() == 0 or not button.is_visible():
            return False
        button.click(timeout=timeout)
        return True
    except Exception as exc:
        logger.warning(f"点击取消失败：{exc}")
        return False


def _click_cancel_anywhere(in_page, timeout, prefer_new=False):
    pages = list(in_page.context.pages)
    if prefer_new:
        pages = [page for page in pages if page != in_page] + [in_page]
    for page in pages:
        try:
            frames = page.frames
        except Exception:
            continue
        for frame in frames:
            try:
                buttons = frame.get_by_role("button", name="取消 (Ctrl+Alt+L)")
                for index in range(min(buttons.count(), 5)):
                    button = buttons.nth(index)
                    if button.is_visible():
                        button.click(timeout=timeout)
                        return True
            except Exception:
                continue
    return False


def _discard_entry(frame, in_page, timeout):
    if not _click_cancel(frame, timeout):
        _click_cancel_anywhere(in_page, timeout)
    in_page.wait_for_timeout(800)
    if _click_cancel_anywhere(in_page, min(timeout, 5000), prefer_new=True):
        logger.info("已点击第二次取消")
    else:
        logger.info("未发现第二次取消按钮，跳过")


def _save_entry(frame, in_page, timeout):
    """保存分录；JDE 保存按钮文字在不同版本可能不同，按候选文字依次尝试。"""
    candidates = ("保存 (Ctrl+Alt+S)", "保存")
    for name in candidates:
        try:
            button = frame.get_by_role("button", name=name).first
            if button.count() and button.is_visible():
                button.click(timeout=timeout)
                in_page.wait_for_timeout(1000)
                logger.info("已点击保存分录")
                return True
        except Exception:
            continue
    raise RuntimeError("未找到分录保存按钮，分录未保存")


def create_journal_entry(in_page, company, description, gl_date, rows,
                         scheme_option=JOURNAL_SCHEME_OPTION,
                         timeout=60000, discard=True):
    """录入日记帐分录并返回 JDE 生成的批号。

    rows 为 (帐号, 借项金额, 贷项金额, 备注) 序列，按 TSV 整块粘贴。
    discard=True 时沿用房产税流程的联调行为：粘贴后取消退出、不保存。
    """
    if not rows:
        raise ValueError("分录没有可录入的数据行")
    home = in_page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()
    logger.info("已返回主页，重新选择页面并进入总帐")
    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    logger.info("打开日记帐分录")
    _click_menu_link(in_page, "日记帐分录", timeout)

    frame.get_by_role("button", name="增加 (Ctrl+Alt+A)").click(timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    batch_box = frame.get_by_role("textbox", name="批号")
    batch_box.wait_for(state="visible", timeout=timeout)
    batch_box.click(timeout=timeout)
    batch = batch_box.input_value(timeout=timeout)
    logger.info(f"JDE 批号：{batch}")

    frame.get_by_role("textbox", name="Document Company").fill(
        str(company).strip(), timeout=timeout
    )
    frame.get_by_role("textbox", name="说明").fill(
        str(description).strip(), timeout=timeout
    )
    frame.get_by_role("textbox", name="总帐日期").fill(
        str(gl_date).strip(), timeout=timeout
    )

    lines = []
    for code, debit, credit, remark in rows:
        lines.append("\t".join("" if value is None else _cell_text(value)
                               for value in (code, debit, credit, remark)))
    _set_clipboard_text("\n".join(lines))
    logger.info(f"分录数据已写入剪贴板：{len(lines)} 行")

    scheme = frame.locator("#gtab0_1")
    scheme.wait_for(state="visible", timeout=timeout)
    scheme.select_option(label="RPA", timeout=timeout)
    logger.info("已选择分录方案：RPA")
    in_page.wait_for_timeout(1000)

    try:
        _verify_journal_columns(frame, timeout)
    except Exception:
        _discard_entry(frame, in_page, timeout)
        raise
    _focus_journal_grid(frame, in_page, timeout)
    in_page.keyboard.press("Control+V")
    in_page.wait_for_timeout(1000)

    if discard:
        _discard_entry(frame, in_page, timeout)
        logger.info("已取消分录录入（未保存）")
    else:
        _save_entry(frame, in_page, timeout)
    return batch


def _cell_text(value):
    """数值去掉多余小数位，避免粘贴时出现科学计数法。"""
    if isinstance(value, float):
        return format(value, "f").rstrip("0").rstrip(".")
    return str(value)


# ---------------------------------------------------------------------------
# R55GL1005 导出
# ---------------------------------------------------------------------------

def download_gl_movement(in_page, ledger_type, from_company, to_company,
                         start_date, end_date, destination, timeout=60000,
                         grid_timeout=300000, download_timeout=300000):
    """导出明细账（预留；企业所得税流程当前只导出 R55GL1005 与利润表）。"""
    for label, value in (("分类帐类型", ledger_type), ("自公司", from_company), ("至公司", to_company),
                         ("起始日期", start_date), ("至日期", end_date)):
        if value is None or not str(value).strip():
            raise ValueError(f"明细账缺少导出参数：{label}")

    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    logger.info("打开按科目帐的帐户查询")
    _click_menu_link(in_page, "按科目帐的帐户查询", timeout)

    def textbox(name, exact=False):
        locator = frame.get_by_role("textbox", name=name, exact=exact)
        locator.wait_for(state="visible", timeout=timeout)
        return locator

    logger.info(f"填写查询条件：{from_company}-{to_company}{start_date}-{end_date}类型 {ledger_type}")
    textbox("分类帐类型 2", exact=True).fill(str(ledger_type).strip())
    textbox("自公司").fill(str(from_company).strip())
    textbox("至公司").fill(str(to_company).strip())
    textbox("起始日期").fill(str(start_date).strip())
    textbox("至日期").fill(str(end_date).strip())

    frame.get_by_role("button", name="查找 (Ctrl+Alt+I)").click()

    logger.info("查询已提交，跳转至提取记录末尾并等待网格加载")
    frame.get_by_role("button", name="转到结束").click(timeout=timeout)
    select_all = frame.get_by_role("checkbox", name="选择所有提取的记录")
    select_all.wait_for(state="visible", timeout=grid_timeout)
    select_all.check(timeout=grid_timeout)
    logger.info("网格记录已就绪并选择全部记录")
    frame.get_by_role(
        "button", name="输出网格数据 (Ctrl + Shift +E)"
    ).click(timeout=timeout)
    local_file = frame.get_by_role("radio", name="本地文件系统")
    local_file.wait_for(state="visible", timeout=timeout)
    local_file.check()
    frame.get_by_role("radio", name="输出到 Excel").check()
    frame.get_by_role("radio", name="输出到新工作簿 输出到新工作簿").check()
    frame.get_by_role("radio", name="所有").check()

    logger.info(f"已点击继续，等待明细账下载，最长 {download_timeout // 1000} 秒")
    with in_page.expect_download(timeout=download_timeout) as download_info:
        frame.get_by_role(
            "cell", name="继续 (Ctrl+Alt+E)", exact=True
        ).locator("a").click(timeout=download_timeout)
    download = download_info.value

    target = Path(destination)
    if not target.suffix:
        target = target.with_suffix(Path(download.suggested_filename).suffix or ".xlsx")
    ensure_dir(target.parent)
    download.save_as(str(target))
    logger.info(f"明细账已导出：{target}")
    return target


def _click_and_collect_download(in_page, link, timeout):
    """点击后可能跳转页面触发下载；下载可能在原页面或新开页面中发起。"""
    context = in_page.context
    existing_pages = list(context.pages)
    collected = []

    def collect(download):
        collected.append(download)

    def watch(page):
        page.on("download", collect)

    for page in existing_pages:
        watch(page)
    context.on("page", watch)
    try:
        link.click(timeout=timeout)
        deadline = time.monotonic() + timeout / 1000
        while not collected:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise PlaywrightTimeoutError(f"点击后 {timeout} 毫秒内未触发下载")
            in_page.wait_for_timeout(min(500, max(1, int(remaining * 1000))))
        return collected[0]
    finally:
        context.remove_listener("page", watch)
        for page in context.pages:
            page.remove_listener("download", collect)


def _close_pages_opened_after(context, existing_pages):
    for page in context.pages:
        if page not in existing_pages:
            page.close()


def _row_text(link):
    try:
        return " ".join(
            link.locator("xpath=ancestor::tr[1]").inner_text(timeout=1000).split()
        )
    except Exception:
        return ""


def _job_number(row_text):
    match = re.search(r"R55\w*\s+(\d{5,10})\b", row_text)
    if match is None:
        match = re.search(r"\b(\d{5,10})\b", row_text)
    return match.group(1) if match else None


def _requery_submitted_jobs(frame, job_name, timeout=1500):
    """在“作业”框重新填入作业名并回车，让列表重新查询以更新状态。"""
    try:
        job_filter = frame.get_by_role("textbox", name=JOB_FILTER_NAME, exact=True)
        job_filter.fill(job_name, timeout=timeout)
        job_filter.press("Enter", timeout=timeout)
        return True
    except Exception as exc:
        logger.warning(f"重新查询作业状态失败：{exc}")
        return False


def _wait_latest_job_completed(frame, timeout,
                               poll_seconds=JOB_STATUS_POLL_SECONDS,
                               job_name=JOB_NAME,
                               link_name=OUTPUT_LINK_NAME):
    """等“已提交作业”最新一行变为完成，返回该行可下载的链接。"""
    links = frame.get_by_role("link", name=link_name)
    deadline = time.monotonic() + timeout / 1000
    while links.count() == 0:
        if time.monotonic() >= deadline:
            raise PlaywrightTimeoutError(f"{timeout} 毫秒内“已提交作业”列表仍为空白")
        time.sleep(0.5)

    first_row = _row_text(links.first)
    job_no = _job_number(first_row)
    if job_no is None:
        raise RuntimeError(f"无法从最新作业行解析作业号：{first_row!r}")
    logger.info(f"本次提交的作业号：{job_no}；当前状态：{first_row}")

    poll_deadline = time.monotonic() + poll_seconds
    last_row = first_row
    requery = False
    while True:
        if requery:
            _requery_submitted_jobs(frame, job_name)

        target = None
        for index in range(links.count()):
            link = links.nth(index)
            row = _row_text(link)
            if job_no in row:
                last_row = row
                target = link
                break

        if target is not None and any(k in last_row for k in JOB_DONE_KEYWORDS):
            logger.info(f"作业 {job_no} 已完成：{last_row}")
            return target

        if any(k in last_row for k in JOB_FAILED_KEYWORDS):
            raise RuntimeError(f"作业 {job_no} 状态异常：{last_row}")

        if time.monotonic() >= poll_deadline:
            raise PlaywrightTimeoutError(
                f"作业 {job_no} 在 {poll_seconds} 秒内未完成，最后状态：{last_row}"
            )

        requery = True
        remaining = poll_deadline - time.monotonic()
        time.sleep(min(JOB_STATUS_POLL_INTERVAL_SECONDS, max(0, remaining)))


def _period_parts(period):
    """报表处理选项的世纪、会计年度（两位）与期间编号；未传值按当前自然月。

    四个数字的年份拆成两段：2026 → ("20", "26")，期间编号为月（10）。
    """
    if period is None or not str(period).strip():
        period = date.today()
    if not isinstance(period, (datetime, date)):
        text = str(period).strip()
        match = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
        if not match:
            raise ValueError(
                f"会计期间格式不正确，应为 YYYY-MM 或 YYYY/MM 等可识别格式；当前值：{text}"
            )
        period = date(int(match.group(1)), int(match.group(2)), 1)
    return f"{period.year // 100:02d}", period.strftime("%y"), str(period.month)


def download_r55gl1005(in_page, company, destination, period=None, timeout=60000):
    """返回主页，提交 R55GL1005 报表并下载生成的 Excel 文件。"""
    home = in_page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()
    logger.info("已返回主页，重新选择页面并进入总帐")

    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    _click_menu_link(in_page, "中方科目总账", timeout)

    version = frame.get_by_role("textbox", name="版本", exact=True)
    version.wait_for(state="visible", timeout=timeout)
    version.fill("TKHXLS001", timeout=timeout)
    version.press("Enter", timeout=timeout)
    _click_checkbox(frame, "选择所有提取的记录", timeout)
    frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click(timeout=timeout)
    _click_checkbox(frame, "数据选择", timeout)
    frame.locator("a").filter(has_text="提交(U)").click(timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    selection = frame.locator('input[name="E1"]')
    selection.wait_for(state="visible", timeout=timeout)
    selection.click(timeout=timeout)
    selection.click(timeout=timeout)
    frame.locator("td").filter(has_text=re.compile(r"^Company$")).click(
        timeout=timeout
    )
    company_value = str(company).strip().lstrip("0") or "0"
    selection.fill("", timeout=timeout)
    selection.fill(company_value, timeout=timeout)

    _, fiscal_year, period_number = _period_parts(period)
    logger.info(f"填写会计年度 {fiscal_year}、期间编号 {period_number}")
    for name, value in (("E2", fiscal_year), ("E3", period_number)):
        field = frame.locator(f'input[name="{name}"]')
        field.wait_for(state="visible", timeout=timeout)
        field.click(timeout=timeout)
        field.fill(value, timeout=timeout)

    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    frame.locator("a").filter(has_text="屏幕(F)").click(timeout=timeout)
    frame.get_by_role("cell", name="已提交作业", exact=True).click(timeout=timeout)

    job_filter = frame.get_by_role("textbox", name=JOB_FILTER_NAME, exact=True)
    job_filter.wait_for(state="visible", timeout=timeout)
    job_filter.fill(JOB_NAME, timeout=timeout)
    job_filter.press("Enter", timeout=timeout)
    output_link = _wait_latest_job_completed(frame, timeout, job_name=JOB_NAME)

    target = Path(destination)
    if not target.suffix:
        target = target.with_suffix(".xls")
    ensure_dir(target.parent)
    existing_pages = list(in_page.context.pages)
    download = _click_and_collect_download(in_page, output_link, timeout)
    download.save_as(str(target))
    _close_pages_opened_after(in_page.context, existing_pages)
    logger.info(f"R55GL1005 已导出：{target}")
    return target


# ---------------------------------------------------------------------------
# 利润表（R55GZ001 / Profit & Loss (Yearly)）导出
# ---------------------------------------------------------------------------
# 主页 → 选择页面 → 所有页面 → 总帐 → 中国政府报表（广州…）相关页签 →
# 利润表 → 版本 TKHXLS001 → 选择所有提取的记录 → 选择 → 数据选择 → 提交(U) →
# 确定 → 处理选项（Century / Fiscal Year / Period Number）→ 确定 →
# 屏幕(F) → 已提交作业 → 作业 R55GZ001_TKHXLS001 → 轮询至完成 → 下载。

def _click_tab(frame, names, timeout):
    """在 #tabBar 中按候选文字点击页签（不同环境页签名带公司前缀）。"""
    tab_bar = frame.locator("#tabBar")
    for name in names:
        tab = tab_bar.get_by_text(name).first
        try:
            tab.wait_for(state="visible", timeout=min(timeout, 3000))
        except PlaywrightTimeoutError:
            continue
        tab.click(timeout=timeout)
        logger.info(f"已切换页签：{name}")
        return
    raise PlaywrightTimeoutError(f"#tabBar 中未找到页签：{' / '.join(names)}")


def _click_button(frame, timeout, *names):
    """点击第一个出现的候选按钮；同一按钮在不同表单的快捷键文字可能不同。"""
    for name in names:
        button = frame.get_by_role("button", name=name).first
        try:
            button.wait_for(state="visible", timeout=min(timeout, 3000))
        except PlaywrightTimeoutError:
            continue
        button.click(timeout=timeout)
        logger.info(f"已点击按钮：{name}")
        return
    raise PlaywrightTimeoutError(f"未找到按钮：{' / '.join(names)}")


def _fill_option(frame, name, value, timeout):
    """填写处理选项编号输入框（E3/E4/E28 等）。"""
    field = frame.locator(f'input[name="{name}"]')
    field.wait_for(state="visible", timeout=timeout)
    field.click(timeout=timeout)
    field.fill(str(value), timeout=timeout)


def _fill_company_literal(frame, company, timeout):
    """在利润表数据选择中把公司代码作为字面值填入，避免导出其它公司数据。"""
    code = str(company).strip()
    if not code:
        raise RuntimeError("利润表导出缺少公司代码")
    if code.isdigit():
        code = code.zfill(5)
    operand = frame.locator("#RightOperand1")
    operand.wait_for(state="visible", timeout=timeout)
    operand.select_option("Literal", timeout=timeout)
    field = frame.locator("#LITtf")
    field.click(timeout=timeout)
    field.click(timeout=timeout)
    field.fill("", timeout=timeout)
    field.fill(code, timeout=timeout)
    logger.info(f"利润表数据选择公司代码：{code}")


def download_income_statement(in_page, company, destination, period=None,
                              menu_name=INCOME_STATEMENT_MENU,
                              version=INCOME_STATEMENT_VERSION,
                              job_name=INCOME_STATEMENT_JOB,
                              timeout=60000):
    """返回主页，提交 R55GZ001 利润表并下载生成的文件。"""
    home = in_page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()
    logger.info("已返回主页，重新选择页面并进入总帐")

    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    tab_names = (menu_name,) + tuple(
        name for name in INCOME_STATEMENT_MENU_CANDIDATES if name != menu_name
    )
    _click_tab(frame, tab_names, timeout)

    logger.info("打开利润表")
    _click_menu_link(in_page, "利润表", timeout)

    version_box = frame.get_by_role("textbox", name="版本", exact=True)
    version_box.wait_for(state="visible", timeout=timeout)
    version_box.fill(version, timeout=timeout)
    version_box.press("Enter", timeout=timeout)
    _click_checkbox(frame, "选择所有提取的记录", timeout)
    frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click(timeout=timeout)
    _click_checkbox(frame, "数据选择", timeout)
    frame.locator("a").filter(has_text="提交(U)").click(timeout=timeout)
    _fill_company_literal(frame, company, timeout)
    # 数据选择后连续两个确定：先确认公司条件，再关闭数据选择提示。
    _click_button(frame, timeout, "确定 (Ctrl+Alt+O)", "确定 (Ctrl+Alt+0)")
    _click_button(frame, timeout, "确定 (Ctrl+Alt+O)", "确定 (Ctrl+Alt+0)")

    century, fiscal_year, period_number = _period_parts(period)
    logger.info(
        f"填写处理选项：Century {century}、Fiscal Year {fiscal_year}、"
        f"Period Number {period_number}（公司 {company}）"
    )
    for name, value in zip(
        INCOME_STATEMENT_OPTION_FIELDS, (century, fiscal_year, period_number)
    ):
        _fill_option(frame, name, value, timeout)

    _click_button(frame, timeout, "确定 (Ctrl+Alt+O)", "确定 (Ctrl+Alt+0)")

    frame.locator("a").filter(has_text="屏幕(F)").click(timeout=timeout)
    frame.get_by_role("cell", name="已提交作业", exact=True).click(timeout=timeout)

    submitted_job = f"{job_name}_{version}"
    job_filter = frame.get_by_role("textbox", name=JOB_FILTER_NAME, exact=True)
    job_filter.wait_for(state="visible", timeout=timeout)
    job_filter.fill(submitted_job, timeout=timeout)
    job_filter.press("Enter", timeout=timeout)
    # 与 R55GL1005 一致：定时重新查询，等最新一行变为完成后再下载。
    output_link = _wait_latest_job_completed(
        frame, timeout, job_name=submitted_job, link_name=submitted_job
    )

    target = Path(destination)
    if not target.suffix:
        target = target.with_suffix(".xls")
    ensure_dir(target.parent)
    existing_pages = list(in_page.context.pages)
    download = _click_and_collect_download(in_page, output_link, timeout)
    download.save_as(str(target))
    _close_pages_opened_after(in_page.context, existing_pages)
    logger.info(f"利润表已导出：{target}")
    return target

