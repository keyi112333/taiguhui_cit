#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 通用操作：外部链接保留、公式重算、JDE 旧版 xls 转换、分录区定位。

重算与转换逻辑复用房产税流程，保证与 JDE 导出模板的既有处理一致。
"""
import os
import re
import time
import uuid
import zipfile
from copy import copy
from pathlib import Path
from tempfile import NamedTemporaryFile

from config.config_loader import ConfigError
from log.loggingConfig import logger

EXTERNAL_LINK_DIR = "xl/externalLinks/"
INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")
JOURNAL_TITLES = ("账号", "借项金额", "贷项金额", "备注")


def external_link_parts(path):
    """读取外部链接部件；openpyxl 保存只写回部分关系，须在保存后回写。"""
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()
                if name.startswith(EXTERNAL_LINK_DIR) and not name.endswith("/")}


def restore_external_links(path, parts):
    """把原始外部链接部件写回，避免 rId 悬空。"""
    if not parts:
        return
    path = Path(path)
    temp = None
    try:
        with zipfile.ZipFile(path) as archive:
            entries = [(item, archive.read(item.filename)) for item in archive.infolist()]
        with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as handle:
            temp = Path(handle.name)
        with zipfile.ZipFile(temp, "w") as archive:
            for item, data in entries:
                archive.writestr(item, parts.get(item.filename, data))
        temp.replace(path)
        temp = None
    finally:
        if temp and temp.exists():
            temp.unlink()


def save_with_openpyxl(book, path, links):
    """openpyxl 保存后回写外部链接部件。"""
    temp_path = None
    try:
        with NamedTemporaryFile(dir=Path(path).parent, suffix=".xlsx", delete=False) as tmp:
            temp_path = Path(tmp.name)
        book.save(temp_path)
        restore_external_links(temp_path, links)
        temp_path.replace(path)
        temp_path = None
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()


def recalculate_file(path, engine="Excel.Application"):
    """启动独立 Excel 实例重算并保存，确保公式产生可读取的缓存值。"""
    import openpyxl
    import pythoncom
    import win32com.client

    path = Path(path).resolve()
    if not path.is_file():
        raise ConfigError(f"公式重算文件不存在：{path}")
    if path.stat().st_size == 0:
        raise ConfigError(f"公式重算文件为空：{path}")
    try:
        with zipfile.ZipFile(path) as archive:
            damaged_member = archive.testzip()
        if damaged_member:
            raise ConfigError(f"公式重算文件损坏，XLSX 成员校验失败：{damaged_member}")
        check_book = openpyxl.load_workbook(path, read_only=True, data_only=False)
        check_book.close()
    except ConfigError:
        raise
    except (OSError, zipfile.BadZipFile, KeyError, ValueError) as exc:
        raise ConfigError(f"公式重算文件不是有效的 XLSX 工作簿：{path}；{exc}") from exc

    pythoncom.CoInitialize()
    app = book = None
    try:
        app = win32com.client.DispatchEx(engine)
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 3
        try:
            last = None
            book = None
            # 新启动的 Excel 实例偶发“Open 方法无效”，重试可覆盖启动竞态。
            for _ in range(3):
                try:
                    book = app.Workbooks.Open(
                        Filename=str(path),
                        UpdateLinks=0,
                        ReadOnly=False,
                        IgnoreReadOnlyRecommended=True,
                        AddToMru=False,
                    )
                    break
                except Exception as exc:
                    last = exc
                    time.sleep(1)
            if book is None:
                raise ConfigError(
                    f"Excel 无法打开公式重算文件：{path}（{path.stat().st_size} 字节）。"
                    f"请确认文件未被占用，并尝试用桌面 Excel 手动打开。原因：{last}"
                )
        except ConfigError:
            raise
        except Exception as exc:
            raise ConfigError(
                f"Excel 无法打开公式重算文件：{path}（{path.stat().st_size} 字节）。"
                "请确认文件未被占用，并尝试用桌面 Excel 手动打开。"
            ) from exc
        if book.ReadOnly:
            raise ConfigError("生成文件被占用，无法保存公式重算结果")
        app.CalculateFullRebuild()
        book.Save()
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


def _normalize_used_values(values, rows, cols):
    """把 COM 的 UsedRange.Value 统一成 rows×cols 的二维元组。"""
    if rows <= 0 or cols <= 0:
        return ()
    if rows == 1 and cols == 1:
        return ((values,),)
    if rows == 1:
        return (tuple(values),)
    if cols == 1:
        return tuple((item,) for item in values)
    return tuple(tuple(row) for row in values)


def _trim_empty_tail(values):
    """去掉末尾的全空行与全空列。"""
    rows = [list(row) for row in values]
    while rows and not any(value is not None for value in rows[-1]):
        rows.pop()
    width = max((len(row) for row in rows), default=0)
    while width:
        column = [row[width - 1] if len(row) >= width else None for row in rows]
        if any(value is not None for value in column):
            break
        width -= 1
    return tuple(tuple(row[:width]) for row in rows)


def _safe_sheet_title(name, used_titles):
    title = INVALID_SHEET_CHARS.sub("_", str(name or "Sheet"))[:31] or "Sheet"
    base = title
    index = 2
    while title in used_titles:
        suffix = f"_{index}"
        title = base[:31 - len(suffix)] + suffix
        index += 1
    used_titles.add(title)
    return title


def _write_sheets_to_xlsx(sheets, destination):
    """按原行列位置把读取到的值写入新的 .xlsx。"""
    import openpyxl

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    used_titles = set()
    for name, values, first_row, first_col in sheets:
        sheet = workbook.create_sheet(title=_safe_sheet_title(name, used_titles))
        for row_index, row in enumerate(values, start=first_row):
            for col_index, value in enumerate(row, start=first_col):
                sheet.cell(row=row_index, column=col_index, value=value)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()


def _trusted_excel_dirs():
    """读取 Excel 受信任位置；只有受信任位置里的文件才会执行模板宏。"""
    import winreg

    found = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Office") as office:
            versions = [
                winreg.EnumKey(office, index)
                for index in range(winreg.QueryInfoKey(office)[0])
            ]
    except OSError:
        return found

    for version in versions:
        key_path = (
            rf"Software\Microsoft\Office\{version}"
            r"\Excel\Security\Trusted Locations"
        )
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as locations:
                count = winreg.QueryInfoKey(locations)[0]
                for index in range(count):
                    name = winreg.EnumKey(locations, index)
                    try:
                        with winreg.OpenKey(locations, name) as location:
                            path, _ = winreg.QueryValueEx(location, "Path")
                    except OSError:
                        continue
                    found.append(path)
        except OSError:
            continue
    return found


def _writable_trusted_dir():
    """返回可写的受信任目录；跳过 XLSTART。"""
    for raw in _trusted_excel_dirs():
        directory = Path(os.path.expandvars(raw))
        if not directory.is_dir():
            continue
        if "xlstart" in str(directory).lower():
            continue
        try:
            probe = directory / f".jde_write_probe_{uuid.uuid4().hex}"
            probe.write_text("", encoding="utf-8")
            probe.unlink()
        except OSError:
            continue
        return directory
    return None


def convert_legacy_excel_to_xlsx(source, destination, engine="Excel.Application"):
    """把 JDE 导出的旧版 .xls 转为 openpyxl 可读的 .xlsx。

    JDE 输出模板自带宏，必须位于 Excel 受信任位置并启用内容才会执行；
    先复制到受信任目录，优先直接另存为 xlsx，失败时退回读取单元格值重建。
    """
    import pythoncom
    import shutil
    import win32com.client

    source = Path(source).resolve()
    destination = Path(destination).resolve()
    if source == destination:
        raise ConfigError("报表转换源文件与目标文件不能相同")
    if not source.is_file():
        raise ConfigError(f"报表源文件不存在：{source}")
    destination.parent.mkdir(parents=True, exist_ok=True)

    trusted_dir = _writable_trusted_dir()
    working_source = source
    temp_copy = None
    temp_xlsx = None
    if trusted_dir is None:
        logger.warning(
            "未找到可写的 Excel 受信任位置，模板宏可能不会运行；"
            "请在信任中心添加受信任位置以得到与“启用内容”一致的结果"
        )
    else:
        token = uuid.uuid4().hex
        temp_copy = trusted_dir / f"_jde_report_{token}.xls"
        temp_xlsx = trusted_dir / f"_jde_report_{token}.xlsx"
        shutil.copy2(source, temp_copy)
        working_source = temp_copy
        logger.info(f"报表已复制到受信任位置以运行模板宏：{trusted_dir}")

    sheets = []
    saved = False
    pythoncom.CoInitialize()
    app = book = None
    try:
        app = win32com.client.DispatchEx(engine)
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 1  # 启用内容：允许 JDE 模板宏运行
        book = app.Workbooks.Open(str(working_source), UpdateLinks=0, ReadOnly=False)
        save_target = temp_xlsx if temp_xlsx is not None else destination
        try:
            book.SaveAs(str(save_target), 51)
            saved = save_target.is_file() and save_target.stat().st_size > 0
            if not saved:
                logger.warning(
                    "Excel 另存为 xlsx 未生成文件（通常是被安全策略阻止），"
                    "改为读取单元格值重建"
                )
        except Exception as exc:
            logger.warning(f"Excel 另存为 xlsx 失败，改为读取单元格值重建：{exc}")
            saved = False

        if not saved:
            logger.info(f"报表已启用模板宏打开：工作表数={book.Sheets.Count}")
            for index in range(1, book.Sheets.Count + 1):
                sheet = book.Sheets(index)
                used = sheet.UsedRange
                logger.info(
                    f"  工作表 {sheet.Name}：{used.Rows.Count} 行 × "
                    f"{used.Columns.Count} 列（起始 {used.Row},{used.Column}）"
                )
                values = _trim_empty_tail(
                    _normalize_used_values(
                        used.Value, used.Rows.Count, used.Columns.Count
                    )
                )
                sheets.append((sheet.Name, values, used.Row, used.Column))
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
                if temp_xlsx is not None and saved:
                    try:
                        shutil.move(str(temp_xlsx), str(destination))
                    except OSError as exc:
                        logger.warning(f"移动转换结果失败：{exc}")
                        saved = False
                for leftover in (temp_copy, temp_xlsx):
                    if leftover is None:
                        continue
                    try:
                        if leftover.exists():
                            leftover.unlink()
                    except OSError as exc:
                        logger.warning(f"删除受信任位置临时文件失败：{exc}")

    if saved:
        logger.info("报表已由 Excel 直接另存为 xlsx（保留原格式）")
        return
    if not sheets:
        raise ConfigError(f"报表未读取到任何工作表：{source}")
    _write_sheets_to_xlsx(sheets, destination)
    if not destination.is_file() or destination.stat().st_size == 0:
        raise ConfigError(f"报表转换后文件无效：{destination}")


def _journal_title(value):
    """分录区标题规范化：忽略空白并统一“帐/账”。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).replace("帐", "账")


def journal_region(sheet):
    """按账号/借项金额/贷项金额/备注四个标题动态定位分录区。

    标题行可以不在第一行，四列也可以不相邻；数据区从表头下一行开始，
    到最后一个有账号的行结束。
    返回 (表头行, {标题: 列号}, 数据行号列表)。
    """
    expected = {_journal_title(title) for title in JOURNAL_TITLES}
    matches = []
    for row in sheet.iter_rows():
        found = {}
        for cell in row:
            title = _journal_title(cell.value)
            if title in expected:
                found[title] = cell.column
        if len(found) == len(expected):
            matches.append((row[0].row, found))
    if not matches:
        raise ConfigError("未找到 账号/借项金额/贷项金额/备注 标题行")
    if len(matches) > 1:
        raise ConfigError(f"无法唯一定位分录区表头：找到 {len(matches)} 处")
    header_row, columns = matches[0]
    account_column = columns["账号"]
    last_row = None
    for row in range(header_row + 1, sheet.max_row + 1):
        value = sheet.cell(row, account_column).value
        if value is not None and str(value).strip():
            last_row = row
    if last_row is None:
        return header_row, columns, []
    return header_row, columns, list(range(header_row + 1, last_row + 1))


def copy_sheet(book, source_sheet, target_name):
    """把源工作表的值、样式、合并区域与批注整体复制到目标工作簿同名页。"""
    if target_name in book.sheetnames:
        book.remove(book[target_name])
    destination = book.create_sheet(target_name)
    for merged in source_sheet.merged_cells.ranges:
        destination.merge_cells(str(merged))
    for row in range(1, source_sheet.max_row + 1):
        for col in range(1, source_sheet.max_column + 1):
            source_cell = source_sheet.cell(row=row, column=col)
            target_cell = destination.cell(row=row, column=col)
            target_cell.value = source_cell.value
            target_cell._style = copy(source_cell._style)
            target_cell.number_format = source_cell.number_format
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)
            if source_cell.comment:
                target_cell.comment = source_cell.comment
    return destination
