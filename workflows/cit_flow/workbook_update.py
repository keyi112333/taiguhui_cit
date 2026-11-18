#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业所得税工作簿更新规则。

实现需求中的：
- 会计期间写入计算底稿；
- profit 表历史金额数值化、当期金额保留公式；
- 申报底稿税款所属期间季度更新；
- 利润表-当季季度文字更新；
- 利润表与 R55GL1005 整表写入；
- 分录区读取与批号回写。
"""
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl

from config.config_loader import ConfigError
from log.loggingConfig import logger
from workflows.cit_flow.excel_ops import (
    external_link_parts,
    journal_region,
    save_with_openpyxl,
)

# 需求给定的当期利润公式；与模板 B18 完全一致。
PROFIT_FORMULA = (
    '=IFERROR(INDEX(利润表!$1:$1048576,MATCH("*利润总额*",利润表!$A:$A,0),'
    'MATCH("本期数",利润表!$5:$5,0)),0)-IFERROR(INDEX(利润表!$1:$1048576,'
    'MATCH("*投资收益*",利润表!$B:$B,0),MATCH("本期数",利润表!$5:$5,0)),0)'
)

PERIOD_LABEL = "会计期间："
BATCH_LABEL = "批号："
DECLARATION_LABEL = "税款所属期间："


def _text(value):
    return str(value).strip() if value is not None else ""


def _label(value):
    return re.sub(r"\s+", "", str(value)) if value is not None else ""


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _as_number(value):
    """把单元格值转成有限数值；公式/文本/空值返回 None。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if number == number and number not in (float("inf"), float("-inf")) else None
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return number if number == number and number not in (float("inf"), float("-inf")) else None


def right_cell(sheet, cell):
    """返回标签右侧第一个可写单元格；标签在合并区域内时跳过整个合并区域。"""
    column = cell.column
    for area in sheet.merged_cells.ranges:
        if area.min_row <= cell.row <= area.max_row and area.min_col <= cell.column <= area.max_col:
            column = area.max_col
            break
    return sheet.cell(cell.row, column + 1)


def input_cell_at(sheet, row, col):
    """返回输入单元格；若落在合并区域内，返回该合并区左上角。"""
    for area in sheet.merged_cells.ranges:
        if area.min_row <= row <= area.max_row and area.min_col <= col <= area.max_col:
            return sheet.cell(area.min_row, area.min_col)
    return sheet.cell(row, col)


def _find_label(sheet, target, max_row=None, max_col=None, exact=True):
    """在工作表中查找标签单元格；找到多处时返回第一处。"""
    max_row = max_row or sheet.max_row
    max_col = max_col or sheet.max_column
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            value = _label(sheet.cell(row, col).value)
            if (value == _label(target)) if exact else (_label(target) in value):
                return sheet.cell(row, col)
    return None


def _merged_top_left(sheet, row, col):
    """若单元格在合并区域内，返回合并区左上角；否则返回自身坐标。"""
    for area in sheet.merged_cells.ranges:
        if area.min_row <= row <= area.max_row and area.min_col <= col <= area.max_col:
            return area.min_row, area.min_col
    return row, col


def set_calculation_period(sheet, period_first):
    """把“会计期间：”右侧单元格写为会计期间首日。"""
    label = _find_label(sheet, PERIOD_LABEL, exact=False)
    if label is None:
        raise ConfigError(f"{sheet.title} 未找到“{PERIOD_LABEL}”单元格")
    target = right_cell(sheet, label)
    target.value = period_first
    target.number_format = "yyyy/m/d"
    logger.info(f"会计期间已更新为 {period_first:%Y/%m/%d}：{sheet.title}!{target.coordinate}")
    return target.coordinate


def _find_profit_table(sheet):
    """定位 profit/CIT/JDE计提CIT 表头行与月份列。"""
    for row in range(1, min(sheet.max_row, 60) + 1):
        profit_col = cit_col = jde_col = None
        for col in range(1, sheet.max_column + 1):
            text = _label(sheet.cell(row, col).value).lower()
            if text == "profit":
                profit_col = col
            elif text.startswith("cit"):
                cit_col = col
            elif "jde" in text and "cit" in text:
                jde_col = col
        if profit_col:
            return row, profit_col, cit_col or profit_col + 1, jde_col or profit_col + 2
    raise ConfigError(f"{sheet.title} 未找到 profit 表头")


def _month_rows(sheet, header_row, date_col):
    """读取表头下方以日期为月份标识的数据行。"""
    rows = {}
    for row in range(header_row + 1, sheet.max_row + 1):
        moment = _as_date(sheet.cell(row, date_col).value)
        if moment is not None:
            rows[(moment.year, moment.month)] = row
    if not rows:
        raise ConfigError(f"{sheet.title} 的 profit 表没有月份行")
    return rows


def _previous_profit_values(previous_path, sheet_name, date_col, profit_col):
    """读取上期运行表已重算的历史 profit 数值。"""
    if not previous_path:
        return {}
    path = Path(previous_path)
    if not path.is_file():
        logger.warning(f"上期运行表不存在，历史金额沿用模板：{path}")
        return {}
    book = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in book.sheetnames:
            logger.warning(f"上期运行表缺少工作表 {sheet_name}，历史金额沿用模板")
            return {}
        sheet = book[sheet_name]
        values = {}
        for row in range(1, sheet.max_row + 1):
            moment = _as_date(sheet.cell(row, date_col).value)
            if moment is None:
                continue
            number = _as_number(sheet.cell(row, profit_col).value)
            if number is not None:
                values[(moment.year, moment.month)] = number
        logger.info(f"已读取上期运行表历史 profit：{path}（{len(values)} 个月）")
        return values
    finally:
        book.close()


def _cached_profit_values(path, sheet_name, date_col, profit_col):
    """读取模板/底稿中公式的缓存值，用于没有上期运行表时的历史金额。"""
    if not path:
        return {}
    path = Path(path)
    if not path.is_file():
        return {}
    book = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in book.sheetnames:
            return {}
        sheet = book[sheet_name]
        values = {}
        for row in range(1, sheet.max_row + 1):
            moment = _as_date(sheet.cell(row, date_col).value)
            if moment is None:
                continue
            number = _as_number(sheet.cell(row, profit_col).value)
            if number is not None:
                values[(moment.year, moment.month)] = number
        return values
    finally:
        book.close()


def update_profit_table(book, sheet_name, period_first, previous_path=None,
                        cached_path=None):
    """更新 profit 表：历史金额数值化，当期金额保留公式。

    历史金额优先取上期运行表；没有上期时取模板/底稿的公式缓存值。
    返回 (当前行号, 表头行, profit 列, 月份行字典)。
    """
    if sheet_name not in book.sheetnames:
        raise ConfigError(f"计算底稿缺少工作表：{sheet_name}")
    sheet = book[sheet_name]
    set_calculation_period(sheet, period_first)

    header_row, profit_col, cit_col, jde_col = _find_profit_table(sheet)
    date_col = 1
    # 月份通常位于 A 列；若 A 列不是日期，退到表头左侧第一列。
    if not any(_as_date(sheet.cell(r, date_col).value) for r in range(header_row + 1, sheet.max_row + 1)):
        for col in range(1, profit_col):
            if any(_as_date(sheet.cell(r, col).value) for r in range(header_row + 1, sheet.max_row + 1)):
                date_col = col
                break
    rows = _month_rows(sheet, header_row, date_col)
    current = (period_first.year, period_first.month)
    if current not in rows:
        raise ConfigError(
            f"{sheet.title} 的 profit 表没有 {period_first:%Y-%m} 月份行"
        )

    previous = _previous_profit_values(previous_path, sheet_name, date_col, profit_col)
    cached = _cached_profit_values(cached_path, sheet_name, date_col, profit_col)
    for key, row in rows.items():
        cell = sheet.cell(row, profit_col)
        if key < current:
            number = previous.get(key)
            if number is None:
                number = cached.get(key)
            if number is None:
                number = _as_number(cell.value)
            if number is None:
                number = 0.0
            cell.value = number
        elif key == current:
            cell.value = PROFIT_FORMULA
        # 未来月份保持模板原值，不做处理。

    logger.info(
        f"profit 表已更新：{sheet.title} 表头第 {header_row} 行，"
        f"当期 {period_first:%Y-%m} 第 {rows[current]} 行；"
        f"历史 {sum(1 for key in rows if key < current)} 个月已数值化"
    )
    return rows[current], header_row, profit_col, rows


def _find_quarter_label_cell(sheet):
    """定位申报底稿的“至”单元格。"""
    label = _find_label(sheet, DECLARATION_LABEL, exact=False)
    if label is None:
        raise ConfigError(f"{sheet.title} 未找到“{DECLARATION_LABEL}”单元格")
    start = right_cell(sheet, label)
    to_cell = None
    for col in range(start.column, sheet.max_column + 1):
        if _label(sheet.cell(label.row, col).value) == "至":
            to_cell = sheet.cell(label.row, col)
            break
    if to_cell is None:
        raise ConfigError(f"{sheet.title} 未找到税款所属期间中的“至”")
    return start, input_cell_at(sheet, to_cell.row, to_cell.column + 1)


def update_declaration_period(book, sheet_name, quarter_start, quarter_end):
    """更新申报底稿“税款所属期间：”起止日期。"""
    if sheet_name not in book.sheetnames:
        raise ConfigError(f"申报底稿缺少工作表：{sheet_name}")
    sheet = book[sheet_name]
    start_cell, end_cell = _find_quarter_label_cell(sheet)
    start_cell.value = quarter_start
    start_cell.number_format = "yyyy-mm-dd"
    end_cell.value = quarter_end
    end_cell.number_format = "yyyy-mm-dd"
    logger.info(
        f"申报底稿税款所属期间已更新：{sheet.title}!{start_cell.coordinate} "
        f"{quarter_start:%Y-%m-%d} 至 {end_cell.coordinate} {quarter_end:%Y-%m-%d}"
    )
    return start_cell.coordinate, end_cell.coordinate


def update_quarter_label(book, sheet_name, quarter_start, quarter_end):
    """更新“利润表-当季”页“利润表”下方的季度文字。"""
    if sheet_name not in book.sheetnames:
        raise ConfigError(f"季度利润表缺少工作表：{sheet_name}")
    sheet = book[sheet_name]
    label = _find_label(sheet, "利润表", max_row=20, exact=True)
    if label is None:
        raise ConfigError(f"{sheet.title} 未找到“利润表”标题")
    row, col = _merged_top_left(sheet, label.row + 1, label.column)
    target = sheet.cell(row, col)
    text = f"{quarter_start.year} 年 {quarter_start.month}-{quarter_end.month} 月"
    target.value = text
    logger.info(f"季度利润表已更新：{sheet.title}!{target.coordinate} = {text}")
    return target.coordinate, text


def _header_columns(sheet, max_row=20):
    """在利润表中定位表头：本期数、本年累计数、行次。"""
    for row in range(1, min(sheet.max_row, max_row) + 1):
        current_col = ytd_col = row_no_col = None
        for col in range(1, sheet.max_column + 1):
            text = _label(sheet.cell(row, col).value)
            if "本期数" in text:
                current_col = col
            elif "本年累计" in text or "累计数" in text:
                ytd_col = col
            elif text in ("行次", "行号"):
                row_no_col = col
        if current_col and ytd_col:
            return row, current_col, ytd_col, row_no_col
    return None


def _source_profit_map(sheet, header_row, current_col, ytd_col, row_no_col):
    """构建源利润表的两套索引：按行次、按项目名称。"""
    by_row_no = {}
    by_label = {}
    for row in range(header_row + 1, sheet.max_row + 1):
        current = sheet.cell(row, current_col).value
        ytd = sheet.cell(row, ytd_col).value
        row_no = None
        if row_no_col:
            row_no = _as_number(sheet.cell(row, row_no_col).value)
        labels = []
        for col in range(1, (row_no_col or current_col)):
            text = _text(sheet.cell(row, col).value)
            if text:
                labels.append(re.sub(r"\s+", "", text))
        if row_no is None and not labels:
            continue
        if row_no is not None:
            by_row_no[int(row_no)] = (current, ytd)
        if labels:
            by_label[labels[-1]] = (current, ytd)
    return by_row_no, by_label


def copy_income_statement(source_path, source_sheet_name, book, target_sheet_name):
    """把 JDE 导出的利润表按行次/项目名称写入模板的利润表页。

    只写本期数与本年累计数，保留模板的单位名称、行次和格式。
    """
    source_book = openpyxl.load_workbook(source_path, data_only=True)
    try:
        if source_sheet_name and source_sheet_name in source_book.sheetnames:
            source = source_book[source_sheet_name]
        else:
            candidates = [name for name in source_book.sheetnames
                          if "利润" in name or "profit" in name.lower()]
            source = source_book[candidates[0] if candidates else source_book.sheetnames[0]]

        source_header = _header_columns(source)
        if source_header is None:
            raise ConfigError(
                f"下载的利润表未找到“本期数/本年累计数”表头：{source_path}"
            )
        header_row, current_col, ytd_col, row_no_col = source_header
        by_row_no, by_label = _source_profit_map(
            source, header_row, current_col, ytd_col, row_no_col
        )
        if not by_row_no and not by_label:
            raise ConfigError(f"下载的利润表没有可写入的数据：{source_path}")

        if target_sheet_name not in book.sheetnames:
            raise ConfigError(f"模板缺少利润表工作表：{target_sheet_name}")
        target = book[target_sheet_name]
        target_header = _header_columns(target)
        if target_header is None:
            raise ConfigError(f"{target.title} 未找到“本期数/本年累计数”表头")
        t_header, t_current, t_ytd, t_row_no = target_header

        # 先清空模板数据区，避免下载报表没有数据的项目保留上月旧值。
        for row in range(t_header + 1, target.max_row + 1):
            input_cell_at(target, row, t_current).value = None
            input_cell_at(target, row, t_ytd).value = None

        written = 0
        for row in range(t_header + 1, target.max_row + 1):
            row_no = _as_number(target.cell(row, t_row_no).value) if t_row_no else None
            values = by_row_no.get(int(row_no)) if row_no is not None else None
            if values is None:
                labels = []
                for col in range(1, t_current):
                    text = _text(target.cell(row, col).value)
                    if text:
                        labels.append(re.sub(r"\s+", "", text))
                if labels:
                    values = by_label.get(labels[-1])
            if values is None:
                continue
            input_cell_at(target, row, t_current).value = _as_number(values[0])
            input_cell_at(target, row, t_ytd).value = _as_number(values[1])
            written += 1

        if written == 0:
            raise ConfigError("下载的利润表与模板项目未能匹配，未写入任何数据")

        # 同步报表期间与单位名称，保持表头与数据一致。
        period_text = None
        for row in source.iter_rows(min_row=1, max_row=6):
            for cell in row:
                if isinstance(cell.value, str) and re.search(r"\d{4}\s*年\s*\d{1,2}\s*月", cell.value):
                    period_text = cell.value
                    break
            if period_text:
                break
        if period_text:
            for row in target.iter_rows(min_row=1, max_row=6):
                for cell in row:
                    if isinstance(cell.value, str) and re.search(r"\d{4}\s*年\s*\d{1,2}\s*月", cell.value):
                        cell.value = period_text
                        break

        company_label = _find_label(source, "单位名称", max_row=10, exact=False)
        target_company_label = _find_label(target, "单位名称", max_row=10, exact=False)
        if company_label is not None and target_company_label is not None:
            company = right_cell(source, company_label).value
            if _text(company):
                right_cell(target, target_company_label).value = company

        logger.info(f"利润表已写入模板：{target.title}，匹配 {written} 个项目")
        return written
    finally:
        source_book.close()


def copy_r55gl1005(source_path, source_sheet_name, book, target_sheet_name="R55GL1005"):
    """用本期 R55GL1005 整表覆盖模板页。"""
    from workflows.cit_flow.excel_ops import copy_sheet

    source_book = openpyxl.load_workbook(source_path, data_only=True)
    try:
        if source_sheet_name and source_sheet_name in source_book.sheetnames:
            source = source_book[source_sheet_name]
        elif target_sheet_name in source_book.sheetnames:
            source = source_book[target_sheet_name]
        else:
            source = source_book[source_book.sheetnames[0]]
        copy_sheet(book, source, target_sheet_name)
        logger.info(
            f"R55GL1005 已整表写入：{target_sheet_name}"
            f"（{source.max_row} 行 × {source.max_column} 列）"
        )
        return target_sheet_name
    finally:
        source_book.close()


def _zero_amount(value):
    """金额为空、空串或数值 0 时视为零。"""
    if value is None:
        return True
    text = str(value).strip()
    if not text:
        return True
    try:
        return float(text) == 0
    except ValueError:
        return False


def journal_rows(path, sheet_name):
    """读取计算底稿分录区，返回 ([(账号, 借项, 贷项, 备注)], 说明)。"""
    book = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in book.sheetnames:
            raise ConfigError(f"计算底稿缺少工作表：{sheet_name}")
        sheet = book[sheet_name]
        _, columns, region = journal_region(sheet)
        rows = []
        for row in region:
            code = sheet.cell(row, columns["账号"]).value
            if code is None or not str(code).strip():
                continue
            debit = sheet.cell(row, columns["借项金额"]).value
            credit = sheet.cell(row, columns["贷项金额"]).value
            if _zero_amount(debit) and _zero_amount(credit):
                continue
            rows.append((
                str(code).strip(),
                debit,
                credit,
                sheet.cell(row, columns["备注"]).value,
            ))
        if not rows:
            return [], ""
        description = str(rows[0][3] or "").strip()
        if not description:
            raise ConfigError("计算底稿分录第一行没有备注，无法作为 JDE 说明")
        return rows, description
    finally:
        book.close()


def write_batch_number(path, sheet_name, batch):
    """把 JDE 批号回写到“批号：”右侧单元格。"""
    if not batch:
        return None
    links = external_link_parts(path)
    book = openpyxl.load_workbook(path)
    try:
        if sheet_name not in book.sheetnames:
            raise ConfigError(f"计算底稿缺少工作表：{sheet_name}")
        sheet = book[sheet_name]
        label = _find_label(sheet, BATCH_LABEL, exact=False)
        if label is None:
            raise ConfigError("计算底稿未找到“批号：”单元格")
        target = right_cell(sheet, label)
        target.value = str(batch)
        target.number_format = "@"
        save_with_openpyxl(book, path, links)
        logger.info(f"JDE 批号 {batch} 已回写：{sheet.title}!{target.coordinate}")
        return target.coordinate
    finally:
        book.close()


def clear_batch_number(book, sheet_name):
    """清空上期批号，避免本次分录失败时沿用旧批号。"""
    if sheet_name not in book.sheetnames:
        return None
    sheet = book[sheet_name]
    label = _find_label(sheet, BATCH_LABEL, exact=False)
    if label is None:
        return None
    target = right_cell(sheet, label)
    target.value = None
    logger.info(f"已清空上期批号：{sheet.title}!{target.coordinate}")
    return target.coordinate
