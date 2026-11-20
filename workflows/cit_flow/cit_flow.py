#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业所得税流程编排。

单行业务链路：
导出 R55GL1005 → 导出利润表 → 以上期运行表为底稿生成本期文件 →
写入利润表/R55GL1005 → 更新会计期间与 profit 表 → 更新申报底稿季度 →
更新季度利润表文字 → Excel 重算 → 录入 JDE 分录 → 回写批号。
"""
import calendar
import re
import shutil
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.workbook.properties import CalcProperties

from config.config_loader import ConfigError
from log.loggingConfig import logger
from utils.helpers import get_timestamp
from web.JDE.jde_function import (
    create_journal_entry,
    download_income_statement,
    download_r55gl1005,
)
from workflows.cit_flow.excel_ops import (
    convert_legacy_excel_to_xlsx,
    external_link_parts,
    recalculate_file,
    save_with_openpyxl,
)
from workflows.cit_flow.workbook_update import (
    clear_batch_number,
    copy_income_statement,
    copy_r55gl1005,
    journal_rows,
    update_declaration_period,
    update_profit_table,
    update_quarter_sheet,
    write_batch_number,
)

RUN_FILE_PATTERN = re.compile(r"^企业所得税\s+(.+?)\s+(\d{6})\.xlsx$", re.IGNORECASE)
OLD_RUN_FILE_PATTERN = re.compile(r"^企业所得税计算-\d{8}\.xlsx$", re.IGNORECASE)
DEFAULT_OUTPUT_NAME = "企业所得税 {公司} {YYYYMM}.xlsx"


class PendingModule(Exception):
    """流程等待录制或报表文件，不代表已完成。"""


def configured_period(value):
    """会计期间配置值；空白（含只有空格）视为未配置，按自然月返回今天。"""
    text = "" if value is None else str(value).strip()
    return text or date.today()


def month_bounds(period):
    """会计期间自然月第一天和最后一天。"""
    if isinstance(period, (datetime, date)):
        year, month = period.year, period.month
    else:
        text = "" if period is None else str(period).strip()
        matchd = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
        if not matchd:
            raise ConfigError(
                f"会计期间格式不正确，应为 YYYY-MM 或 YYYY/MM 或 YYYYMM 等可识别格式；当前值：{text}"
            )
        year, month = int(matchd.group(1)), int(matchd.group(2))
    if not (1 <= month <= 12):
        raise ConfigError(f"会计期间月份不正确，应为 1-12；当前值：{month}")
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def quarter_bounds(period_first):
    """按自然季度返回（季度首日, 季度末日）：1-3、4-6、7-9、10-12。"""
    quarter = (period_first.month - 1) // 3 + 1
    start_month = 3 * (quarter - 1) + 1
    end_month = start_month + 2
    start = date(period_first.year, start_month, 1)
    end = date(period_first.year, end_month, calendar.monthrange(period_first.year, end_month)[1])
    return start, end


def format_jde_date(value, label):
    """把日期格式化为 JDE 的 MM/DD/YYYY。"""
    if isinstance(value, (datetime, date)):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, pattern).strftime("%m/%d/%Y")
        except ValueError:
            continue
    raise ConfigError(f"{label} 日期格式不正确，应为 YYYY-MM-DD 或 MM/DD/YYYY 等可识别格式；当前值：{text}")


def resolve_template(shared_path, short_name, explicit=None):
    """在共享盘目录解析企业所得税模板。

    优先使用显式配置；其次按公司简称匹配固定命名；最后回退到最新的
    “企业所得税*.xlsx”（上期运行表本身也可作为下一次的底稿）。
    """
    shared_path = Path(shared_path)
    if explicit:
        template = Path(explicit)
        if not template.is_absolute():
            template = shared_path / template
        if not template.is_file():
            raise ConfigError(f"模板文件不存在：{template}")
        return template

    for name in (
        f"企业所得税计算底稿 {short_name}.xlsx",
        f"企业所得税计算 {short_name}.xlsx",
        "企业所得税计算底稿.xlsx",
    ):
        candidate = shared_path / name
        if candidate.is_file():
            return candidate

    files = [path for path in shared_path.glob("*企业所得税*.xlsx")
             if not path.name.startswith("~$")]
    dedicated = [path for path in files
                 if not RUN_FILE_PATTERN.match(path.name)
                 and not OLD_RUN_FILE_PATTERN.match(path.name)]
    pool = dedicated or files
    if pool:
        template = max(pool, key=lambda path: path.stat().st_mtime)
        logger.warning(f"未找到固定命名模板，改用共享盘最新企业所得税工作簿：{template}")
        return template
    raise ConfigError(
        f"共享盘缺少企业所得税模板：{shared_path}；"
        f"请放置“企业所得税计算底稿 {short_name}.xlsx”或配置 企业所得税配置.模板文件"
    )


def find_previous_run(shared_path, output_path=None, company=None, period_first=None):
    """查找上一个会计期间的运行表：企业所得税 公司 YYYYMM.xlsx。

    优先精确匹配上一个自然月；找不到时返回 None，由调用方回退到模板数据。
    """
    shared_path = Path(shared_path)
    current = period_first or date.today()
    year, month = current.year, current.month - 1
    if month == 0:
        year, month = year - 1, 12
    period_label = f"{year}{month:02d}"
    output_resolved = Path(output_path).resolve() if output_path else None

    if company:
        target = shared_path / f"企业所得税 {company} {period_label}.xlsx"
        if target.is_file() and (output_resolved is None or target.resolve() != output_resolved):
            return target
        return None

    candidates = []
    for path in shared_path.glob("企业所得税 *.xlsx"):
        if path.name.startswith("~$"):
            continue
        match = RUN_FILE_PATTERN.match(path.name)
        if match is None:
            continue
        if output_resolved is not None and path.resolve() == output_resolved:
            continue
        if match.group(2) < f"{current:%Y%m}":
            candidates.append((match.group(2), path))
    if not candidates:
        return None
    return max(candidates, key=lambda item: item[0])[1]


def build_output_path(shared_path, company, period_first, name_template=None):
    """生成共享盘本期运行表路径。"""
    template = str(name_template or DEFAULT_OUTPUT_NAME).strip() or DEFAULT_OUTPUT_NAME
    name = template.replace("{公司}", str(company))
    name = name.replace("{YYYYMM}", f"{period_first:%Y%m}")
    name = name.replace("{YYYYMMDD}", date.today().strftime("%Y%m%d"))
    if not Path(name).suffix:
        name += ".xlsx"
    return Path(shared_path) / name


class CITFlow:
    """企业所得税单公司流程。"""

    def __init__(self, context):
        self.context = context
        self.config = context.config

    def get(self, key, default=None):
        return self.config.get(f"企业所得税配置.{key}", default)

    def accounting_period(self):
        first, _ = month_bounds(configured_period(self.get("会计期间")))
        return first

    def _resolve_relative(self, value, path_type):
        path = Path(value)
        if not path.is_absolute():
            path = self.config.get_path(path_type) / path
        return path

    def export_r55gl1005(self):
        """导出本期 R55GL1005 并转换为 xlsx。"""
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；无法导出 R55GL1005")

        company = self.get("明细账自公司")
        if company is None or not str(company).strip():
            raise ConfigError("导出 R55GL1005 缺少 企业所得税配置.明细账自公司")

        destination = self.context.get("jde_r55gl1005_path") or self.get("R55GL1005文件")
        if destination:
            destination = self._resolve_relative(destination, "input")
        else:
            destination = self.config.get_path("output") / f"R55GL1005_{get_timestamp()}.xlsx"
        destination = Path(destination).with_suffix(".xlsx")
        destination.parent.mkdir(parents=True, exist_ok=True)

        period, _ = month_bounds(configured_period(self.get("会计期间")))
        raw_path = download_r55gl1005(
            page,
            company=company,
            destination=destination.with_suffix(".xls"),
            period=period,
        )
        convert_legacy_excel_to_xlsx(raw_path, destination)
        logger.info(f"R55GL1005 已转换为工作簿格式：{destination}")
        self.context.set("jde_r55gl1005_path", destination)
        return destination

    def export_income_statement(self):
        """从 JDE 导出本期利润表；每次运行、每行配置都重新导出，不复用历史文件。"""
        company = self.get("明细账自公司")
        if company is None or not str(company).strip():
            raise ConfigError("导出利润表缺少 企业所得税配置.明细账自公司")

        destination = self.get("利润表文件")
        if destination:
            destination = self._resolve_relative(destination, "input")
        else:
            destination = self.config.get_path("output") / f"利润表_{get_timestamp()}.xlsx"
        destination = Path(destination)
        if destination.suffix.lower() not in (".xlsx", ".xls"):
            destination = destination.with_suffix(".xlsx")
        destination.parent.mkdir(parents=True, exist_ok=True)

        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；无法导出利润表")

        period, _ = month_bounds(configured_period(self.get("会计期间")))
        raw_path = download_income_statement(
            page,
            company=company,
            destination=destination.with_suffix(".xls"),
            period=period,
        )
        convert_legacy_excel_to_xlsx(raw_path, destination)
        logger.info(f"利润表已转换为工作簿格式：{destination}")
        self.context.set("jde_income_statement_path", destination)
        return destination

    def prepare_output(self, template, previous, output):
        """以上期运行表为底稿复制出本期文件；没有上期时使用模板。"""
        source = Path(previous) if previous and Path(previous).is_file() else Path(template)
        output = Path(output)
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            output.unlink()
            logger.info(f"已覆盖同名运行表：{output}")
        shutil.copy2(source, output)
        logger.info(f"本期运行表已生成：{output}（底稿来源：{source}）")
        return output

    def update_workbook(self, output, r55_path, income_path, previous_path):
        """写入报表并更新会计期间、profit、申报底稿与季度文字。"""
        calc_sheet = self.get("计算底稿工作表", "企业所得税计算底稿")
        declaration_sheet = self.get("申报底稿工作表", "企业所得税申报底稿")
        quarter_sheet = self.get("利润表-当季工作表", "利润表-当季")
        profit_sheet = self.get("利润表目标工作表", "利润表")

        period_first = self.accounting_period()
        quarter_start, quarter_end = quarter_bounds(period_first)

        links = external_link_parts(output)
        book = openpyxl.load_workbook(output)
        try:
            copy_r55gl1005(
                r55_path,
                self.get("R55GL1005工作表"),
                book,
                "R55GL1005",
            )
            copy_income_statement(
                income_path,
                self.get("利润表源工作表"),
                book,
                profit_sheet,
            )
            update_profit_table(book, calc_sheet, period_first, previous_path, output)
            clear_batch_number(book, calc_sheet)
            update_declaration_period(book, declaration_sheet, quarter_start, quarter_end)
            update_quarter_sheet(
                book, quarter_sheet, period_first, profit_sheet, previous_path
            )
            book.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
            save_with_openpyxl(book, output, links)
        finally:
            book.close()
        logger.info(f"工作簿更新完成：{output}")
        return output

    def run(self):
        self.context.set("cit_status", "running")
        try:
            period_first = self.accounting_period()
            shared_path = Path(self.get("输出目录") or self.config.get_path("output"))
            if not shared_path.is_absolute():
                shared_path = self.config.get_path("output") / shared_path

            template = self.get("模板文件")
            if not template:
                template = resolve_template(shared_path, self.get("公司") or "")
            template = self._resolve_relative(template, "input")

            output = self.get("输出文件")
            if output:
                output = Path(output)
            else:
                output = build_output_path(
                    shared_path, self.get("公司") or "", period_first,
                    self.get("输出文件名"),
                )

            previous = self.get("上期运行表")
            if previous:
                previous = self._resolve_relative(previous, "input")
            else:
                previous = find_previous_run(
                    shared_path, output, self.get("公司") or "", period_first
                )
            if previous:
                logger.info(f"历史数据来源：{previous}")
            else:
                logger.info("未找到上期运行表，历史金额使用模板缓存值")

            r55_path = self.export_r55gl1005()
            income_path = self.export_income_statement()
            self.prepare_output(template, previous, output)
            self.update_workbook(output, r55_path, income_path, previous)

            engine = str(self.get("重算引擎", "Excel.Application")).strip()
            if engine.lower() == "openpyxl":
                raise ConfigError(
                    "企业所得税流程需要 Excel 重算以取得分录金额；"
                    "重算引擎不能配置为 openpyxl"
                )
            recalculate_file(output, engine)

            calc_sheet = self.get("分录工作表", "企业所得税计算底稿")
            rows, description = journal_rows(output, calc_sheet)
            if not rows:
                logger.info("本期计提金额为 0，计算底稿没有需要录入的 JDE 分录")
                self.context.set("jde_journal_rows", 0)
                self.context.set("cit_path", output)
                self.context.set("cit_status", "completed")
                return
            batch = create_journal_entry(
                self.context.get_browser("JDE").get_page(),
                company=self.get("分录公司", ""),
                description=description,
                gl_date=format_jde_date(date.today(), "总帐日期"),
                rows=rows,
                discard=not self._save_journal(),
            )
            self.context.set("jde_journal_batch", batch)
            self.context.set("jde_journal_rows", len(rows))
            self.context.set("cit_path", output)
            write_batch_number(output, calc_sheet, batch)

            # 回写批号用 openpyxl 保存会丢弃公式缓存；再重算一次，保证
            # 下月读取本期运行表时历史 profit 仍有数值。
            recalculate_file(output, engine)
            self.context.set("cit_status", "completed")
            logger.info(f"企业所得税流程完成：{output}（批号 {batch}）")
        except PendingModule as exc:
            self.context.set("cit_status", "pending")
            self.context.set("cit_pending_reason", str(exc))
            logger.warning(str(exc))
        except Exception:
            self.context.set("cit_status", "failed")
            raise

    def _save_journal(self):
        value = self.get("分录保存", False)
        if isinstance(value, str):
            return value.strip().lower() in ("true", "1", "yes", "是")
        return bool(value)
