"""按「登录配置」逐行运行：每行一个账号 + 一个共享盘目录，完整跑一遍企业所得税流程。"""

from datetime import datetime
from pathlib import Path

from config.config_loader import ConfigError
from core.context import TaskContext
from core.task_base import TaskBase
from log.loggingConfig import logger
from utils.credentials import CredentialNotFoundError, read_windows_credential
from utils.helpers import get_timestamp
from web.browser import BrowserManager
from workflows.cit_flow.cit_flow import (
    build_output_path,
    configured_period,
    find_previous_run,
    month_bounds,
    resolve_template,
)
from workflows.cit_flow.login_task import JDELoginTask
from workflows.cit_flow.run_report import (
    RESULT_FILE_TEMPLATE,
    append_run_result,
    build_notification_message,
    notify_wechat_robots,
)


RUN_SHEET = "登录配置"
REQUIRED_COLUMNS = ("window凭据", "凭据账号", "共享盘路径")


class SkipRun(Exception):
    """当前行凭据未配置，跳过该行且不计入失败。"""


def parse_shared_path(value):
    """从共享盘路径末段目录（如“太古汇 2501”）解析公司简称与公司代码。"""
    path = Path(str(value).strip())
    tokens = path.name.split()
    if len(tokens) != 2 or not tokens[1].isdigit():
        raise ConfigError(
            f"共享盘路径末段需为“公司简称 公司代码”（如 太古汇 2501），当前：{path.name}"
        )
    return path, tokens[0], tokens[1]


def read_headless(value):
    if isinstance(value, str):
        text = value.strip().lower()
        if text not in ("true", "false", "1", "0", "yes", "no"):
            raise ConfigError("JDE配置.headless 必须为 true 或 false")
        return text in ("true", "1", "yes")
    return bool(value)


class CITMainTask(TaskBase):
    def __init__(self, config):
        context = TaskContext()
        context.config = config
        super().__init__("企业所得税流程", context)

    def execute(self) -> bool:
        config = self.context.config
        if not config.get("JDE配置.jde_url", ""):
            raise ConfigError("请配置 JDE配置.jde_url")
        runs = config.get_rows(RUN_SHEET)
        if not runs:
            raise ConfigError(
                f"{RUN_SHEET} 页没有数据行，请填写 "
                f"{'、'.join(REQUIRED_COLUMNS)}；每行会完整跑一遍流程"
            )
        period = self.period_label()
        # 每行会覆盖下面这些键；先保存全局默认值，避免下一行沿用上一行的共享盘路径。
        defaults = {
            "模板文件": config.get("企业所得税配置.模板文件"),
            "上期运行表": config.get("企业所得税配置.上期运行表"),
            "输出文件": config.get("企业所得税配置.输出文件"),
        }
        failures = []
        records = []
        result_file = None
        for index, run in enumerate(runs, 1):
            label = f"[{index}/{len(runs)}] {run.get('共享盘路径')}"
            logger.info(f"开始执行 {label}")
            started = datetime.now()
            info = {}
            batch = ""
            result, error = "成功", ""
            for key, value in defaults.items():
                config.set(f"企业所得税配置.{key}", value)
            self.context.set("cit_status", None)
            self.context.set("jde_journal_batch", None)
            try:
                info = self.prepare(run)
                self.run_one(run, info)
                status = self.context.get("cit_status")
                if status == "pending":
                    result = "待补运行"
                    error = str(self.context.get("cit_pending_reason") or "")
                    logger.warning(f"{label} 待补运行：{error}")
                elif status not in (None, "completed"):
                    result = "失败"
                    error = str(self.context.get("cit_pending_reason") or "")
                    failures.append(f"{label}：{error}")
                    logger.error(f"{label} 执行失败：{error}")
                else:
                    logger.info(f"{label} 执行完成")
                batch = self.context.get("jde_journal_batch") or ""
            except SkipRun as exc:
                result, error = "跳过", str(exc)
                logger.warning(f"{label} 跳过：{exc}")
            except Exception as exc:
                result, error = "失败", str(exc)
                logger.exception(f"{label} 执行失败：{exc}")
                failures.append(f"{label}：{exc}")
            finally:
                short_name, company = self._run_identity(run)
                record = {
                    "会计期间": period,
                    "公司": info.get("short_name") or short_name,
                    "公司代码": info.get("company") or company,
                    "分录批号": batch,
                    "运行结果": result,
                    "异常情况": error,
                    "备注": "",
                    "运行开始时间": started,
                    "运行结束时间": datetime.now(),
                }
                records.append(record)
                try:
                    result_file = append_run_result(config, record) or result_file
                except Exception as exc:
                    logger.warning(f"运行结果表写入失败：{exc}")
        self._notify(config, period, records, result_file)
        if failures:
            raise ConfigError("以下账号执行失败：" + "；".join(failures))
        return True

    def period_label(self):
        """本次运行的会计期间（YYYYMM），未配置时按自然月。"""
        first, _ = month_bounds(
            configured_period(self.context.config.get("企业所得税配置.会计期间"))
        )
        return f"{first:%Y%m}"

    @staticmethod
    def _run_identity(run):
        """从共享盘路径解析公司简称与代码；解析失败时保留原始路径便于排查。"""
        try:
            _, short_name, code = parse_shared_path(str(run.get("共享盘路径") or ""))
            return short_name, code.zfill(5)
        except Exception:
            return str(run.get("共享盘路径") or ""), ""

    def _notify(self, config, period, records, result_file):
        skipped = sum(1 for record in records if record.get("运行结果") == "跳过")
        total = len(records) - skipped
        success = sum(1 for record in records if record.get("运行结果") == "成功")
        failed = total - success
        companies = []
        for record in records:
            name = str(record.get("公司") or "").strip()
            if name and name not in companies:
                companies.append(name)
        company = "、".join(companies) or str(
            config.get("企业所得税配置.公司", "") or "太古汇"
        )
        filename = (
            Path(result_file).name if result_file
            else RESULT_FILE_TEMPLATE.format(year=datetime.now().year)
        )
        message = build_notification_message(
            company, period, total, success, failed, filename
        )
        try:
            notify_wechat_robots(config, message, result_file)
        except Exception as exc:
            logger.warning(f"企业微信通知失败：{exc}")

    def prepare(self, run):
        """校验本行配置并把公司、模板、路径写入企业所得税配置。"""
        config = self.context.config
        values = {name: str(run.get(name) or "").strip() for name in REQUIRED_COLUMNS}
        missing = [name for name, value in values.items() if not value]
        if missing:
            raise ConfigError(f"{RUN_SHEET} 缺少：{', '.join(missing)}")

        shared_path, short_name, code = parse_shared_path(values["共享盘路径"])
        if not shared_path.is_dir():
            raise ConfigError(f"共享盘路径不存在：{shared_path}")

        template = resolve_template(
            shared_path, short_name, config.get("企业所得税配置.模板文件")
        )

        try:
            username, _ = read_windows_credential(values["window凭据"])
        except CredentialNotFoundError:
            raise SkipRun(f"Windows 凭据 {values['window凭据']} 不存在，已跳过") from None
        if username.split("\\")[-1].strip().lower() != values["凭据账号"].lower():
            logger.warning(
                f"Windows 凭据 {values['window凭据']} 的用户名 {username} "
                f"与「凭据账号」{values['凭据账号']} 不一致，请确认数据归属"
            )

        period_value = config.get("企业所得税配置.会计期间")
        first_day, _ = month_bounds(configured_period(period_value))
        period = f"{first_day:%Y%m}"
        company = code.zfill(5)
        stamp = get_timestamp()
        data_dir = config.get_path("output") / f"{short_name} {code} {period}"
        data_dir.mkdir(parents=True, exist_ok=True)

        output = config.get("企业所得税配置.输出文件")
        if output:
            output = Path(output)
        else:
            output = build_output_path(
                shared_path, short_name, first_day,
                config.get("企业所得税配置.输出文件名"),
            )

        previous = config.get("企业所得税配置.上期运行表")
        if previous:
            previous = Path(previous)
            if not previous.is_absolute():
                previous = shared_path / previous
        else:
            previous = find_previous_run(shared_path, output, short_name, first_day)

        overrides = {
            "JDE配置.jde_credential_name": values["window凭据"],
            "企业所得税配置.公司": short_name,
            "企业所得税配置.模板文件": str(template),
            "企业所得税配置.输出目录": str(shared_path),
            "企业所得税配置.输出文件": str(output),
            "企业所得税配置.上期运行表": str(previous) if previous else "",
            "企业所得税配置.明细账自公司": company,
            "企业所得税配置.明细账至公司": company,
            "企业所得税配置.分录公司": company,
            "企业所得税配置.R55GL1005文件": str(data_dir / f"R55GL1005_{stamp}.xlsx"),
            "企业所得税配置.利润表文件": str(data_dir / f"利润表_{stamp}.xlsx"),
        }
        for key, value in overrides.items():
            config.set(key, value)
        logger.info(
            f"运行参数：凭据={values['window凭据']} 账号={values['凭据账号']} "
            f"公司={company} 共享盘={shared_path} 数据源={data_dir} "
            f"本期运行表={output} 上期运行表={previous or '无'}"
        )
        return {
            "credential": values["window凭据"],
            "company": company,
            "short_name": short_name,
            "period": period,
            "shared_path": shared_path,
            "template": template,
            "output": output,
            "previous": previous,
            "data_dir": data_dir,
        }

    def run_one(self, run, info=None):
        info = info or self.prepare(run)
        config = self.context.config
        browser = BrowserManager(headless=read_headless(config.get("JDE配置.headless", False)))
        self.context.set_browser("JDE", browser)
        try:
            browser.start()
            if not JDELoginTask(self.context).run():
                raise RuntimeError("JDE 登录失败")
            from workflows.cit_flow.cit_flow import CITFlow

            CITFlow(self.context).run()
        except Exception:
            page = browser.get_page()
            if page is not None:
                suffix = "_".join(filter(None, (info.get("short_name"), info.get("company"))))
                try:
                    page.screenshot(
                        path=str(config.get_path(
                            "screenshot", f"jde_failure_{suffix}.png" if suffix else "jde_failure.png"
                        ))
                    )
                except Exception:
                    logger.warning("JDE错误截图保存失败")
            raise
        finally:
            self.context.close_all_browsers()
