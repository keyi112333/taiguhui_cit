#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业所得税计税 RPA 框架入口。
"""
import time
from pathlib import Path

from config import config
from config.config import SYSTEM_CONFIG_PATH, USER_CONFIG_PATH
from init.init import initEnv
from log.loggingConfig import logger
from utils.helpers import format_time
from utils.period_dialog import ask_accounting_period
from workflows.cit_flow.main_task import CITMainTask


def main():
    """主函数：加载配置、输入会计期间、初始化环境并执行主任务。"""
    flow_start_time = time.time()

    config_path = Path(USER_CONFIG_PATH)
    if not config_path.is_absolute():
        config_path = Path(__file__).resolve().parent / config_path
    if config_path.exists():
        system_path = Path(SYSTEM_CONFIG_PATH)
        if not system_path.is_absolute():
            system_path = Path(__file__).resolve().parent / system_path
        config.load(str(config_path), str(system_path))
        config.ensure_directories()
    else:
        logger.error(f"配置文件不存在，请检查路径是否正确: {config_path}")
        raise Exception(f"配置文件不存在，请检查路径是否正确: {config_path}")

    period = ask_accounting_period()
    config.set("企业所得税配置.会计期间", period or "")
    if period:
        logger.info(f"本次运行会计期间：{period}（弹窗输入）")
    else:
        logger.info("未输入会计期间，按自然月运行")

    initEnv()

    logger.info(f"{config.project_name}开始执行")
    main_task = CITMainTask(config)

    if main_task.run():
        elapsed_time = time.time() - flow_start_time
        logger.info(f"{config.project_name}完成执行, 总耗时: {format_time(elapsed_time)}")
        return True

    elapsed_time = time.time() - flow_start_time
    logger.error(
        f"{config.project_name}执行失败，状态: {main_task.status}, "
        f"总耗时: {format_time(elapsed_time)}"
    )
    return False


if __name__ == '__main__':
    raise SystemExit(0 if main() else 1)
