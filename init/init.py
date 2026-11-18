#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
初始化流程：关闭配置中指定的进程。
"""
import psutil

from config import config
from log.loggingConfig import logger


def set_resolution(in_dicConfig):
    """预留：按配置调整分辨率（当前主流程未调用）。"""
    import win32api
    import win32con

    width_value = in_dicConfig.get('Resolution_Width')
    height_value = in_dicConfig.get('Resolution_Height')

    if width_value and height_value and width_value != '' and height_value != '':
        try:
            width = int(width_value)
            height = int(height_value)
            logger.info(f"使用配置中的分辨率: {width}x{height}")
        except (ValueError, TypeError):
            logger.info("配置中的分辨率值无效，使用默认值1920x1080")
            width, height = 1920, 1080

        devmode = win32api.EnumDisplaySettings(None, win32con.ENUM_CURRENT_SETTINGS)
        devmode.PelsWidth = width
        devmode.PelsHeight = height
        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

        result = win32api.ChangeDisplaySettings(devmode, win32con.CDS_UPDATEREGISTRY)
        if result == win32con.DISP_CHANGE_SUCCESSFUL:
            logger.info(f"分辨率已更改为: {width}x{height}")
            return True
        logger.error("分辨率更改失败")
        return False

    logger.info("width或height配置不存在或为空，不设置")
    return True


def initEnv(list_process: list = []):
    """按进程名子串匹配关闭进程（配置项：基础配置.初始化进程列表）。"""
    logger.info("开始初始化环境")
    if not list_process:
        str_list_process = str(config.get("基础配置.初始化进程列表", "")).replace("，", ",")
        list_process = str_list_process.split(",")

    list_process = [name.strip() for name in list_process if name.strip()]

    for process_name in list_process:
        for process in psutil.process_iter():
            try:
                proc_name = process.name().lower()
                if process_name.lower() in proc_name:
                    if process.is_running():
                        process.terminate()
                        logger.info(f"已关闭进程： {process_name} (PID: {process.pid})")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                logger.error(f"关闭进程应用异常：{e}")
    logger.info("完成初始化环境")
