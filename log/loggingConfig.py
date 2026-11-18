#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日志初始化：控制台与 log_data/YYYYMMDD_机器名.log 双输出。
"""
import logging
import os
import socket
import sys
from datetime import datetime

from config.config import PYHON_LOG_FOLDER_PATH

today = datetime.now().strftime("%Y%m%d")
log_directory = os.path.expanduser(PYHON_LOG_FOLDER_PATH)
str_machine_name = socket.gethostname()

if not os.path.exists(log_directory):
    os.makedirs(log_directory)

logFilePath = os.path.join(log_directory, f"{today}_{str_machine_name}.log")
if not os.path.exists(logFilePath):
    with open(logFilePath, "w"):
        pass

logger = logging.getLogger("logger")
logger.setLevel(logging.DEBUG)
logger.propagate = False

if not logger.handlers:
    console_handler = logging.StreamHandler(stream=sys.stderr)
    file_handler = logging.FileHandler(logFilePath, encoding="utf-8")
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
