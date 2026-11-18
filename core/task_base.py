#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基础任务抽象类
"""
import time
from abc import ABC, abstractmethod

from log.loggingConfig import logger
from utils.helpers import format_time


class TaskBase(ABC):
    """所有任务继承此类，统一日志、状态与耗时管理。"""

    def __init__(self, name: str, context=None):
        self.name = name
        self.context = context
        self.status = "pending"
        self.result = None
        self.start_time = None
        self.end_time = None
        self.elapsed_time = None

    @abstractmethod
    def execute(self) -> bool:
        """执行任务（子类实现）。"""
        pass

    def run(self) -> bool:
        """运行任务并捕获异常。"""
        logger.info(f"开始任务: {self.name}")
        self.status = "running"
        self.start_time = time.time()
        try:
            self.result = self.execute()
            self.status = "completed"
            self.end_time = time.time()
            self.elapsed_time = self.end_time - self.start_time
            logger.info(f"完成任务: {self.name}, 耗时: {format_time(self.elapsed_time)}")
            return True
        except Exception as e:
            self.status = "failed"
            self.end_time = time.time()
            self.elapsed_time = self.end_time - self.start_time
            logger.error(f"任务失败: {self.name} - {e}, 耗时: {format_time(self.elapsed_time)}")
            return False

    def get_status(self) -> str:
        return self.status
