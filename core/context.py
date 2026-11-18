#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务上下文
"""
from typing import Any, Dict


class TaskContext:
    """任务上下文，用于在任务之间共享状态。"""

    def __init__(self):
        self._data: Dict[str, Any] = {}
        self.browsers: Dict[str, Any] = {}
        self.clients: Dict[str, Any] = {}
        self.config = None

    def set(self, key: str, value: Any) -> None:
        self._data[key] = value

    def get(self, key: str, default: Any = None) -> Any:
        return self._data.get(key, default)

    def set_browser(self, system_name: str, browser: Any) -> None:
        self.browsers[system_name] = browser

    def get_browser(self, system_name: str) -> Any:
        return self.browsers.get(system_name)

    def clear(self) -> None:
        self._data.clear()

    def close_all_browsers(self) -> None:
        for browser in self.browsers.values():
            try:
                if browser and hasattr(browser, 'close'):
                    browser.close()
            except Exception:
                pass
        self.browsers.clear()
