# -*- coding: utf-8 -*-
"""Selenium fixtures for UI tests."""

import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from pages.login_page import LoginPage


def pytest_addoption(parser):
    parser.addoption(
        "--base-url",
        action="store",
        default=os.environ.get("BID_REVIEW_BASE_URL", "http://localhost:8080"),
        help="招投标审查平台根地址，如 https://example.com",
    )
    parser.addoption(
        "--headless",
        action="store_true",
        default=os.environ.get("SELENIUM_HEADLESS", "").lower() in ("1", "true", "yes"),
        help="无头模式运行 Chrome",
    )


@pytest.fixture
def base_url(request):
    return request.config.getoption("--base-url").rstrip("/")


def _resolve_chrome_service() -> Service:
    """每次解析并安装与当前 Chrome 匹配的最新 chromedriver（可关闭缓存）."""
    override = os.environ.get("CHROMEDRIVER")
    if override:
        return Service(override)

    # webdriver-manager：cache_valid_range=0 表示不沿用本地缓存，每次向源站检查最新驱动
    from webdriver_manager.chrome import ChromeDriverManager

    path = ChromeDriverManager(cache_valid_range=0).install()
    return Service(path)


@pytest.fixture
def driver(request):
    options = Options()
    if request.config.getoption("--headless"):
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    service = _resolve_chrome_service()
    drv = webdriver.Chrome(service=service, options=options)

    drv.implicitly_wait(0)
    yield drv
    drv.quit()


@pytest.fixture
def login_page(driver, base_url):
    return LoginPage(driver, base_url)
