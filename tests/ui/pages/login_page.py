# -*- coding: utf-8 -*-
"""Page object: 招投标审查平台 - 登录页."""

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class LoginPage:
    """登录页：微信登录 / 手机登录 / 账号登录."""

    PATH = "/"

    # 文案定位（与界面一致，便于无稳定 id 时使用）
    TITLE = (By.XPATH, "//*[contains(text(),'招投标审查平台')]")
    WELCOME = (By.XPATH, "//*[contains(text(),'您好，欢迎使用')]")
    TAB_WECHAT = (By.XPATH, "//*[contains(text(),'微信登录')]")
    TAB_PHONE = (By.XPATH, "//*[contains(text(),'手机登录')]")
    TAB_ACCOUNT = (By.XPATH, "//*[contains(text(),'账号登录')]")
    WECHAT_HINT = (By.XPATH, "//*[contains(text(),'使用微信扫一扫登录')]")
    LINK_HELP = (By.XPATH, "//a[contains(text(),'帮助')]|//*[contains(text(),'帮助')]")
    LINK_PRIVACY = (By.XPATH, "//a[contains(text(),'隐私')]|//*[contains(text(),'隐私')]")
    LINK_TERMS = (By.XPATH, "//a[contains(text(),'条款')]|//*[contains(text(),'条款')]")
    COPYRIGHT = (By.XPATH, "//*[contains(text(),'南方财经全媒体集团')]")

    # 手机登录 / 账号登录（切换后出现，选择器需按实际 DOM 调整）
    PHONE_INPUT = (
        By.CSS_SELECTOR,
        "input[type='tel'], input[placeholder*='手机'], input[name*='phone'], input[name*='mobile']",
    )
    SMS_CODE_INPUT = (
        By.CSS_SELECTOR,
        "input[placeholder*='验证码'], input[name*='code'], input[name*='captcha']",
    )
    USERNAME_INPUT = (
        By.CSS_SELECTOR,
        "input[type='text'][placeholder*='账号'], input[name*='user'], input[name*='account']",
    )
    PASSWORD_INPUT = (
        By.CSS_SELECTOR,
        "input[type='password']",
    )
    LOGIN_BUTTON = (
        By.XPATH,
        "//button[contains(.,'登录')]|//span[contains(.,'登录')]/ancestor::button",
    )
    QR_IMAGE = (By.CSS_SELECTOR, "img[src*='qr'], canvas, .qrcode img, .qr-code")

    def __init__(self, driver, base_url: str, timeout: int = 15):
        self.driver = driver
        self.base_url = base_url.rstrip("/")
        self.wait = WebDriverWait(driver, timeout)

    def open(self):
        self.driver.get(f"{self.base_url}{self.PATH}")

    def wait_visible(self, locator):
        return self.wait.until(EC.visibility_of_element_located(locator))

    def click_tab(self, locator):
        el = self.wait.until(EC.element_to_be_clickable(locator))
        el.click()

    def is_tab_active(self, tab_locator) -> bool:
        """通过父级 class 或下划线样式判断激活态（可按实际页面改写）."""
        el = self.driver.find_element(*tab_locator)
        cls = (el.get_attribute("class") or "") + (el.get_attribute("aria-selected") or "")
        parent_cls = ""
        try:
            parent_cls = el.find_element(By.XPATH, "./..").get_attribute("class") or ""
        except Exception:
            pass
        active_markers = ("active", "selected", "is-active", "ant-tabs-tab-active", "true")
        combined = f"{cls} {parent_cls}".lower()
        return any(m in combined for m in active_markers)

    def has_visible(self, locator) -> bool:
        try:
            el = self.driver.find_element(*locator)
            return el.is_displayed()
        except Exception:
            return False
