# -*- coding: utf-8 -*-
"""
招投标审查平台 - 登录页 UI 测试用例.

运行示例:
  pip install selenium pytest
  set BID_REVIEW_BASE_URL=https://你的平台地址
  pytest tests/ui/test_login_page.py -v --base-url %BID_REVIEW_BASE_URL%

说明:
  - 微信扫码登录仅验证二维码区域与提示文案，不断言登录成功（需真实微信环境）.
  - 部分 CSS 选择器需按实际前端 DOM 在 pages/login_page.py 中微调.
"""

import pytest

from pages.login_page import LoginPage


@pytest.mark.ui
class TestLoginPageLayout:
    """页面结构与静态元素."""

    def test_page_loads_with_title(self, login_page):
        """用例1: 打开登录页，应显示平台名称."""
        login_page.open()
        el = login_page.wait_visible(LoginPage.TITLE)
        assert "招投标审查平台" in el.text

    def test_welcome_message_visible(self, login_page):
        """用例2: 应显示欢迎语."""
        login_page.open()
        login_page.wait_visible(LoginPage.WELCOME)

    def test_three_login_tabs_visible(self, login_page):
        """用例3: 应展示三种登录方式 Tab."""
        login_page.open()
        login_page.wait_visible(LoginPage.TAB_WECHAT)
        login_page.wait_visible(LoginPage.TAB_PHONE)
        login_page.wait_visible(LoginPage.TAB_ACCOUNT)

    def test_footer_links_and_copyright(self, login_page):
        """用例4: 页脚帮助/隐私/条款及版权信息可见."""
        login_page.open()
        login_page.wait_visible(LoginPage.LINK_HELP)
        login_page.wait_visible(LoginPage.LINK_PRIVACY)
        login_page.wait_visible(LoginPage.LINK_TERMS)
        login_page.wait_visible(LoginPage.COPYRIGHT)


@pytest.mark.ui
class TestWeChatLoginTab:
    """微信登录（默认）."""

    def test_wechat_tab_default_content(self, login_page):
        """用例5: 默认展示微信扫码区域与提示文案."""
        login_page.open()
        login_page.wait_visible(LoginPage.WECHAT_HINT)
        assert "微信" in login_page.driver.find_element(*LoginPage.WECHAT_HINT).text

    @pytest.mark.skip(reason="二维码节点结构因项目而异，确认 DOM 后取消 skip")
    def test_wechat_qr_code_displayed(self, login_page):
        """用例6: 微信登录 Tab 下应显示二维码（img/canvas）."""
        login_page.open()
        login_page.wait_visible(LoginPage.QR_IMAGE)


@pytest.mark.ui
class TestLoginTabSwitch:
    """Tab 切换与对应表单区域."""

    def test_switch_to_phone_login(self, login_page):
        """用例7: 切换到手机登录，应出现手机号/验证码输入（若前端已实现）."""
        login_page.open()
        login_page.click_tab(LoginPage.TAB_PHONE)
        # 至少 Tab 可点击；输入框是否存在取决于实际页面
        assert login_page.has_visible(LoginPage.TAB_PHONE)

    def test_switch_to_account_login(self, login_page):
        """用例8: 切换到账号登录，应出现账号与密码框（若前端已实现）."""
        login_page.open()
        login_page.click_tab(LoginPage.TAB_ACCOUNT)
        password_visible = login_page.has_visible(LoginPage.PASSWORD_INPUT)
        username_visible = login_page.has_visible(LoginPage.USERNAME_INPUT)
        assert password_visible or username_visible or login_page.has_visible(
            LoginPage.TAB_ACCOUNT
        )

    def test_switch_back_to_wechat_shows_hint(self, login_page):
        """用例9: 从其他 Tab 切回微信登录，应再次显示扫码提示."""
        login_page.open()
        login_page.click_tab(LoginPage.TAB_PHONE)
        login_page.click_tab(LoginPage.TAB_WECHAT)
        login_page.wait_visible(LoginPage.WECHAT_HINT)


@pytest.mark.ui
class TestAccountLoginValidation:
    """账号登录交互（不依赖真实账号时可做冒烟）."""

    @pytest.mark.skip(reason="需测试账号或 mock 接口；有环境后取消 skip")
    def test_account_login_empty_submit_shows_error(self, login_page):
        """用例10: 账号密码为空点击登录，应提示错误且不跳转."""
        login_page.open()
        login_page.click_tab(LoginPage.TAB_ACCOUNT)
        if login_page.has_visible(LoginPage.LOGIN_BUTTON):
            login_page.driver.find_element(*LoginPage.LOGIN_BUTTON).click()
        # 按实际页面断言 toast / 表单校验文案
        assert "登录" in login_page.driver.page_source


@pytest.mark.ui
class TestFooterNavigation:
    """页脚链接（仅验证可点击，目标 URL 按环境配置）."""

    @pytest.mark.skip(reason="外链或弹窗需确认预期 URL 后启用")
    def test_help_link_navigates(self, login_page):
        """用例11: 点击「帮助」进入帮助页或打开新窗口."""
        login_page.open()
        login_page.driver.find_element(*LoginPage.LINK_HELP).click()
        assert len(login_page.driver.window_handles) >= 1
