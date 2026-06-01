# 招投标审查平台 - 登录页 UI 测试

## 用例清单

| 编号 | 类 | 用例 | 说明 |
|------|-----|------|------|
| 1 | 布局 | `test_page_loads_with_title` | 平台标题 |
| 2 | 布局 | `test_welcome_message_visible` | 欢迎语 |
| 3 | 布局 | `test_three_login_tabs_visible` | 三个 Tab |
| 4 | 布局 | `test_footer_links_and_copyright` | 页脚链接与版权 |
| 5 | 微信 | `test_wechat_tab_default_content` | 扫码提示文案 |
| 6 | 微信 | `test_wechat_qr_code_displayed` | 二维码（需对齐 DOM 后启用） |
| 7 | 切换 | `test_switch_to_phone_login` | 手机登录 Tab |
| 8 | 切换 | `test_switch_to_account_login` | 账号登录 Tab |
| 9 | 切换 | `test_switch_back_to_wechat_shows_hint` | 切回微信 |
| 10 | 账号 | `test_account_login_empty_submit_shows_error` | 空提交校验（需环境） |
| 11 | 页脚 | `test_help_link_navigates` | 帮助链接（需环境） |

## 安装与运行

```bash
pip install -r tests/ui/requirements-ui.txt
```

设置平台地址（必填）：

```bash
# Windows
set BID_REVIEW_BASE_URL=https://你的招投标审查平台地址

# 运行全部 UI 用例
cd tests/ui
pytest test_login_page.py -v --base-url %BID_REVIEW_BASE_URL%

# 无头模式
pytest test_login_page.py -v --base-url %BID_REVIEW_BASE_URL% --headless
```

需本机已安装 **Chrome** 浏览器即可；**无需手动下载 ChromeDriver**。

### 驱动自动更新

每次启动浏览器（每个测试用例的 `driver` fixture）会通过 **webdriver-manager** 按本机 Chrome 版本检查并安装**最新匹配的 chromedriver**（`cache_valid_range=0`，不沿用旧缓存）。

- 若需固定使用本地驱动：设置环境变量 `CHROMEDRIVER=完整路径\chromedriver.exe`
- 首次或每次运行需能访问外网（下载驱动）

## 维护说明

- 选择器以界面文案为主；若前端使用 Vue/React 组件库，请在 `pages/login_page.py` 中改为稳定的 `data-testid`。
- 微信扫码登录无法自动化完成授权，仅做 UI 展示类断言。
