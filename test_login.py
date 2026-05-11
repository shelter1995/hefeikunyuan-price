#!/usr/bin/env python3
"""测试 mysteel.com 网站登录功能"""

import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')


def parse_credentials(account_file: Path) -> tuple[str, str]:
    """解析账号密码文件"""
    text = account_file.read_text(encoding="utf-8", errors="ignore") if account_file.exists() else ""
    user = pwd = ""
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if ("账号" in s or "用户名" in s) and ("：" in s or ":" in s):
            user = s.split("：")[-1].split(":")[-1].strip()
        if ("密码" in s or "pass" in s.lower()) and ("：" in s or ":" in s):
            pwd = s.split("：")[-1].split(":")[-1].strip()
    return user, pwd


def test_login():
    """测试登录流程"""
    print("=" * 60)
    print("测试 mysteel.com 网站登录")
    print("=" * 60)

    # 读取账号密码
    account_file = Path(__file__).parent / "网站账号密码.txt"
    user, pwd = parse_credentials(account_file)
    print(f"\n[1] 账号信息:")
    print(f"    - 用户名: {user}")
    print(f"    - 密码: {'*' * len(pwd)}")

    # 使用持久化用户数据目录
    user_data_dir = Path(__file__).parent / ".chrome_user_data"
    user_data_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n[2] 启动浏览器（有头模式）...")
    print(f"    - 用户数据目录: {user_data_dir}")

    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=False,
            locale="zh-CN",
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = browser.pages[0] if browser.pages else browser.new_page()
        print(f"[OK] 浏览器启动成功")

        # 访问网站
        print(f"\n[3] 访问 mysteel.com...")
        page.goto("https://www.mysteel.com/", timeout=60000, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        print(f"    - 页面标题: {page.title()}")

        # 检查是否已登录
        print(f"\n[4] 检查登录状态...")
        login_markers = [
            "a:has-text('退出')",
            ".topbar-nav-logout:visible",
            ".topbar-user-name:visible",
        ]

        already_logged = False
        for selector in login_markers:
            try:
                if page.locator(selector).count() > 0:
                    print(f"    [OK] 已登录，找到标记: {selector}")
                    already_logged = True
                    break
            except:
                pass

        if already_logged:
            print(f"\n[OK] 已处于登录状态，无需重新登录")
        else:
            print(f"    [INFO] 未登录，开始登录流程...")

            # 点击登录按钮
            print(f"\n[5] 点击登录按钮...")
            login_btn_selectors = [
                ".topbar-nav-login:visible",
                ".topbar-nav-login",
                "text=登录",
                "a:has-text('登录')",
            ]

            clicked = False
            for selector in login_btn_selectors:
                try:
                    if page.locator(selector).count() > 0:
                        page.locator(selector).first.click()
                        print(f"    [OK] 点击登录按钮: {selector}")
                        clicked = True
                        break
                except:
                    continue

            if not clicked:
                print(f"    [FAIL] 找不到登录按钮")
                browser.close()
                return

            page.wait_for_timeout(2000)

            # 切换到账号密码登录
            print(f"\n[6] 切换到账号密码登录...")
            account_tab_selectors = [
                ".form-tab-account:visible",
                ".form-tab-account",
                "text=账号登录",
                "button:has-text('账号登录')",
            ]

            for selector in account_tab_selectors:
                try:
                    if page.locator(selector).count() > 0:
                        page.locator(selector).first.click()
                        print(f"    [OK] 点击账号登录标签: {selector}")
                        break
                except:
                    continue

            page.wait_for_timeout(1000)

            # 填写账号密码
            print(f"\n[7] 填写账号密码...")
            try:
                user_input = page.locator('input[placeholder="请输入用户名"]:visible').first
                pwd_input = page.locator('input[placeholder="请输入密码"]:visible').first

                if user_input.count() == 0 or pwd_input.count() == 0:
                    print(f"    [FAIL] 找不到输入框")
                    browser.close()
                    return

                user_input.fill(user)
                pwd_input.fill(pwd)
                print(f"    [OK] 已填写账号密码")
            except Exception as e:
                print(f"    [FAIL] 填写失败: {e}")
                browser.close()
                return

            # 点击登录按钮
            print(f"\n[8] 点击登录按钮...")
            login_btn_selectors = [
                ".form-button-login:visible",
                ".form-button-login",
                "button:has-text('登录')",
            ]

            login_clicked = False
            for selector in login_btn_selectors:
                try:
                    if page.locator(selector).count() > 0:
                        page.locator(selector).first.click()
                        print(f"    [OK] 点击登录按钮: {selector}")
                        login_clicked = True
                        break
                except:
                    continue

            if not login_clicked:
                print(f"    [FAIL] 找不到登录按钮")
                browser.close()
                return

            # 等待登录完成
            print(f"\n[9] 等待登录完成...")
            page.wait_for_timeout(5000)

            # 验证登录结果
            print(f"\n[10] 验证登录结果...")
            for selector in login_markers:
                try:
                    if page.locator(selector).count() > 0:
                        print(f"    [OK] 登录成功，找到标记: {selector}")
                        already_logged = True
                        break
                except:
                    pass

            if not already_logged:
                print(f"    [FAIL] 登录失败，未找到登录标记")
                # 截图保存
                screenshot_path = "login_failed.png"
                page.screenshot(path=screenshot_path)
                print(f"    [INFO] 已保存截图: {screenshot_path}")
            else:
                print(f"\n[OK] 登录成功！")

        # 尝试访问价格页面验证
        if already_logged:
            print(f"\n[11] 访问价格页面验证...")
            page.goto(
                "https://jiancai.mysteel.com/market/pa228aa010101a0a01010715aaaa1.html",
                timeout=60000,
                wait_until="domcontentloaded"
            )
            page.wait_for_timeout(3000)
            print(f"    - 页面标题: {page.title()}")

            # 检查是否被重定向到登录页
            if "login" in page.url or "登录" in page.title():
                print(f"    [FAIL] 被重定向到登录页，登录可能未生效")
            else:
                print(f"    [OK] 成功访问价格页面")

        print(f"\n[INFO] 按 Enter 键关闭浏览器...")
        input()
        browser.close()
        print(f"[OK] 浏览器已关闭")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)


if __name__ == "__main__":
    test_login()
