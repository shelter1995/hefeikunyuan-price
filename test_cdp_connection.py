#!/usr/bin/env python3
"""测试持久化浏览器状态方案

此脚本测试两种方案：
1. CDP 连接已运行的 Chrome（Chrome MCP）
2. 使用持久化用户数据目录启动浏览器（保持登录状态）
"""

import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')


def test_persistent_browser():
    """测试使用持久化用户数据目录启动浏览器"""
    print("=" * 60)
    print("测试持久化浏览器状态方案")
    print("=" * 60)

    user_data_dir = Path(__file__).parent / ".chrome_user_data"
    user_data_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n[1] 用户数据目录: {user_data_dir}")
    print(f"    此目录用于保存 Cookie 和登录状态")

    with sync_playwright() as p:
        print(f"\n[2] 启动浏览器（有头模式）...")

        # 使用持久化上下文启动浏览器
        browser = p.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=False,
            locale="zh-CN",
            args=["--disable-blink-features=AutomationControlled"],
        )

        page = browser.pages[0] if browser.pages else browser.new_page()
        print(f"[OK] 浏览器启动成功")

        # 访问 mysteel.com
        print(f"\n[3] 访问 mysteel.com...")
        page.goto("https://www.mysteel.com/", timeout=60000, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        print(f"    - 标题: {page.title()}")

        # 检测登录状态
        print(f"\n[4] 检测登录状态...")
        selectors = [
            "a:has-text('退出')",
            ".topbar-nav-logout:visible",
            ".topbar-user-name:visible",
            ".topbar-nav-news:visible",
            ".topbar-nav-msg:visible",
        ]

        logged_in = False
        for selector in selectors:
            try:
                count = page.locator(selector).count()
                if count > 0:
                    print(f"    [OK] 找到登录标记: {selector}")
                    logged_in = True
                    break
            except:
                pass

        if not logged_in:
            print(f"    [WARN] 未找到登录标记，需要手动登录")
            print(f"\n[!] 请在浏览器中手动登录 mysteel.com")
            print(f"    登录成功后，关闭浏览器即可保存状态")
            input(f"\n按 Enter 键关闭浏览器并保存状态...")
        else:
            print(f"    [OK] 已检测到登录状态!")
            print(f"\n[5] 尝试访问价格页面...")
            page.goto(
                "https://jiancai.mysteel.com/market/pa228aa010101a0a01010715aaaa1.html",
                timeout=60000,
                wait_until="domcontentloaded"
            )
            page.wait_for_timeout(3000)
            print(f"    [OK] 价格页面加载完成")
            print(f"    - 标题: {page.title()}")
            input(f"\n按 Enter 键关闭浏览器...")

        # 关闭浏览器时会自动保存状态到 user_data_dir
        browser.close()
        print(f"\n[OK] 浏览器已关闭，状态已保存到: {user_data_dir}")

    print("\n" + "=" * 60)
    print("测试完成!")
    print("=" * 60)
    print(f"\n说明:")
    print(f"1. 首次运行需要手动登录")
    print(f"2. 登录状态会保存在 {user_data_dir}")
    print(f"3. 下次运行时会自动加载登录状态")
    print(f"4. web_price.py 已集成此方案")


if __name__ == "__main__":
    test_persistent_browser()
