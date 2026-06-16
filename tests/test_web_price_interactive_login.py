from __future__ import annotations

import json
from pathlib import Path

import pytest

from ocr_price import web_price


class _FakeLocator:
    def __init__(self, count: int = 0) -> None:
        self._count = count
        self.first = self

    def count(self) -> int:
        return self._count

    def fill(self, value: str) -> None:
        return None


class _FakePage:
    def __init__(self, visible_login: bool = False, visible_inputs: bool = False) -> None:
        self.waits: list[int] = []
        self.urls: list[str] = []
        self.visible_login = visible_login
        self.visible_inputs = visible_inputs

    def goto(self, url: str, **kwargs) -> None:
        self.urls.append(url)

    def wait_for_timeout(self, timeout_ms: int) -> None:
        self.waits.append(timeout_ms)

    def locator(self, selector: str) -> _FakeLocator:
        if selector == ".topbar-nav-login:visible":
            return _FakeLocator(1 if self.visible_login else 0)
        if "请输入用户名" in selector or "请输入密码" in selector:
            return _FakeLocator(1 if self.visible_inputs else 0)
        return _FakeLocator(0)

    def title(self) -> str:
        return "登录失败页"

    def content(self) -> str:
        return "<html>login failed</html>"

    def screenshot(self, path: str, full_page: bool = True) -> None:
        Path(path).write_bytes(b"png")


def test_login_waits_for_manual_login_when_auto_login_fails(monkeypatch):
    page = _FakePage()
    states = [
        (False, "未检测到登录态"),
        (False, "未检测到登录态"),
        (False, "未检测到登录态"),
        (True, "userinfo接口返回已登录"),
    ]

    def fake_detect_logged_in(page_arg, retry=False):
        return states.pop(0)

    monkeypatch.setattr(web_price, "_detect_logged_in", fake_detect_logged_in)

    proof = web_price._login(
        page,
        "user",
        "pwd",
        allow_manual_login=True,
        manual_login_timeout_seconds=2,
        manual_login_poll_interval_seconds=1,
    )

    assert "人工登录成功" in proof
    assert 1000 in page.waits


def test_login_does_not_wait_for_manual_login_when_disabled(monkeypatch):
    page = _FakePage()

    def fake_detect_logged_in(page_arg, retry=False):
        return False, "未检测到登录态"

    monkeypatch.setattr(web_price, "_detect_logged_in", fake_detect_logged_in)

    with pytest.raises(web_price.WebPriceError, match="登录失败"):
        web_price._login(
            page,
            "user",
            "pwd",
            allow_manual_login=False,
            manual_login_timeout_seconds=2,
            manual_login_poll_interval_seconds=1,
        )


def test_login_falls_back_to_manual_when_auto_submit_fails(monkeypatch):
    page = _FakePage(visible_login=True, visible_inputs=True)
    states = [
        (False, "未检测到登录态"),
        (False, "未检测到登录态"),
        (False, "未检测到登录态"),
        (True, "userinfo接口返回已登录"),
    ]
    click_results = [True, True, False]

    def fake_detect_logged_in(page_arg, retry=False):
        return states.pop(0)

    def fake_click_any(page_arg, selectors, timeout_ms=5000):
        return click_results.pop(0)

    monkeypatch.setattr(web_price, "_detect_logged_in", fake_detect_logged_in)
    monkeypatch.setattr(web_price, "_click_any", fake_click_any)

    proof = web_price._login(
        page,
        "user",
        "pwd",
        allow_manual_login=True,
        manual_login_timeout_seconds=2,
        manual_login_poll_interval_seconds=1,
    )

    assert "人工登录成功" in proof


def test_login_failure_writes_diagnostics(monkeypatch, tmp_path: Path):
    page = _FakePage()

    def fake_detect_logged_in(page_arg, retry=False):
        return False, "未检测到登录态"

    monkeypatch.setattr(web_price, "_detect_logged_in", fake_detect_logged_in)

    with pytest.raises(web_price.WebPriceError):
        web_price._login(
            page,
            "user",
            "pwd",
            allow_manual_login=False,
            manual_login_timeout_seconds=1,
            manual_login_poll_interval_seconds=1,
            diagnostics_dir=tmp_path,
        )

    summary = json.loads((tmp_path / "login_diagnostics.json").read_text(encoding="utf-8"))
    assert summary["reason"] == "未检测到登录态"
    assert (tmp_path / "login_failure.png").exists()
    assert (tmp_path / "login_failure.html").exists()


def test_login_force_manual_waits_for_user_confirmation(monkeypatch):
    page = _FakePage()
    states = [
        (True, "顶部导航已显示登录态标记"),
    ]
    prompts: list[str] = []

    def fake_detect_logged_in(page_arg, retry=False):
        return states.pop(0)

    def fake_input(prompt):
        prompts.append(prompt)
        return ""

    monkeypatch.setattr(web_price, "_detect_logged_in", fake_detect_logged_in)
    monkeypatch.setattr(web_price, "input", fake_input, raising=False)

    proof = web_price._login(
        page,
        "user",
        "pwd",
        allow_manual_login=True,
        manual_login_timeout_seconds=2,
        manual_login_poll_interval_seconds=1,
        force_manual_login=True,
    )

    assert "人工确认登录完成" in proof
    assert prompts


def test_fetch_web_prices_falls_back_from_cdp_to_persistent_browser(monkeypatch, tmp_path: Path):
    calls: list[str] = []
    page = _FakePage()

    class _FakePlaywright:
        chromium = object()

    class _FakeSyncPlaywright:
        def __enter__(self):
            return _FakePlaywright()

        def __exit__(self, exc_type, exc, tb):
            return False

    class _FakeBrowser:
        def close(self):
            calls.append("close")

    def fake_sync_playwright():
        return _FakeSyncPlaywright()

    def fake_launch_browser_with_state(p, headless, chrome_cdp_url=None):
        calls.append(f"launch:{headless}")
        return _FakeBrowser(), _FakeBrowser(), page, "cdp" if len(calls) == 1 else "persistent"

    def fake_login(page_arg, user, pwd, allow_manual_login, manual_login_timeout_seconds, diagnostics_dir=None, force_manual_login=False):
        calls.append(f"login:{allow_manual_login}")
        if len([x for x in calls if x.startswith("login:")]) == 1:
            raise web_price.WebPriceError("CDP登录态失效")
        return "人工登录成功"

    monkeypatch.setattr(web_price, "sync_playwright", fake_sync_playwright)
    monkeypatch.setattr(web_price, "_parse_credentials", lambda *args, **kwargs: ("u", "p"))
    monkeypatch.setattr(web_price, "_launch_browser_with_state", fake_launch_browser_with_state)
    monkeypatch.setattr(web_price, "_login", fake_login)
    monkeypatch.setattr(web_price, "_latest_url_from_list", lambda *args, **kwargs: "https://example.com/detail")
    monkeypatch.setattr(web_price, "_extract_detail_rows", lambda page_arg: [{"钢厂/产地": "测试钢厂"}])
    monkeypatch.setattr(web_price, "_parse_date", lambda text: "2026-06-16")
    monkeypatch.setattr(web_price, "_export_raw", lambda rows, output_excel: output_excel.write_text("raw", encoding="utf-8"))
    monkeypatch.setattr(web_price, "_source_mills", lambda rows: ["测试钢厂"])

    report = web_price.fetch_web_prices(
        project_excel=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        location="安徽合肥",
        list_url="https://example.com/list",
        detail_url=None,
        account_file=tmp_path / "account.txt",
        username=None,
        password=None,
        output_excel=tmp_path / "raw.xlsx",
        report_out=tmp_path / "report.json",
        headless=True,
    )

    assert report["login_proof"] == "人工登录成功"
    assert calls[:4] == ["launch:True", "login:False", "close", "launch:False"]


def test_fetch_web_prices_retries_after_encrypted_detail_with_manual_login(monkeypatch, tmp_path: Path):
    calls: list[str] = []
    page = _FakePage()

    class _FakePlaywright:
        chromium = object()

    class _FakeSyncPlaywright:
        def __enter__(self):
            return _FakePlaywright()

        def __exit__(self, exc_type, exc, tb):
            return False

    class _FakeBrowser:
        def close(self):
            calls.append("close")

    extract_calls = {"count": 0}

    def fake_extract_detail_rows(page_arg):
        extract_calls["count"] += 1
        if extract_calls["count"] == 1:
            raise web_price.WebPriceError("详情页价格仍为加密态（icon-key存在），请检查登录状态")
        return [{"钢厂/产地": "测试钢厂"}]

    def fake_login(page_arg, user, pwd, allow_manual_login, manual_login_timeout_seconds, diagnostics_dir=None, force_manual_login=False):
        calls.append(f"login:{force_manual_login}")
        return "登录成功"

    monkeypatch.setattr(web_price, "sync_playwright", lambda: _FakeSyncPlaywright())
    monkeypatch.setattr(web_price, "_parse_credentials", lambda *args, **kwargs: ("u", "p"))
    monkeypatch.setattr(web_price, "_launch_browser_with_state", lambda *args, **kwargs: (_FakeBrowser(), _FakeBrowser(), page, "persistent"))
    monkeypatch.setattr(web_price, "_login", fake_login)
    monkeypatch.setattr(web_price, "_latest_url_from_list", lambda *args, **kwargs: "https://example.com/detail")
    monkeypatch.setattr(web_price, "_extract_detail_rows", fake_extract_detail_rows)
    monkeypatch.setattr(web_price, "_parse_date", lambda text: "2026-06-16")
    monkeypatch.setattr(web_price, "_export_raw", lambda rows, output_excel: output_excel.write_text("raw", encoding="utf-8"))
    monkeypatch.setattr(web_price, "_source_mills", lambda rows: ["测试钢厂"])

    report = web_price.fetch_web_prices(
        project_excel=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        location="安徽合肥",
        list_url="https://example.com/list",
        detail_url=None,
        account_file=tmp_path / "account.txt",
        username=None,
        password=None,
        output_excel=tmp_path / "raw.xlsx",
        report_out=tmp_path / "report.json",
        headless=False,
    )

    assert report["row_count"] == 1
    assert calls[:2] == ["login:False", "login:True"]


def test_fetch_web_prices_passes_chrome_cdp_url_to_launcher(monkeypatch, tmp_path: Path):
    seen: dict[str, str | None] = {}
    page = _FakePage()

    class _FakePlaywright:
        chromium = object()

    class _FakeSyncPlaywright:
        def __enter__(self):
            return _FakePlaywright()

        def __exit__(self, exc_type, exc, tb):
            return False

    class _FakeBrowser:
        def close(self):
            return None

    def fake_launch_browser_with_state(p, headless, chrome_cdp_url=None):
        seen["chrome_cdp_url"] = chrome_cdp_url
        return _FakeBrowser(), _FakeBrowser(), page, "cdp"

    monkeypatch.setattr(web_price, "sync_playwright", lambda: _FakeSyncPlaywright())
    monkeypatch.setattr(web_price, "_parse_credentials", lambda *args, **kwargs: ("u", "p"))
    monkeypatch.setattr(web_price, "_launch_browser_with_state", fake_launch_browser_with_state)
    monkeypatch.setattr(web_price, "_login", lambda *args, **kwargs: "已登录")
    monkeypatch.setattr(web_price, "_latest_url_from_list", lambda *args, **kwargs: "https://example.com/detail")
    monkeypatch.setattr(web_price, "_extract_detail_rows", lambda page_arg: [{"钢厂/产地": "测试钢厂"}])
    monkeypatch.setattr(web_price, "_parse_date", lambda text: "2026-06-16")
    monkeypatch.setattr(web_price, "_export_raw", lambda rows, output_excel: output_excel.write_text("raw", encoding="utf-8"))
    monkeypatch.setattr(web_price, "_source_mills", lambda rows: ["测试钢厂"])

    web_price.fetch_web_prices(
        project_excel=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        location="安徽合肥",
        list_url="https://example.com/list",
        detail_url=None,
        account_file=tmp_path / "account.txt",
        username=None,
        password=None,
        output_excel=tmp_path / "raw.xlsx",
        report_out=tmp_path / "report.json",
        headless=True,
        chrome_cdp_url="http://127.0.0.1:9222",
    )

    assert seen["chrome_cdp_url"] == "http://127.0.0.1:9222"
