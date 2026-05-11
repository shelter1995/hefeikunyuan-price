from __future__ import annotations

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
