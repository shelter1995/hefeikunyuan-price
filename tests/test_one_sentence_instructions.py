from pathlib import Path


def test_agents_doc_defines_one_sentence_user_entrypoint():
    text = Path("AGENTS.md").read_text(encoding="utf-8")

    assert "小白用户一句话入口" in text
    assert "更新景明苑报价" in text
    assert "用户不需要提供命令行参数" in text
    assert "先 dry-run，报告 Manifest，等待用户确认" in text


def test_quote_update_skill_defines_natural_language_flow():
    text = Path("skills/quote-update/SKILL.md").read_text(encoding="utf-8")

    assert "小白用户一句话入口" in text
    assert "只更新网价" in text
    assert "批量更新报价" in text
    assert "如果匹配到多个项目文件" in text
    assert "禁止要求用户自己拼命令" in text
    assert "确认映射后禁止重新 dry-run" in text
    assert "apply_confirmations.py" in text
