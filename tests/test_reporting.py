from ocr_price.reporting import render_single_report_markdown


def test_render_single_report_markdown_contains_required_sections():
    result = {
        "project": "项目报价/测试项目.xlsx",
        "mode": "both",
        "started_at": "2026-05-11T09:00:00",
        "ended_at": "2026-05-11T09:01:00",
        "status": "ok",
        "web": {
            "status": "ok",
            "apply_summary": {
                "updated_count": 1,
                "skipped_count": 0,
                "updated_items": [
                    {
                        "项目文件Sheet": "徐钢",
                        "G1": {"old": None, "new": "网价[2026-05-11]"},
                        "G3": {"old": 3300, "new": 3400},
                        "G4": {"old": 3100, "new": 3200},
                    }
                ],
                "skipped_items": [],
            },
        },
        "image_doc": {
            "status": "ok",
            "apply_summary": {
                "updated_count": 1,
                "skipped_count": 0,
                "updated_items": [
                    {
                        "项目文件Sheet": "闽源",
                        "H1": {"old": None, "new": "报价[2026-05-11]"},
                        "H3": {"old": 3320, "new": 3450},
                        "H4": {"old": 3140, "new": 3180},
                    }
                ],
                "skipped_items": [],
            },
        },
    }

    markdown = render_single_report_markdown(result, json_report_path="运行产物/report.json")
    assert "项目报价单文件更新完成" in markdown
    assert "网价更新" in markdown
    assert "图片/文档价更新" in markdown
    assert "徐钢" in markdown
    assert "闽源" in markdown
    assert "运行产物/report.json" in markdown
