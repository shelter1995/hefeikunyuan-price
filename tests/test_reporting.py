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


def test_render_single_report_markdown_lists_inventory_details():
    result = {
        "project": "项目报价/测试项目.xlsx",
        "mode": "both",
        "started_at": "2026-05-13T09:00:00",
        "ended_at": "2026-05-13T09:02:00",
        "status": "ok",
        "image_doc": {
            "status": "ok",
            "apply_summary": {
                "updated_count": 0,
                "skipped_count": 0,
                "updated_items": [],
                "skipped_items": [],
            },
            "inventory_report": {
                "status": "ok",
                "cleared_count": 3,
                "applied_count": 2,
                "applied": [
                    {
                        "mill": "桂鑫报价",
                        "sheet_mill": "桂鑫",
                        "warehouse": "厂内",
                        "product": "螺纹",
                        "spec": "12",
                        "length": "9",
                        "material": "HRB400E",
                        "status": "充足",
                        "cell": "P16",
                        "source_spec": "9米HRB400E螺纹12",
                        "confidence_basis": "MiniMax视觉库存表",
                    },
                    {
                        "mill": "贵航报价",
                        "sheet_mill": "贵航",
                        "warehouse": "",
                        "product": "圆钢",
                        "spec": "16",
                        "length": "",
                        "material": "",
                        "status": "告警",
                        "cell": "R18",
                        "source_spec": "圆钢16（少）",
                        "confidence_basis": "JSON库存字段",
                    },
                ],
            },
        },
    }

    markdown = render_single_report_markdown(result, json_report_path="运行产物/report.json")

    assert "### 3. 库存颜色标注明细" in markdown
    assert "| 厂家 | 仓库 | 材质 | 产品 | 规格 | 长度 | 库存情况 | 单元格 | 来源描述 |" in markdown
    assert "| 桂鑫 | 厂内 | HRB400E | 螺纹 | 12 | 9 | 充足 | P16 | 9米HRB400E螺纹12；MiniMax视觉库存表 |" in markdown
    assert "| 贵航 |  |  | 圆钢 | 16 |  | 告警 | R18 | 圆钢16（少）；JSON库存字段 |" in markdown
    assert "蓝色（充足）" not in markdown
    assert "新标注颜色：" not in markdown


def test_render_single_report_markdown_shows_pending_and_inventory_error():
    result = {
        "project": "项目报价/测试项目.xlsx",
        "mode": "both",
        "started_at": "2026-05-23T09:00:00",
        "ended_at": "2026-05-23T09:01:00",
        "status": "pending_confirmation",
        "web": {
            "status": "pending_confirmation",
            "reason": "网价对照存在待确认项",
            "pending_mapping_json": "运行产物/厂家对照表_安徽合肥_待确认.json",
            "pending_details": {
                "pending_matches": [
                    {
                        "项目文件Sheet": "徐钢",
                        "最新清单厂家Sheet": "徐钢",
                        "状态": "待确认匹配",
                        "说明": "需人工确认",
                    }
                ],
                "pending_new": [],
            },
        },
        "image_doc": {
            "status": "ok",
            "apply_summary": {
                "updated_items": [],
                "skipped_items": [],
            },
            "inventory_report": {
                "status": "error",
                "error": "报价表sheet不存在",
            },
        },
    }

    markdown = render_single_report_markdown(result, json_report_path="运行产物/report.json")

    assert "待确认事项" in markdown
    assert "网价对照存在待确认项" in markdown
    assert "厂家对照表_安徽合肥_待确认.json" in markdown
    assert "徐钢" in markdown
    assert "库存颜色标注异常" in markdown
    assert "报价表sheet不存在" in markdown


def test_render_single_report_markdown_shows_inventory_review_conflicts():
    result = {
        "project": "项目报价/测试项目.xlsx",
        "mode": "image_doc",
        "started_at": "2026-05-30T09:00:00",
        "ended_at": "2026-05-30T09:01:00",
        "status": "ok",
        "image_doc": {
            "status": "ok",
            "apply_summary": {
                "updated_items": [],
                "skipped_items": [],
            },
            "inventory_report": {
                "status": "review_only",
                "reason": "dry-run模式不修改库存颜色",
                "review": {
                    "raw_count": 2,
                    "selected_count": 1,
                    "duplicate_group_count": 1,
                    "conflict_group_count": 1,
                    "conflict_groups": [
                        {
                            "company": "徐钢",
                            "warehouse": "",
                            "product": "螺纹",
                            "spec": "12",
                            "length": "9",
                            "material": "",
                            "statuses": ["充足", "告警"],
                            "selected": {
                                "status": "充足",
                                "source_file": "ocr价格提取_徐钢.json",
                                "source_spec": "9米螺纹12E",
                            },
                        }
                    ],
                },
            },
        },
    }

    markdown = render_single_report_markdown(result, json_report_path="运行产物/report.json")

    assert "### 3. 库存归并检查" in markdown
    assert "冲突组：1" in markdown
    assert "| 徐钢 |  | 螺纹 12 9 | 充足 / 告警 | 充足 |" in markdown
