from ocr_price.events import PipelineEvent, append_event, event_to_dict


def test_event_to_dict_is_json_ready():
    event = PipelineEvent(stage="web", level="info", message="开始抓取", data={"location": "安徽合肥"})

    payload = event_to_dict(event)

    assert payload["stage"] == "web"
    assert payload["level"] == "info"
    assert payload["message"] == "开始抓取"
    assert payload["data"] == {"location": "安徽合肥"}
    assert "time" in payload


def test_append_event_keeps_result_events_list():
    result = {}

    append_event(result, stage="image_doc", level="blocked", message="存在待确认厂家")

    assert result["events"][0]["stage"] == "image_doc"
    assert result["events"][0]["level"] == "blocked"
    assert result["events"][0]["message"] == "存在待确认厂家"
