import json
import re

import export_schema
import trip_model as tm
from conftest import ROOT


def test_committed_schema_is_current():
    committed = json.loads((ROOT / "schema" / "trip.schema.json").read_text(encoding="utf-8"))
    assert committed == export_schema.build_schema()


def test_every_property_has_description():
    schema = export_schema.build_schema()
    objects = [schema, *schema["$defs"].values()]
    for obj in objects:
        for name, prop in obj["properties"].items():
            assert prop.get("description"), name


def test_time_field_editor_hints():
    item = export_schema.build_schema()["$defs"]["Item"]["properties"]
    for key in ("start", "end"):
        assert re.match(item[key]["pattern"], "19:28")
        assert not re.match(item[key]["pattern"], "7:28")
        assert "23:50" in item[key]["examples"]
        assert len(item[key]["examples"]) == 144
    assert item["type"]["enum"] == list(tm.ITEM_TYPES)
    assert item["status"]["enum"] == list(tm.STATUSES)
    assert "anyOf" not in item["booking_at"]
    assert re.match(item["booking_at"]["pattern"], "2026-08-02 14:00")
    for key in ("price", "qty", "cost"):
        assert item[key]["type"] == "number" and item[key]["minimum"] == 0


def test_schema_has_only_standard_keywords():
    text = json.dumps(export_schema.build_schema())
    for keyword in ('"ge"', '"le"', '"gt"', '"lt"'):
        assert keyword not in text


def test_snippets_present():
    schema = export_schema.build_schema()
    labels = [s["label"] for s in schema["$defs"]["Day"]["properties"]["items"]["defaultSnippets"]]
    assert "新事项" in labels
    assert schema["properties"]["days"]["defaultSnippets"][0]["label"] == "新的一天"
