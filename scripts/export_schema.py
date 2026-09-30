#!/usr/bin/env python3
"""从 trip_model 导出 schema/trip.schema.json，并补充 VS Code 编辑提示。"""

from __future__ import annotations

import json
from pathlib import Path

import trip_model as tm

SCHEMA_PATH = Path(__file__).resolve().parents[1] / "schema" / "trip.schema.json"
BOOKING_PATTERN = r"^\d{4}-\d{2}-\d{2}( ([01]\d|2[0-3]):[0-5]\d)?$"
TIME_EXAMPLES = [f"{h:02d}:{m:02d}" for h in range(24) for m in range(0, 60, 10)]

DESCRIPTIONS = {
    "Trip": {
        "title": "旅行标题",
        "start": "旅行开始日期，如 2026-08-02",
        "end": "旅行结束日期，不早于 start",
        "travelers": "人数，用于计算人均花销",
        "notes": "整体说明，如行程原则",
        "days": "按日期递增排列的每日行程",
    },
    "Day": {
        "date": "日期，必须在旅行起止日期内",
        "label": "当天标题，如“苏州 → 上海”；为空时显示日期",
        "items": "按执行时间排列的事项",
    },
    "Item": {
        "title": "事项名称，如 高铁、入住、早餐",
        "type": "类型",
        "start": "开始时间，带引号的 \"HH:MM\"，可手输精确时间如 \"19:28\"",
        "end": "结束时间，早于 start 表示跨过午夜",
        "note": "路线、车次、地址、营业时间等",
        "price": "单价（人民币）",
        "qty": "数量；消费 = price × qty",
        "cost": "总价，仅在无法拆成单价 × 数量时填写，不能与 price / qty 同时出现",
        "booking_at": "开售 / 预约时间，如 2026-08-01 或 2026-08-02 14:00",
        "status": "预订状态，默认“无需预订”",
        "place": "地图关键词，填写后页面显示地图按钮",
    },
}

ITEM_SNIPPET = {
    "label": "新事项",
    "description": "插入一条行程事项",
    "body": [{
        "title": "$1",
        "type": "${2|" + ",".join(tm.ITEM_TYPES) + "|}",
        "start": "^\"${3:09:00}\"",
        "end": "^\"${4:10:00}\"",
        "status": "${5|" + ",".join(tm.STATUSES) + "|}",
    }],
}
DAY_SNIPPET = {
    "label": "新的一天",
    "description": "插入一天及一条事项",
    "body": [{"date": "${1:2026-01-01}", "label": "$2", "items": [{"title": "$3", "type": "${4|" + ",".join(tm.ITEM_TYPES) + "|}"}]}],
}


def _drop_null(prop: dict) -> dict:
    """编辑器里不需要写 null，去掉 Optional 产生的 anyOf null 分支和 null 默认值。"""
    options = [o for o in prop.pop("anyOf", []) if o.get("type") != "null"]
    if len(options) == 1:
        prop.update(options[0])
    elif options:
        prop["anyOf"] = options
    if prop.get("default", 0) is None:
        del prop["default"]
    prop.pop("title", None)
    return prop


def build_schema() -> dict:
    schema = tm.Trip.model_json_schema()
    schema.pop("title", None)
    objects = {"Trip": schema, **schema["$defs"]}
    for name, obj in objects.items():
        obj.pop("title", None)
        for key, prop in obj["properties"].items():
            _drop_null(prop)
            prop["description"] = DESCRIPTIONS[name][key]
    item = schema["$defs"]["Item"]["properties"]
    for key in ("start", "end"):
        item[key].update(type="string", pattern=tm.TIME_PATTERN, examples=TIME_EXAMPLES)
    for key in ("price", "qty", "cost"):
        item[key] = {"type": "number", "minimum": 0, "description": DESCRIPTIONS["Item"][key]}
    item["booking_at"] = {
        "type": "string",
        "pattern": BOOKING_PATTERN,
        "description": DESCRIPTIONS["Item"]["booking_at"],
    }
    schema["$defs"]["Day"]["properties"]["items"]["defaultSnippets"] = [ITEM_SNIPPET]
    schema["properties"]["days"]["defaultSnippets"] = [DAY_SNIPPET]
    return {"$schema": "http://json-schema.org/draft-07/schema#", **schema}


def main() -> None:
    SCHEMA_PATH.parent.mkdir(exist_ok=True)
    SCHEMA_PATH.write_text(
        json.dumps(build_schema(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"已导出 {SCHEMA_PATH}")


if __name__ == "__main__":
    main()
