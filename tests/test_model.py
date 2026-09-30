from datetime import date, datetime
from pathlib import Path

import pytest

import trip_model as tm

VALID = """\
title: 测试旅行
start: 2026-08-02
end: 2026-08-03
travelers: 2
days:
  - date: 2026-08-02
    label: 苏州 → 上海
    items:
      - title: 高铁
        type: 交通
        start: "07:28"
        end: "07:59"
        price: 38
        qty: 2
        booking_at: 2026-08-01 20:00
        status: 已购买
        place: 上海虹桥站
      - title: 夜车
        type: 交通
        start: "22:10"
        end: "06:30"
        cost: 300
  - date: 2026-08-03
    items: []
"""


def write(tmp_path: Path, text: str, name: str = "2026-test.yaml") -> Path:
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


def errors_for(tmp_path: Path, text: str) -> list[str]:
    with pytest.raises(tm.TripLoadError) as info:
        tm.load_trip(write(tmp_path, text))
    return info.value.messages


def test_valid_trip_loads(tmp_path):
    loaded = tm.load_trip(write(tmp_path, VALID))
    assert loaded.id == "2026-test"
    item = loaded.trip.days[0].items[0]
    assert item.start == "07:28"
    assert item.amount == 76
    assert item.booking_at == datetime(2026, 8, 1, 20, 0)
    assert loaded.trip.days[0].items[1].amount == 300
    assert loaded.trip.days[1].date == date(2026, 8, 3)


def test_date_only_booking(tmp_path):
    text = VALID.replace("booking_at: 2026-08-01 20:00", "booking_at: 2026-08-01")
    item = tm.load_trip(write(tmp_path, text)).trip.days[0].items[0]
    assert item.booking_at == date(2026, 8, 1)
    assert not isinstance(item.booking_at, datetime)


def test_defaults_and_blank_strings(tmp_path):
    text = VALID.replace("place: 上海虹桥站", 'place: "  "')
    trip = tm.load_trip(write(tmp_path, text)).trip
    assert trip.days[0].items[0].place is None
    assert trip.days[0].items[1].status == "无需预订"


def test_explicit_null_optional_fields(tmp_path):
    text = VALID.replace('start: "07:28"', "start:").replace('end: "07:59"', 'end: ""')
    item = tm.load_trip(write(tmp_path, text)).trip.days[0].items[0]
    assert item.start is None and item.end is None


def test_todo_flag(tmp_path):
    text = VALID.replace("status: 已购买", "status: 待购买")
    item = tm.load_trip(write(tmp_path, text)).trip.days[0].items[0]
    assert item.is_todo


def test_unknown_field_rejected(tmp_path):
    messages = errors_for(tmp_path, VALID.replace("place:", "palce:"))
    assert any("days[0].items[0].palce" in m for m in messages)


def test_numeric_time_rejected(tmp_path):
    messages = errors_for(tmp_path, VALID.replace('start: "07:28"', "start: 728"))
    assert any("days[0].items[0].start" in m for m in messages)


def test_bad_time_format_rejected(tmp_path):
    messages = errors_for(tmp_path, VALID.replace('"07:28"', '"7:28"'))
    assert any("days[0].items[0].start" in m for m in messages)


def test_invalid_enum_rejected(tmp_path):
    messages = errors_for(tmp_path, VALID.replace("status: 已购买", "status: 已买"))
    assert any("days[0].items[0].status" in m for m in messages)


def test_cost_with_price_rejected(tmp_path):
    messages = errors_for(tmp_path, VALID.replace("cost: 300", "cost: 300\n        price: 1"))
    assert any("days[0].items[1]" in m and "cost" in m for m in messages)


def test_day_out_of_range_rejected(tmp_path):
    messages = errors_for(tmp_path, VALID.replace("- date: 2026-08-03", "- date: 2026-08-09"))
    assert any("超出" in m for m in messages)


def test_days_must_increase(tmp_path):
    messages = errors_for(tmp_path, VALID.replace("- date: 2026-08-03", "- date: 2026-08-02"))
    assert any("递增" in m for m in messages)


def test_bad_trip_id_rejected(tmp_path):
    with pytest.raises(tm.TripLoadError) as info:
        tm.load_trip(write(tmp_path, VALID, name="Bad_Name.yaml"))
    assert any("文件名" in m for m in info.value.messages)


def test_yaml_syntax_error_reports_line(tmp_path):
    messages = errors_for(tmp_path, "title: [unclosed\n")
    assert any("行" in m for m in messages)


def test_load_trips_aggregates_errors(tmp_path):
    write(tmp_path, VALID, "2026-good.yaml")
    write(tmp_path, VALID.replace("status: 已购买", "status: 已买"), "2026-bad-a.yaml")
    write(tmp_path, "title: [x\n", "2026-bad-b.yaml")
    trips, errors = tm.load_trips(tmp_path)
    assert [t.id for t in trips] == ["2026-good"]
    assert any(e.startswith("2026-bad-a.yaml") for e in errors)
    assert any(e.startswith("2026-bad-b.yaml") for e in errors)


def test_dump_round_trip(tmp_path):
    trip = tm.load_trip(write(tmp_path, VALID)).trip
    text = tm.dump_trip(trip)
    assert 'start: "07:28"' in text
    assert "booking_at: 2026-08-01 20:00" in text
    assert "price: 38\n" in text
    again = tm.load_trip(write(tmp_path, text, "2026-again.yaml")).trip
    assert again == trip


def test_dump_has_no_yaml_anchors():
    shared = date(2026, 8, 2)
    trip = tm.Trip.model_validate({
        "title": "锚点", "start": shared, "end": shared,
        "days": [{"date": shared, "items": []}],
    })
    text = tm.dump_trip(trip)
    assert "&id" not in text and "*id" not in text


@pytest.mark.parametrize("content", [
    VALID.replace("- date: 2026-08-03", "- date: 2026-02-30"),
    "title: \x01\n",
])
def test_bad_input_reported_not_raised(tmp_path, content):
    messages = errors_for(tmp_path, content)
    assert messages and messages[0].startswith("2026-test.yaml")


def test_non_utf8_reported(tmp_path):
    path = tmp_path / "2026-gbk.yaml"
    path.write_bytes(VALID.encode("gbk"))
    with pytest.raises(tm.TripLoadError) as info:
        tm.load_trip(path)
    assert "UTF-8" in info.value.messages[0]


@pytest.mark.parametrize("name", ["assets.yaml", "icons.yaml"])
def test_reserved_trip_id_rejected(tmp_path, name):
    with pytest.raises(tm.TripLoadError) as info:
        tm.load_trip(write(tmp_path, VALID, name=name))
    assert "保留" in info.value.messages[0]


def test_yml_extension_reported(tmp_path):
    write(tmp_path, VALID, "2026-a.yml")
    _, errors = tm.load_trips(tmp_path)
    assert any("2026-a.yml" in e and ".yaml" in e for e in errors)


@pytest.mark.parametrize("value", [".inf", "1e999", "true"])
def test_bad_numbers_rejected(tmp_path, value):
    messages = errors_for(tmp_path, VALID.replace("price: 38", f"price: {value}"))
    assert any("days[0].items[0].price" in m for m in messages)
