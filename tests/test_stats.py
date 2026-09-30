from datetime import date

import pytest

import trip_model as tm


def make_trip(items_by_day, travelers=2):
    return tm.Trip.model_validate({
        "title": "统计测试",
        "start": date(2026, 8, 1),
        "end": date(2026, 8, 2),
        "travelers": travelers,
        "days": [{"date": day, "items": items} for day, items in items_by_day],
    })


def test_compute_stats():
    trip = make_trip([
        (date(2026, 8, 1), [
            {"title": "高铁", "type": "交通", "price": 50, "qty": 2, "status": "已购买"},
            {"title": "酒店", "type": "住宿", "cost": 300, "status": "待购买"},
            {"title": "取消的门票", "type": "游览", "cost": 999, "status": "已取消"},
            {"title": "取消的预约", "type": "游览", "status": "待预约"},
        ]),
        (date(2026, 8, 2), [
            {"title": "午饭", "type": "餐饮", "cost": 100},
            {"title": "起床", "type": "日常"},
        ]),
    ])
    stats = tm.compute_stats(trip)
    assert stats.total == 500
    assert stats.per_person == 250
    assert stats.paid == 100
    assert stats.unpaid == 300
    assert stats.todo_count == 2
    assert stats.day_totals == {date(2026, 8, 1): 400, date(2026, 8, 2): 100}
    assert [(t, a) for t, a, _ in stats.by_type] == [("住宿", 300), ("交通", 100), ("餐饮", 100)]
    assert stats.by_type[0][2] == pytest.approx(0.6)


def test_cancelled_todo_not_counted():
    trip = make_trip([(date(2026, 8, 1), [
        {"title": "x", "type": "游览", "status": "已取消"},
    ])], travelers=None)
    stats = tm.compute_stats(trip)
    assert stats.total == 0
    assert stats.per_person is None
    assert stats.by_type == []
    assert stats.todo_count == 0
