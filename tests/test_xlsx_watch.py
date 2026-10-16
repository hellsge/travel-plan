import re

import pytest

import trip_xlsx as tx
from test_xlsx_sync import TRIP_ID, edited, make_trip, put_yaml, xlsx_path, yaml_path


def test_events_coalesce_into_one():
    debouncer = tx.Debouncer()
    for now in (0.0, 0.5, 1.0):
        debouncer.notify("a", now)
    assert debouncer.due(2.4) == []
    assert debouncer.due(2.5) == ["a"]
    assert debouncer.due(10.0) == []


def test_not_due_before_quiet_period():
    debouncer = tx.Debouncer()
    debouncer.notify("a", 0.0)
    debouncer.notify("b", 1.0)
    assert debouncer.due(1.49) == []
    assert debouncer.due(1.5) == ["a"]
    assert debouncer.due(2.5) == ["b"]


def test_retry_waits_five_seconds():
    debouncer = tx.Debouncer()
    debouncer.retry_later("a", 10.0)
    assert debouncer.due(14.9) == []
    assert debouncer.due(15.0) == ["a"]


@pytest.mark.parametrize("name, expected", [
    ("2026-a.yaml", "2026-a"),
    ("2026-a.xlsx", "2026-a"),
    ("~$2026-a.xlsx", None),
    (".2026-a.k3j9.tmp", None),
    (".2026-a.xlsx", None),
    ("2026-a.txt", None),
    ("2026-a.yaml.bak", None),
])
def test_trip_id_for_path(tmp_path, name, expected):
    assert tx.trip_id_for(tmp_path / name) == expected


def counting_sync(calls):
    def sync(trip_id, trips_dir):
        calls.append(trip_id)
        return tx.sync_trip(trip_id, trips_dir)
    return sync


def test_start_syncs_everything_and_own_writes_are_noop(tmp_path, capsys):
    put_yaml(tmp_path, make_trip())
    put_yaml(tmp_path, edited(), trip_id="2026-other")
    calls = []
    loop = tx.WatchLoop(tmp_path, sync=counting_sync(calls))
    loop.start(0.0)
    assert calls == ["2026-other", TRIP_ID]
    assert xlsx_path(tmp_path).exists() and xlsx_path(tmp_path, "2026-other").exists()
    capsys.readouterr()

    # 自身写入触发的事件合并为一次 sync，按 digest 判定为已同步，不写文件也不打印。
    before = xlsx_path(tmp_path).read_bytes()
    loop.on_path(yaml_path(tmp_path))
    loop.on_path(xlsx_path(tmp_path))
    loop.on_path(tmp_path / f"~${TRIP_ID}.xlsx")
    loop.tick(0.1)
    loop.tick(2.0)
    assert calls == ["2026-other", TRIP_ID, TRIP_ID]
    assert xlsx_path(tmp_path).read_bytes() == before
    assert capsys.readouterr().out == ""

    # 之后只是读文件之类的事件，文件没变，不再调用 sync。
    loop.on_path(xlsx_path(tmp_path))
    loop.tick(3.0)
    loop.tick(5.0)
    assert calls == ["2026-other", TRIP_ID, TRIP_ID]


def test_external_edit_is_synced_and_reported(tmp_path, capsys):
    put_yaml(tmp_path, make_trip())
    loop = tx.WatchLoop(tmp_path)
    loop.start(0.0)
    capsys.readouterr()

    put_yaml(tmp_path, edited())
    loop.on_path(yaml_path(tmp_path))
    loop.tick(0.1)
    loop.tick(1.0)
    assert tx.digest(tx.parse_workbook(xlsx_path(tmp_path))) == tx.digest(make_trip())
    loop.tick(2.0)
    assert tx.digest(tx.parse_workbook(xlsx_path(tmp_path))) == tx.digest(edited())
    out = capsys.readouterr().out
    assert re.search(rf"^\d\d:\d\d:\d\d {TRIP_ID}: {tx.TO_XLSX}，.*总额", out, re.M)


def test_locked_file_is_retried(tmp_path, monkeypatch, capsys):
    put_yaml(tmp_path, make_trip())
    real_replace = tx.os.replace

    def locked(src, dst):
        raise PermissionError(13, "denied")

    monkeypatch.setattr(tx.os, "replace", locked)
    loop = tx.WatchLoop(tmp_path)
    loop.start(0.0)
    assert "请先关闭 Excel" in capsys.readouterr().err
    assert not xlsx_path(tmp_path).exists()

    # 仍被占用时继续重试，但不重复打印同一条提示。
    loop.tick(5.0)
    assert capsys.readouterr().err == ""

    monkeypatch.setattr(tx.os, "replace", real_replace)
    loop.tick(9.9)
    assert not xlsx_path(tmp_path).exists()
    loop.tick(10.0)
    assert xlsx_path(tmp_path).exists()
    assert tx.TO_XLSX in capsys.readouterr().out


def test_error_does_not_stop_later_events(tmp_path, capsys):
    calls = []

    def flaky(trip_id, trips_dir):
        calls.append(trip_id)
        if len(calls) == 1:
            raise RuntimeError("boom")
        return tx.SyncResult(trip_id, tx.NOOP)

    (tmp_path / "a.yaml").write_text("x", encoding="utf-8")
    loop = tx.WatchLoop(tmp_path, sync=flaky)
    loop.on_path(tmp_path / "a.yaml")
    loop.tick(0.0)
    loop.tick(2.0)
    assert "boom" in capsys.readouterr().err

    loop.on_path(tmp_path / "a.yaml")
    loop.tick(2.0)
    loop.tick(4.0)
    assert calls == ["a", "a"]


def test_deleting_both_files_is_silent(tmp_path, capsys):
    loop = tx.WatchLoop(tmp_path)
    loop.on_path(yaml_path(tmp_path))
    loop.tick(0.0)
    loop.tick(2.0)
    captured = capsys.readouterr()
    assert captured.out == "" and captured.err == ""


@pytest.mark.parametrize("extra", [["2026-a"], ["--check"], ["2026-a", "--prefer", "yaml"],
                                   ["--print-written"]])
def test_watch_rejects_other_options(extra):
    with pytest.raises(SystemExit):
        tx.parse_args(["sync", "--watch", *extra])
