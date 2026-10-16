from datetime import date

import pytest

import trip_model as tm
import trip_xlsx as tx

TRIP_ID = "2026-sync"


def make_trip(**changes):
    items = [
        {"title": "高铁", "type": "交通", "price": 38, "qty": 2, "status": "已购买"},
        {"title": "酒店", "type": "住宿", "cost": 300, "status": "待购买"},
    ]
    data = {
        "title": "同步测试", "start": date(2026, 8, 1), "end": date(2026, 8, 2),
        "days": [{"date": date(2026, 8, 1), "items": items},
                 {"date": date(2026, 8, 2), "items": []}],
    }
    for key, value in changes.items():
        if key == "items":
            data["days"][0]["items"] = value
        else:
            data[key] = value
    return tm.Trip.model_validate(data)


def edited():
    return make_trip(items=[
        {"title": "高铁", "type": "交通", "price": 50, "qty": 2, "status": "已购买"},
        {"title": "晚饭", "type": "餐饮", "cost": 80, "status": "待预约"},
    ])


def yaml_path(d, trip_id=TRIP_ID):
    return d / f"{trip_id}.yaml"


def xlsx_path(d, trip_id=TRIP_ID):
    return d / f"{trip_id}.xlsx"


def put_yaml(d, trip, trip_id=TRIP_ID):
    yaml_path(d, trip_id).write_text(tm.dump_trip(trip), encoding="utf-8")


def put_xlsx(d, trip, meta="auto", trip_id=TRIP_ID):
    wb = tx.build_workbook(trip, tx.digest(trip) if meta == "auto" else meta or "")
    if meta is None:
        del wb[tx.META_SHEET]
    tx.write_workbook(wb, xlsx_path(d, trip_id))


def snapshot(d):
    return {p.name: p.read_bytes() for p in sorted(d.iterdir())}


def yaml_trip(d, trip_id=TRIP_ID):
    return tm.load_trip(yaml_path(d, trip_id)).trip


def xlsx_state(d, trip_id=TRIP_ID):
    trip, meta, _ = tx._load_xlsx(xlsx_path(d, trip_id))
    return tx.digest(trip), meta


def test_in_sync_writes_nothing(tmp_path):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, make_trip())
    before = snapshot(tmp_path)
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert (result.action, result.written, result.error) == (tx.NOOP, [], None)
    assert snapshot(tmp_path) == before


@pytest.mark.parametrize("meta", [None, "stale"])
def test_same_content_missing_meta_reexports(tmp_path, meta):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, make_trip(), meta=meta)
    yaml_before = yaml_path(tmp_path).read_bytes()
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.REEXPORT and result.written == [xlsx_path(tmp_path)]
    assert xlsx_state(tmp_path) == (tx.digest(make_trip()),) * 2
    assert yaml_path(tmp_path).read_bytes() == yaml_before


def test_only_xlsx_changed_writes_yaml(tmp_path):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, edited(), meta=tx.digest(make_trip()))
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.TO_YAML
    assert result.written == [yaml_path(tmp_path), xlsx_path(tmp_path)]
    assert tx.digest(yaml_trip(tmp_path)) == tx.digest(edited())
    assert xlsx_state(tmp_path) == (tx.digest(edited()),) * 2


def test_only_yaml_changed_exports_xlsx(tmp_path):
    put_yaml(tmp_path, edited())
    put_xlsx(tmp_path, make_trip())
    yaml_before = yaml_path(tmp_path).read_bytes()
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.TO_XLSX and result.written == [xlsx_path(tmp_path)]
    assert xlsx_state(tmp_path) == (tx.digest(edited()),) * 2
    assert yaml_path(tmp_path).read_bytes() == yaml_before


@pytest.mark.parametrize("meta", [None, "unrelated"])
def test_conflict_writes_nothing(tmp_path, meta):
    put_yaml(tmp_path, edited())
    put_xlsx(tmp_path, make_trip(title="另一边"), meta=meta)
    before = snapshot(tmp_path)
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.CONFLICT and result.written == []
    assert "--prefer" in result.error
    assert result.changes
    assert snapshot(tmp_path) == before


def test_prefer_yaml(tmp_path):
    put_yaml(tmp_path, edited())
    put_xlsx(tmp_path, make_trip(title="另一边"), meta=None)
    result = tx.sync_trip(TRIP_ID, tmp_path, prefer="yaml")
    assert result.action == tx.TO_XLSX
    assert xlsx_state(tmp_path) == (tx.digest(edited()),) * 2


def test_prefer_xlsx(tmp_path):
    put_yaml(tmp_path, edited())
    other = make_trip(title="另一边")
    put_xlsx(tmp_path, other, meta=None)
    result = tx.sync_trip(TRIP_ID, tmp_path, prefer="xlsx")
    assert result.action == tx.TO_YAML
    assert tx.digest(yaml_trip(tmp_path)) == tx.digest(other)
    assert xlsx_state(tmp_path) == (tx.digest(other),) * 2


def test_prefer_overrides_meta(tmp_path):
    # 只改了 yaml 的情况下 --prefer xlsx 仍以 xlsx 为准，用于放弃 yaml 中的修改。
    put_yaml(tmp_path, edited())
    put_xlsx(tmp_path, make_trip())
    assert tx.sync_trip(TRIP_ID, tmp_path, prefer="xlsx").action == tx.TO_YAML
    assert tx.digest(yaml_trip(tmp_path)) == tx.digest(make_trip())


def test_only_xlsx_creates_yaml(tmp_path):
    put_xlsx(tmp_path, edited(), meta="copied-from-another-trip")
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.TO_YAML
    assert result.changes.startswith("新建 2 项事项")
    assert tx.digest(yaml_trip(tmp_path)) == tx.digest(edited())
    assert xlsx_state(tmp_path) == (tx.digest(edited()),) * 2


def test_only_yaml_creates_xlsx(tmp_path):
    put_yaml(tmp_path, make_trip())
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.TO_XLSX and result.written == [xlsx_path(tmp_path)]
    assert xlsx_state(tmp_path) == (tx.digest(make_trip()),) * 2


def test_change_summary(tmp_path):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, edited(), meta=tx.digest(make_trip()))
    changes = tx.sync_trip(TRIP_ID, tmp_path).changes
    assert "新增 1 项、删除 1 项、修改 1 项" in changes
    assert "总额 ¥376.00 → ¥180.00" in changes
    assert "待办 1 → 1" in changes


def test_change_summary_moved_and_duplicate_titles():
    a = {"title": "散步", "type": "游览"}
    old = make_trip(items=[a, a])
    new = tm.Trip.model_validate({**old.model_dump(), "days": [
        {"date": date(2026, 8, 1), "items": [{**a, "note": "改了"}, {"title": "新的", "type": "其他"}]},
        {"date": date(2026, 8, 2), "items": [a]},
    ]})
    # 挪到另一天也算修改。
    assert tx._item_changes(old, new) == (1, 0, 2)
    assert tx._item_changes(old, make_trip(items=[a])) == (0, 1, 0)
    assert tx._change_summary(old, old).startswith("事项未变")


def test_yaml_line_endings_preserved(tmp_path):
    yaml_path(tmp_path).write_bytes(tm.dump_trip(make_trip()).replace("\n", "\r\n").encode())
    put_xlsx(tmp_path, edited(), meta=tx.digest(make_trip()))
    tx.sync_trip(TRIP_ID, tmp_path)
    data = yaml_path(tmp_path).read_bytes()
    assert data.count(b"\r\n") == data.count(b"\n") > 0

    yaml_path(tmp_path).write_bytes(tm.dump_trip(make_trip()).encode())
    put_xlsx(tmp_path, edited(), meta=tx.digest(make_trip()))
    tx.sync_trip(TRIP_ID, tmp_path)
    assert b"\r\n" not in yaml_path(tmp_path).read_bytes()


def test_check_mode_reports_without_writing(tmp_path):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, edited(), meta=tx.digest(make_trip()))
    before = snapshot(tmp_path)
    result = tx.sync_trip(TRIP_ID, tmp_path, check=True)
    assert result.action == tx.TO_YAML and result.written == []
    assert snapshot(tmp_path) == before


@pytest.mark.parametrize("setup, code", [
    (lambda d: (put_yaml(d, make_trip()), put_xlsx(d, make_trip())), 0),
    (lambda d: (put_yaml(d, make_trip()), put_xlsx(d, make_trip(), meta=None)), 0),
    (lambda d: (put_yaml(d, edited()), put_xlsx(d, make_trip())), 1),
    (lambda d: put_yaml(d, make_trip()), 1),
    (lambda d: put_xlsx(d, make_trip()), 1),
], ids=["in-sync", "meta-only", "differs", "no-xlsx", "no-yaml"])
def test_main_check_exit_code(tmp_path, setup, code):
    setup(tmp_path)
    before = snapshot(tmp_path)
    assert tx.main(["sync", "--check", "--trips", str(tmp_path)]) == code
    assert snapshot(tmp_path) == before


def test_locked_xlsx_reports_hint(tmp_path, monkeypatch, capsys):
    put_yaml(tmp_path, make_trip())

    def locked(src, dst):
        raise PermissionError(13, "denied")

    monkeypatch.setattr(tx.os, "replace", locked)
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.error == f"请先关闭 Excel：{xlsx_path(tmp_path)}"
    assert result.written == []
    assert tx.main(["sync", "--trips", str(tmp_path)]) == 1
    assert "请先关闭 Excel" in capsys.readouterr().err
    assert sorted(p.name for p in tmp_path.iterdir()) == [f"{TRIP_ID}.yaml"]


def test_yaml_written_even_if_xlsx_refresh_locked(tmp_path, monkeypatch):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, edited(), meta=tx.digest(make_trip()))
    real_replace = tx.os.replace

    def lock_xlsx(src, dst):
        if str(dst).endswith(".xlsx"):
            raise PermissionError(13, "denied")
        real_replace(src, dst)

    monkeypatch.setattr(tx.os, "replace", lock_xlsx)
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.TO_YAML and "请先关闭 Excel" in result.error
    assert result.written == [yaml_path(tmp_path)]
    assert tx.digest(yaml_trip(tmp_path)) == tx.digest(edited())


@pytest.mark.parametrize("name", ["Bad Name.xlsx", "assets.xlsx"])
def test_invalid_file_name(tmp_path, name):
    put_xlsx(tmp_path, make_trip(), trip_id=name.removesuffix(".xlsx"))
    [result] = tx.sync_all(tmp_path)
    assert result.action == tx.ERROR and name in result.error
    assert sorted(p.name for p in tmp_path.iterdir()) == [name]


def test_path_traversal_id_rejected(tmp_path):
    inner = tmp_path / "trips"
    inner.mkdir()
    put_yaml(tmp_path, make_trip(), trip_id="outside")
    result = tx.sync_trip("../outside", inner)
    assert result.action == tx.ERROR and result.written == []
    assert not (tmp_path / "outside.xlsx").exists()


def test_missing_trip(tmp_path):
    result = tx.sync_trip("nothing", tmp_path)
    assert result.action == tx.ERROR and "找不到" in result.error


def test_invalid_yaml_leaves_xlsx(tmp_path):
    put_xlsx(tmp_path, make_trip())
    yaml_path(tmp_path).write_text("title: x\n", encoding="utf-8")
    before = snapshot(tmp_path)
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.ERROR and f"{TRIP_ID}.yaml" in result.error
    assert snapshot(tmp_path) == before


def test_invalid_xlsx_reports_file_and_row(tmp_path):
    put_yaml(tmp_path, make_trip())
    put_xlsx(tmp_path, make_trip())
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path(tmp_path))
    wb.worksheets[0][f"D{tx.FIRST_ROW + 1}"] = "吃饭"
    wb.save(xlsx_path(tmp_path))
    result = tx.sync_trip(TRIP_ID, tmp_path)
    assert result.action == tx.ERROR
    assert f"{TRIP_ID}.xlsx: 第 {tx.FIRST_ROW + 1} 行 D 列" in result.error


def test_prefer_recovers_from_broken_other_side(tmp_path):
    put_yaml(tmp_path, make_trip())
    xlsx_path(tmp_path).write_text("not a zip")
    assert tx.sync_trip(TRIP_ID, tmp_path).action == tx.ERROR
    result = tx.sync_trip(TRIP_ID, tmp_path, prefer="yaml")
    assert result.action == tx.TO_XLSX and result.error is None
    assert xlsx_state(tmp_path) == (tx.digest(make_trip()),) * 2

    yaml_path(tmp_path).write_text("title: x\n", encoding="utf-8")
    assert tx.sync_trip(TRIP_ID, tmp_path, prefer="xlsx").action == tx.TO_YAML
    assert tx.digest(yaml_trip(tmp_path)) == tx.digest(make_trip())


def test_sync_all_partial_failure(tmp_path, capsys):
    put_yaml(tmp_path, make_trip(), trip_id="good")
    yaml_path(tmp_path, "broken").write_text("title: x\n", encoding="utf-8")
    (tmp_path / "~$good.xlsx").write_bytes(b"lock")
    assert tx.main(["sync", "--trips", str(tmp_path)]) == 1
    assert xlsx_path(tmp_path, "good").exists()
    out = capsys.readouterr()
    assert "good: yaml → xlsx" in out.out
    assert "broken" in out.err and "~$" not in out.out + out.err


def test_main_single_id_and_prefer(tmp_path):
    put_yaml(tmp_path, edited())
    put_xlsx(tmp_path, make_trip(title="另一边"), meta=None)
    assert tx.main(["sync", TRIP_ID, "--trips", str(tmp_path)]) == 1
    assert tx.main(["sync", TRIP_ID, "--prefer", "yaml", "--trips", str(tmp_path)]) == 0
    assert tx.main(["sync", TRIP_ID, "--trips", str(tmp_path)]) == 0


def test_main_prefer_requires_id(tmp_path):
    with pytest.raises(SystemExit):
        tx.main(["sync", "--prefer", "yaml", "--trips", str(tmp_path)])


def test_main_empty_dir(tmp_path):
    assert tx.main(["sync", "--trips", str(tmp_path)]) == 1


def test_main_print_written_lists_only_paths(tmp_path, monkeypatch, capsys):
    trips = tmp_path / "trips"
    trips.mkdir()
    monkeypatch.chdir(tmp_path)
    put_yaml(trips, make_trip())
    args = ["sync", "--print-written", "--trips", str(trips)]
    assert tx.main(args) == 0
    out = capsys.readouterr()
    assert out.out.splitlines() == [f"trips/{TRIP_ID}.xlsx"]
    assert tx.TO_XLSX in out.err and "同步完成" in out.err

    put_xlsx(trips, edited(), meta=tx.digest(make_trip()))
    assert tx.main(args) == 0
    assert capsys.readouterr().out.splitlines() == [f"trips/{TRIP_ID}.yaml",
                                                    f"trips/{TRIP_ID}.xlsx"]
    assert tx.main(args) == 0
    assert capsys.readouterr().out == ""


def test_main_print_written_on_partial_failure(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    put_yaml(tmp_path, make_trip(), trip_id="good")
    yaml_path(tmp_path, "broken").write_text("title: x\n", encoding="utf-8")
    assert tx.main(["sync", "--print-written", "--trips", str(tmp_path)]) == 1
    out = capsys.readouterr()
    assert out.out.splitlines() == ["good.xlsx"]
    assert "broken" in out.err
