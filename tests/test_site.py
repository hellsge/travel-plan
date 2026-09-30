import json
import re
from pathlib import Path

import pytest

import generate_site as gs

TRIP_A = """\
title: 甲旅行 <script>alert(1)</script>
start: 2026-08-01
end: 2026-08-02
travelers: 2
days:
  - date: 2026-08-01
    label: 第一天
    items:
      - title: 高铁 & 地铁
        type: 交通
        start: "07:28"
        end: "07:59"
        price: 50
        qty: 2
        status: 待购买
        place: 上海虹桥站
  - date: 2026-08-02
    items:
      - title: 午饭
        type: 餐饮
        cost: 100
"""

TRIP_B = """\
title: 乙旅行
start: 2025-10-01
end: 2025-10-01
days:
  - date: 2025-10-01
    items: []
"""


@pytest.fixture
def trips_dir(tmp_path: Path) -> Path:
    directory = tmp_path / "trips"
    directory.mkdir()
    (directory / "2026-a.yaml").write_text(TRIP_A, encoding="utf-8")
    (directory / "2025-b.yaml").write_text(TRIP_B, encoding="utf-8")
    return directory


@pytest.fixture
def site(trips_dir: Path, tmp_path: Path) -> Path:
    output = tmp_path / "output"
    assert gs.main(["--trips", str(trips_dir), "--output", str(output)]) == 0
    return output


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_home_lists_trips_newest_first(site):
    home = read(site / "index.html")
    links = re.findall(r'class="trip-card" href="([^"]+)"', home)
    assert links == ["2026-a/", "2025-b/"]


def test_trip_pages_exist_and_link_back(site):
    for trip_id in ("2026-a", "2025-b"):
        page = read(site / trip_id / "index.html")
        table = read(site / trip_id / "table.html")
        assert 'href="../"' in page and 'href="table.html"' in page
        assert 'href="../"' in table and 'href="./"' in table
        assert 'data-root="../"' in page


def test_html_escaped(site):
    page = read(site / "2026-a" / "index.html")
    assert "<script>alert(1)</script>" not in page
    assert "&lt;script&gt;" in page
    assert "高铁 &amp; 地铁" in page


def test_timeline_content(site):
    page = read(site / "2026-a" / "index.html")
    assert 'data-start="2026-08-01T07:28"' in page
    assert "¥50.00 × 2" in page
    assert 'data-location="上海虹桥站"' in page
    assert "8月2日 周日" in page
    assert "花销统计" in page


def test_table_content(site):
    table = read(site / "2026-a" / "table.html")
    assert 'class="type-cell type-交通"' in table
    assert 'class="status-cell status-待购买"' in table
    assert "¥200.00" in table


def test_service_worker_precaches_all_pages(site):
    worker = read(site / "service-worker.js")
    assets = json.loads(re.search(r"const ASSETS = (\[.*?\]);", worker, re.S).group(1))
    for path in ("./", "./2026-a/", "./2026-a/table.html", "./2025-b/", "./assets/app.css",
                 "./assets/common.js", "./assets/timeline.js", "./assets/home.js",
                 "./manifest.webmanifest", "./icons/icon-192.png"):
        assert path in assets
    manifest = json.loads(read(site / "manifest.webmanifest"))
    assert manifest["start_url"] == "./" and manifest["scope"] == "./"


def test_cache_version_changes_with_content(trips_dir, tmp_path):
    out1, out2 = tmp_path / "o1", tmp_path / "o2"
    gs.main(["--trips", str(trips_dir), "--output", str(out1)])
    (trips_dir / "2025-b.yaml").write_text(TRIP_B.replace("乙旅行", "乙旅行改"), encoding="utf-8")
    gs.main(["--trips", str(trips_dir), "--output", str(out2)])
    version = lambda p: re.search(r'CACHE_NAME = "([^"]+)"', read(p / "service-worker.js")).group(1)
    assert version(out1) != version(out2)


def test_check_prints_summary(trips_dir, capsys):
    assert gs.main(["--trips", str(trips_dir), "--check"]) == 0
    out = capsys.readouterr().out
    assert "2026-a" in out and "¥200.00" in out and "待办 1" in out
    assert "交通 ¥100.00" in out


def test_check_fails_on_invalid(trips_dir, capsys, tmp_path):
    (trips_dir / "2026-bad.yaml").write_text(TRIP_B.replace("items: []", "items: [{title: x, type: 飞行}]"), encoding="utf-8")
    output = tmp_path / "output"
    assert gs.main(["--trips", str(trips_dir), "--output", str(output)]) == 1
    assert "2026-bad.yaml" in capsys.readouterr().err
    assert not output.exists()


def test_empty_trips_dir_fails(tmp_path, capsys):
    empty = tmp_path / "empty"
    empty.mkdir()
    assert gs.main(["--trips", str(empty), "--check"]) == 1


def test_refuses_to_clear_foreign_directory(trips_dir, tmp_path, capsys):
    foreign = tmp_path / "foreign"
    foreign.mkdir()
    (foreign / "keep.txt").write_text("x", encoding="utf-8")
    assert gs.main(["--trips", str(trips_dir), "--output", str(foreign)]) == 1
    assert (foreign / "keep.txt").exists()


def test_regenerate_into_previous_output(site, trips_dir):
    assert gs.main(["--trips", str(trips_dir), "--output", str(site)]) == 0


def test_item_end_attributes(trips_dir, tmp_path):
    (trips_dir / "2025-b.yaml").write_text(TRIP_B.replace(
        "items: []", 'items: [{title: 夜车, type: 交通, start: "23:00", end: "01:00"}, {title: 游, type: 游览, start: "10:00"}]'
    ).replace("end: 2025-10-01", "end: 2025-10-02"), encoding="utf-8")
    output = tmp_path / "out"
    assert gs.main(["--trips", str(trips_dir), "--output", str(output)]) == 0
    page = read(output / "2025-b" / "index.html")
    assert 'data-start="2025-10-01T23:00" data-end="2025-10-02T01:00"' in page
    assert 'data-start="2025-10-01T10:00" data-end=""' in page
