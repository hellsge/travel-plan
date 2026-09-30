#!/usr/bin/env python3
"""由 trips/*.yaml 生成多旅行 PWA 站点；--check 只校验并输出统计摘要。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional, Union

from jinja2 import Environment, FileSystemLoader, StrictUndefined, select_autoescape

import site_assets
import trip_model as tm

ROOT = Path(__file__).resolve().parents[1]
TEMPLATES = ROOT / "templates"
WEEKDAYS = "一二三四五六日"
OUTPUT_MARKER = ".generated-site"


@dataclass
class TripEntry:
    id: str
    trip: tm.Trip
    stats: tm.TripStats

    @property
    def day_count(self) -> int:
        return (self.trip.end - self.trip.start).days + 1


def money(value: Optional[float]) -> str:
    return f"¥{value or 0:,.2f}"


def qty(value: float) -> str:
    return f"{value:g}"


def booking(value: Union[datetime, date]) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    return value.strftime("%Y-%m-%d")


def day_label(day: tm.Day) -> str:
    return day.label or f"{day.date.month}月{day.date.day}日 周{WEEKDAYS[day.date.weekday()]}"


def copy_text(item: tm.Item, day: tm.Day) -> str:
    when = item.start or "待定"
    if item.end:
        when += f"–{item.end}"
    parts = (f"{day.date.isoformat()} {when}", item.title, item.place, item.note)
    return "\n".join(part for part in parts if part)


def environment() -> Environment:
    env = Environment(
        loader=FileSystemLoader(TEMPLATES),
        autoescape=select_autoescape(["html"]),
        undefined=StrictUndefined,
        trim_blocks=True,
        lstrip_blocks=True,
    )
    env.filters.update(money=money, qty=qty, booking=booking, day_label=day_label,
                       copy_text=copy_text, end_attr=end_attr)
    return env


def build_entries(loaded: list[tm.LoadedTrip]) -> list[TripEntry]:
    entries = [TripEntry(t.id, t.trip, tm.compute_stats(t.trip)) for t in loaded]
    return sorted(entries, key=lambda e: (e.trip.start, e.id), reverse=True)


class OutputDirError(Exception):
    pass


def prepare_output(output: Path) -> None:
    """只清空空目录或带生成标记的目录，防止 --output 指错时删除其他文件。"""
    if output.exists() and any(output.iterdir()) and not (output / OUTPUT_MARKER).exists():
        raise OutputDirError(f"{output} 不是空目录，也不是之前生成的站点，拒绝清空")
    # 只清空内容、保留目录本身：Windows 上目录被本地预览服务器占用时无法删除。
    output.mkdir(parents=True, exist_ok=True)
    for child in output.iterdir():
        shutil.rmtree(child) if child.is_dir() else child.unlink()


def end_attr(item: tm.Item, day: tm.Day) -> str:
    """供前端判断当前事项；跨午夜的结束时间落在次日，无结束时间留空。"""
    if not item.end:
        return ""
    end_day = day.date + timedelta(days=1) if item.start and item.end < item.start else day.date
    return f"{end_day.isoformat()}T{item.end}"


def render_site(entries: list[TripEntry], output: Path) -> None:
    env = environment()
    pages: dict[str, str] = {"index.html": env.get_template("home.html").render(trips=entries, root="")}
    for entry in entries:
        context = dict(trip=entry.trip, trip_id=entry.id, stats=entry.stats, root="../")
        pages[f"{entry.id}/index.html"] = env.get_template("trip.html").render(**context)
        pages[f"{entry.id}/table.html"] = env.get_template("table.html").render(**context)

    prepare_output(output)
    digest = hashlib.sha256()
    for rel, content in sorted(pages.items()):
        path = output / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        digest.update(rel.encode() + content.encode("utf-8"))
    shutil.copytree(TEMPLATES / "assets", output / "assets")
    static = sorted(f"assets/{p.name}" for p in (output / "assets").iterdir())
    for rel in static:
        digest.update(rel.encode() + (output / rel).read_bytes())

    (output / OUTPUT_MARKER).write_text("由 generate_site.py 生成，整个目录会在下次生成时被清空。\n", encoding="utf-8")
    urls = [rel.removesuffix("index.html") for rel in sorted(pages)]
    assets = urls + static + [site_assets.write_manifest(output)] + site_assets.write_icons(output)
    site_assets.write_service_worker(output, assets, digest.hexdigest()[:12])


def summary(entry: TripEntry) -> str:
    stats = entry.stats
    types = "，".join(f"{kind} {money(amount)}" for kind, amount, _ in stats.by_type) or "无消费"
    return (f"{entry.id}: {entry.trip.start} 至 {entry.trip.end}，总额 {money(stats.total)}，"
            f"待办 {stats.todo_count}；{types}")


def parse_args(argv: Optional[list[str]]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="由 trips/*.yaml 生成旅行计划站点")
    parser.add_argument("--trips", default=str(ROOT / "trips"), help="旅行 YAML 目录")
    parser.add_argument("--output", default=str(ROOT / "output"), help="输出目录")
    parser.add_argument("--check", action="store_true", help="只校验并输出统计摘要，不生成")
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    loaded, errors = tm.load_trips(Path(args.trips))
    if errors:
        print("校验失败：\n" + "\n".join(f"  {e}" for e in errors), file=sys.stderr)
        return 1
    if not loaded:
        print(f"错误：{args.trips} 中没有旅行文件", file=sys.stderr)
        return 1
    entries = build_entries(loaded)
    if args.check:
        print("\n".join(summary(e) for e in entries))
        print(f"校验通过：{len(entries)} 次旅行")
        return 0
    try:
        render_site(entries, Path(args.output))
    except OutputDirError as error:
        print(f"错误：{error}", file=sys.stderr)
        return 1
    print(f"已生成站点：{args.output}（{len(entries)} 次旅行）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
