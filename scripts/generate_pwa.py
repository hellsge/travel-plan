#!/usr/bin/env python3
"""从 travel.xlsx 的新版旅行 Sheet 生成手机友好的离线 HTML。"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import struct
import sys
import zlib
from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
from typing import Any
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


DEFAULT_WORKBOOK = "travel.xlsx"
DEFAULT_SHEET = "26暑期云南"
HEADERS = [
    "日期", "开始", "结束", "类型", "项目", "路线 / 备注",
    "单价", "数量", "消费", "开售 / 预约时间", "状态",
]
TODO_STATUSES = {"待购买", "待预约"}
TYPE_NAMES = {"交通", "住宿", "游览", "餐饮", "日常", "购物", "其他"}


@dataclass
class TripItem:
    row: int
    day: date
    start: time | None
    end: time | None
    kind: str
    title: str
    note: str
    unit_price: float | None
    quantity: float | None
    cost: float | None
    booking_at: datetime | date | None
    status: str


@dataclass
class TripDay:
    day: date
    label: str
    items: list[TripItem]


WEEKDAYS = "一二三四五六日"


def as_date(value: Any, epoch: datetime) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        converted = from_excel(value, epoch)
        return converted.date() if isinstance(converted, datetime) else converted
    return None


def as_time(value: Any, epoch: datetime) -> time | None:
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, time):
        return value.replace(microsecond=0)
    if isinstance(value, (int, float)):
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime):
            return converted.time().replace(microsecond=0)
        if isinstance(converted, time):
            return converted.replace(microsecond=0)
    return None


def as_booking_time(value: Any, epoch: datetime) -> datetime | date | None:
    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        converted = from_excel(value, epoch)
        if isinstance(converted, (datetime, date)):
            return converted
    return None


def as_number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def calculate_cost(formula_value: Any, cached_value: Any, unit: Any, quantity: Any) -> float | None:
    cached = as_number(cached_value)
    if cached is not None:
        return cached
    direct = as_number(formula_value)
    if direct is not None:
        return direct
    unit_number = as_number(unit)
    quantity_number = as_number(quantity)
    if unit_number is not None and quantity_number is not None:
        return unit_number * quantity_number
    return None


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def default_day_label(day: date) -> str:
    return f"{day.month}月{day.day}日 周{WEEKDAYS[day.weekday()]}"


def validate_sheet(sheet: Any) -> None:
    actual = [clean_text(sheet.cell(4, column).value) for column in range(1, 12)]
    if actual != HEADERS:
        raise ValueError(
            f"Sheet“{sheet.title}”不是新版模板：第 4 行应为 {' / '.join(HEADERS)}"
        )


def read_trip(workbook_path: Path, sheet_name: str) -> tuple[str, date | None, date | None, list[TripDay]]:
    formulas = load_workbook(workbook_path, data_only=False)
    cached = load_workbook(workbook_path, data_only=True)
    try:
        if sheet_name not in formulas.sheetnames:
            raise ValueError(f"找不到 Sheet“{sheet_name}”")
        sheet = formulas[sheet_name]
        cached_sheet = cached[sheet_name]
        validate_sheet(sheet)

        title = clean_text(sheet["A1"].value) or sheet_name
        start_day = as_date(sheet["B2"].value, formulas.epoch)
        end_day = as_date(sheet["C2"].value, formulas.epoch)
        separators: dict[date, str] = {}
        items_by_day: dict[date, list[TripItem]] = {}

        merged_rows = {
            cell_range.min_row
            for cell_range in sheet.merged_cells.ranges
            if cell_range.min_col == 1 and cell_range.max_col >= 11
        }
        for row in range(5, sheet.max_row + 1):
            if row in merged_rows:
                text = clean_text(sheet.cell(row, 1).value)
                match = re.search(r"(\d{1,2})月(\d{1,2})日", text)
                if match:
                    year = start_day.year if start_day else date.today().year
                    try:
                        separators[date(year, int(match.group(1)), int(match.group(2)))] = text
                    except ValueError:
                        pass
                continue

            day = as_date(sheet.cell(row, 1).value, formulas.epoch)
            values = [sheet.cell(row, column).value for column in range(1, 12)]
            if day is None or not any(value not in (None, "") for value in values[1:]):
                continue

            unit = sheet.cell(row, 7).value
            quantity = sheet.cell(row, 8).value
            item = TripItem(
                row=row,
                day=day,
                start=as_time(sheet.cell(row, 2).value, formulas.epoch),
                end=as_time(sheet.cell(row, 3).value, formulas.epoch),
                kind=clean_text(sheet.cell(row, 4).value) or "其他",
                title=clean_text(sheet.cell(row, 5).value) or "未命名事项",
                note=clean_text(sheet.cell(row, 6).value),
                unit_price=as_number(unit),
                quantity=as_number(quantity),
                cost=calculate_cost(
                    sheet.cell(row, 9).value,
                    cached_sheet.cell(row, 9).value,
                    unit,
                    quantity,
                ),
                booking_at=as_booking_time(sheet.cell(row, 10).value, formulas.epoch),
                status=clean_text(sheet.cell(row, 11).value) or "无需预订",
            )
            items_by_day.setdefault(day, []).append(item)

        days = [
            TripDay(day=day, label=separators.get(day, default_day_label(day)), items=items)
            for day, items in sorted(items_by_day.items())
        ]
        if not days:
            raise ValueError(f"Sheet“{sheet_name}”中没有可生成的行程事项")
        return title, start_day, end_day, days
    finally:
        formulas.close()
        cached.close()


def esc(value: Any) -> str:
    return html.escape(str(value), quote=True)


def money(value: float) -> str:
    return f"¥{value:,.2f}"


def time_range(item: TripItem) -> str:
    start = item.start.strftime("%H:%M") if item.start else "待定"
    if not item.end:
        return start
    return f"{start}–{item.end.strftime('%H:%M')}"


def booking_text(value: datetime | date) -> str:
    if isinstance(value, datetime) and value.time() != time(0, 0):
        return value.strftime("%Y-%m-%d %H:%M")
    return value.strftime("%Y-%m-%d")


def slug_class(value: str, allowed: set[str], fallback: str) -> str:
    return value if value in allowed else fallback


def render_item(item: TripItem) -> str:
    kind = slug_class(item.kind, TYPE_NAMES, "其他")
    todo = item.status in TODO_STATUSES
    details: list[str] = []
    if item.note:
        details.append(f'<div class="note">{esc(item.note)}</div>')
    if item.booking_at:
        details.append(
            f'<div class="booking">⏰ 开售 / 预约：{esc(booking_text(item.booking_at))}</div>'
        )
    if item.cost is not None:
        price_detail = ""
        if item.unit_price is not None and item.quantity is not None:
            price_detail = f'<span class="price-detail">{money(item.unit_price)} × {item.quantity:g}</span>'
        details.append(f'<div class="cost">{price_detail}<strong>{money(item.cost)}</strong></div>')

    action_text = item.note or item.title
    map_url = "https://uri.amap.com/search?keyword=" + quote(action_text)
    actions = (
        f'<div class="actions"><a href="{esc(map_url)}" target="_blank" rel="noopener">地图搜索</a>'
        f'<button type="button" class="copy" data-copy="{esc(action_text)}">复制地点</button></div>'
    )
    return f'''<article class="item type-{esc(kind)}{' todo' if todo else ''}" data-todo="{'1' if todo else '0'}">
      <div class="time">{esc(time_range(item))}</div>
      <div class="item-body">
        <div class="item-heading"><span class="type">{esc(item.kind)}</span><h3>{esc(item.title)}</h3></div>
        {''.join(details)}
        <div class="item-footer"><span class="status status-{esc(item.status)}">{esc(item.status)}</span>{actions}</div>
      </div>
    </article>'''


def render_html(title: str, sheet_name: str, start_day: date | None, end_day: date | None, days: list[TripDay]) -> str:
    all_items = [item for day in days for item in day.items]
    total_cost = sum(item.cost or 0 for item in all_items)
    todo_count = sum(item.status in TODO_STATUSES for item in all_items)
    date_text = ""
    if start_day and end_day:
        date_text = f"{start_day:%Y-%m-%d} 至 {end_day:%Y-%m-%d}"

    nav = "".join(
        f'<a href="#day-{day.day.isoformat()}" data-day-link="{day.day.isoformat()}"><b>{day.day.day}</b><span>{day.day.month}月</span></a>'
        for day in days
    )
    sections: list[str] = []
    for trip_day in days:
        day_cost = sum(item.cost or 0 for item in trip_day.items)
        day_todos = sum(item.status in TODO_STATUSES for item in trip_day.items)
        sections.append(f'''<section class="day" id="day-{trip_day.day.isoformat()}" data-date="{trip_day.day.isoformat()}">
  <header class="day-header">
    <div><span class="day-date">{trip_day.day:%m月%d日}</span><h2>{esc(trip_day.label)}</h2></div>
    <div class="day-summary"><strong>{money(day_cost)}</strong><span>{len(trip_day.items)} 项{f' · {day_todos} 待办' if day_todos else ''}</span></div>
  </header>
  <div class="timeline">{''.join(render_item(item) for item in trip_day.items)}</div>
</section>''')

    return f'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="color-scheme" content="light">
<meta name="theme-color" content="#2563eb">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="旅行计划">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icons/icon-192.png">
<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">
<title>{esc(title)}</title>
<style>
:root{{--bg:#f4f7fb;--card:#fff;--ink:#172033;--muted:#667085;--line:#dbe4ef;--primary:#2563eb;--shadow:0 8px 24px rgba(30,55,90,.08)}}
*{{box-sizing:border-box}}html{{scroll-behavior:smooth;scroll-padding-top:calc(82px + env(safe-area-inset-top))}}body{{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;-webkit-text-size-adjust:100%;text-rendering:optimizeLegibility}}button,a{{font:inherit;-webkit-tap-highlight-color:transparent}}.hero{{padding:calc(28px + env(safe-area-inset-top)) max(18px,env(safe-area-inset-right)) 20px max(18px,env(safe-area-inset-left));background:linear-gradient(135deg,#173b73,#2563eb);color:#fff}}.hero-inner,main{{max-width:760px;margin:auto}}.eyebrow{{margin:0 0 7px;font-size:12px;opacity:.76}}h1{{font-size:25px;line-height:1.25;margin:0}}.date-range{{margin:8px 0 0;opacity:.86;font-size:14px}}.stats{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:20px}}.stat{{padding:13px 14px;border:1px solid rgba(255,255,255,.18);border-radius:14px;background:rgba(255,255,255,.12);backdrop-filter:blur(6px)}}.stat span{{display:block;font-size:12px;opacity:.8}}.stat strong{{display:block;margin-top:3px;font-size:20px}}.toolbar{{position:sticky;z-index:20;top:0;background:rgba(244,247,251,.94);border-bottom:1px solid var(--line);-webkit-backdrop-filter:blur(12px);backdrop-filter:blur(12px)}}.toolbar-inner{{display:flex;align-items:center;gap:10px;max-width:760px;margin:auto;padding:9px max(14px,env(safe-area-inset-right)) 9px max(14px,env(safe-area-inset-left))}}.day-nav{{display:flex;gap:7px;overflow-x:auto;scrollbar-width:none;flex:1}}.day-nav::-webkit-scrollbar{{display:none}}.day-nav a{{flex:0 0 46px;min-height:44px;text-align:center;text-decoration:none;color:#344054;background:#fff;border:1px solid var(--line);border-radius:11px;padding:5px 4px}}.day-nav b,.day-nav span{{display:block}}.day-nav b{{font-size:16px}}.day-nav span{{font-size:10px;color:var(--muted)}}.day-nav a.active{{color:#fff;background:var(--primary);border-color:var(--primary)}}.day-nav a.active span{{color:#dbeafe}}.filter{{flex:0 0 auto;min-height:44px;border:1px solid #f2c94c;background:#fff8db;color:#704d00;border-radius:11px;padding:10px 11px;cursor:pointer}}.filter.active{{background:#facc15}}main{{padding:16px max(14px,env(safe-area-inset-right)) calc(50px + env(safe-area-inset-bottom)) max(14px,env(safe-area-inset-left))}}.day{{margin:0 0 18px;scroll-margin-top:74px}}.day-header{{display:flex;align-items:end;justify-content:space-between;gap:12px;margin-bottom:9px;padding:0 2px}}.day-header h2{{font-size:17px;margin:2px 0 0}}.day-date{{color:var(--primary);font-weight:800;font-size:12px}}.day-summary{{text-align:right;white-space:nowrap}}.day-summary strong,.day-summary span{{display:block}}.day-summary strong{{font-size:15px}}.day-summary span{{font-size:11px;color:var(--muted);margin-top:2px}}.day.today .day-header{{border-left:4px solid var(--primary);padding-left:9px}}.timeline{{position:relative}}.timeline:before{{content:"";position:absolute;left:62px;top:18px;bottom:18px;width:2px;background:var(--line)}}.item{{--type:#64748b;position:relative;display:grid;grid-template-columns:54px 1fr;gap:18px;margin-bottom:10px}}.item:before{{content:"";position:absolute;left:58px;top:20px;width:10px;height:10px;border:3px solid var(--bg);border-radius:50%;background:var(--type);z-index:1}}.time{{padding-top:14px;font-size:11px;font-weight:700;color:#475467;text-align:right}}.item-body{{min-width:0;padding:13px 13px 11px;border-radius:15px;background:var(--card);border:1px solid #e8edf4;box-shadow:var(--shadow)}}.item.todo .item-body{{border-color:#f4cf5d}}.item-heading{{display:flex;align-items:center;gap:8px}}.item-heading h3{{font-size:16px;margin:0;min-width:0}}.type{{flex:0 0 auto;padding:3px 7px;border-radius:7px;color:var(--type);background:color-mix(in srgb,var(--type) 13%,white);font-size:11px;font-weight:800}}.note{{margin-top:8px;color:#475467;line-height:1.55;font-size:13px;white-space:pre-wrap;overflow-wrap:anywhere}}.booking{{margin-top:8px;padding:7px 9px;border-radius:8px;background:#fff8db;color:#765600;font-size:12px}}.cost{{display:flex;align-items:center;justify-content:flex-end;gap:7px;margin-top:8px}}.cost strong{{font-size:14px}}.price-detail{{font-size:11px;color:var(--muted)}}.item-footer{{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:10px}}.status{{padding:4px 8px;border-radius:99px;background:#e7edf4;color:#344054;font-size:11px;font-weight:800}}.status-待购买,.status-待预约{{background:#fef0b8;color:#704d00}}.status-已购买,.status-已预约,.status-已完成{{background:#d7f5df;color:#176b32}}.status-已取消{{background:#e5e7eb;color:#6b7280}}.actions{{display:flex;gap:5px}}.actions a,.actions button{{display:inline-flex;align-items:center;justify-content:center;min-height:36px;border:0;background:#eef4ff;color:#2457a7;border-radius:8px;padding:7px 9px;text-decoration:none;cursor:pointer;font-size:11px}}.type-交通{{--type:#2673d9}}.type-住宿{{--type:#7c4dcc}}.type-游览{{--type:#219653}}.type-餐饮{{--type:#dc7318}}.type-日常{{--type:#667085}}.type-购物{{--type:#d94f8a}}.type-其他{{--type:#80685b}}.empty{{display:none;text-align:center;color:var(--muted);padding:40px 10px}}footer{{text-align:center;color:#98a2b3;font-size:11px;padding:8px 16px 28px}}@media(max-width:420px){{h1{{font-size:23px}}.hero{{padding-bottom:18px}}.stats{{margin-top:16px}}.stat{{padding:11px 12px}}.stat strong{{font-size:18px}}.item{{grid-template-columns:50px 1fr;gap:15px}}.timeline:before{{left:58px}}.item:before{{left:54px}}.item-body{{padding:12px 12px 10px}}.actions{{gap:4px}}}}@media(max-width:370px){{.actions a{{display:none}}}}@media(display-mode:standalone){{.toolbar{{top:env(safe-area-inset-top)}}}}@media(prefers-reduced-motion:reduce){{html{{scroll-behavior:auto}}}}
</style>
</head>
<body>
<header class="hero"><div class="hero-inner"><p class="eyebrow">由 Excel 自动生成 · {esc(sheet_name)}</p><h1>{esc(title)}</h1><p class="date-range">{esc(date_text)}</p><div class="stats"><div class="stat"><span>预计总消费</span><strong>{money(total_cost)}</strong></div><div class="stat"><span>待处理事项</span><strong>{todo_count}</strong></div></div></div></header>
<nav class="toolbar"><div class="toolbar-inner"><div class="day-nav">{nav}</div><button id="todoFilter" class="filter" type="button">只看待办（{todo_count}）</button></div></nav>
<main>{''.join(sections)}<div id="empty" class="empty">没有待处理事项</div></main>
<footer>内容来源：{esc(DEFAULT_WORKBOOK)} / {esc(sheet_name)}；请修改 Excel 后重新生成。</footer>
<script>
(()=>{{
  const filter=document.getElementById('todoFilter'),empty=document.getElementById('empty');
  filter.addEventListener('click',()=>{{
    const active=filter.classList.toggle('active');
    document.querySelectorAll('.item').forEach(el=>el.hidden=active&&el.dataset.todo!=='1');
    document.querySelectorAll('.day').forEach(day=>day.hidden=active&&!day.querySelector('.item[data-todo="1"]'));
    empty.style.display=active&&!document.querySelector('.item[data-todo="1"]')?'block':'none';
    filter.setAttribute('aria-pressed',String(active));
  }});
  document.addEventListener('click',async event=>{{
    const button=event.target.closest('.copy');if(!button)return;
    try{{await navigator.clipboard.writeText(button.dataset.copy)}}catch(_){{
      const area=document.createElement('textarea');area.value=button.dataset.copy;document.body.append(area);area.select();document.execCommand('copy');area.remove();
    }}
    const old=button.textContent;button.textContent='已复制';setTimeout(()=>button.textContent=old,1200);
  }});
  const days=[...document.querySelectorAll('.day')],today=new Date(),local=`${{today.getFullYear()}}-${{String(today.getMonth()+1).padStart(2,'0')}}-${{String(today.getDate()).padStart(2,'0')}}`;
  let current=days.find(day=>day.dataset.date===local)||days.find(day=>day.dataset.date>local)||days[days.length-1];
  if(current){{current.classList.add('today');document.querySelector(`[data-day-link="${{current.dataset.date}}"]`)?.classList.add('active')}}
  const observer=new IntersectionObserver(entries=>{{
    const visible=entries.filter(entry=>entry.isIntersecting).sort((a,b)=>b.intersectionRatio-a.intersectionRatio)[0];if(!visible)return;
    document.querySelectorAll('[data-day-link]').forEach(link=>link.classList.toggle('active',link.dataset.dayLink===visible.target.dataset.date));
  }},{{rootMargin:'-75px 0px -55% 0px',threshold:[0,.2,.5]}});days.forEach(day=>observer.observe(day));
  if('serviceWorker' in navigator && location.protocol.startsWith('http')){{
    window.addEventListener('load',()=>navigator.serviceWorker.register('./service-worker.js').catch(error=>console.warn('离线缓存注册失败',error)));
  }}
}})();
</script>
</body>
</html>
'''


def safe_filename(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return cleaned or "travel"


def compatible_sheets(workbook_path: Path) -> list[str]:
    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    try:
        result = []
        for name in workbook.sheetnames:
            sheet = workbook[name]
            actual = [clean_text(sheet.cell(4, column).value) for column in range(1, 12)]
            if actual == HEADERS:
                result.append(name)
        return result
    finally:
        workbook.close()


def png_chunk(chunk_type: bytes, data: bytes) -> bytes:
    return (
        struct.pack(">I", len(data))
        + chunk_type
        + data
        + struct.pack(">I", zlib.crc32(chunk_type + data) & 0xFFFFFFFF)
    )


def write_icon(path: Path, size: int) -> None:
    """生成不依赖第三方图像库的蓝色旅行图标。"""
    rows = bytearray()
    center = (size - 1) / 2
    radius = size * 0.39
    for y in range(size):
        rows.append(0)
        for x in range(size):
            distance = ((x - center) ** 2 + (y - center) ** 2) ** 0.5
            if distance <= radius:
                color = (255, 255, 255, 255)
            else:
                color = (37, 99, 235, 255)
            rows.extend(color)
    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    content = (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", header)
        + png_chunk(b"IDAT", zlib.compress(bytes(rows), 9))
        + png_chunk(b"IEND", b"")
    )
    path.write_bytes(content)


def write_pwa_files(output_dir: Path, title: str, html_filename: str, content_version: str) -> None:
    icons_dir = output_dir / "icons"
    icons_dir.mkdir(parents=True, exist_ok=True)
    write_icon(icons_dir / "icon-192.png", 192)
    write_icon(icons_dir / "icon-512.png", 512)
    write_icon(icons_dir / "apple-touch-icon.png", 180)

    manifest = {
        "name": title,
        "short_name": "旅行计划",
        "description": "从 Excel 自动生成的离线旅行计划",
        "lang": "zh-CN",
        "start_url": f"./{html_filename}",
        "scope": "./",
        "display": "standalone",
        "background_color": "#f4f7fb",
        "theme_color": "#2563eb",
        "icons": [
            {"src": "icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
        ],
    }
    (output_dir / "manifest.webmanifest").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    cache_name = f"travel-{content_version}"
    assets = [
        f"./{html_filename}",
        "./manifest.webmanifest",
        "./icons/icon-192.png",
        "./icons/icon-512.png",
        "./icons/apple-touch-icon.png",
    ]
    service_worker = f'''const CACHE_NAME = {json.dumps(cache_name)};
const ASSETS = {json.dumps(assets, ensure_ascii=False, indent=2)};

self.addEventListener('install', event => {{
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting()));
}});

self.addEventListener('activate', event => {{
  event.waitUntil(caches.keys().then(names => Promise.all(
    names.filter(name => name.startsWith('travel-') && name !== CACHE_NAME).map(name => caches.delete(name))
  )).then(() => self.clients.claim()));
}});

self.addEventListener('fetch', event => {{
  if (event.request.method !== 'GET') return;
  event.respondWith(fetch(event.request).then(response => {{
    if (response.ok && new URL(event.request.url).origin === self.location.origin) {{
      const copy = response.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
    }}
    return response;
  }}).catch(() => caches.match(event.request).then(cached => cached || caches.match('./{html_filename}'))));
}});
'''
    (output_dir / "service-worker.js").write_text(service_worker, encoding="utf-8")


def generate(workbook_path: Path, sheet_name: str, output_path: Path) -> None:
    title, start_day, end_day, days = read_trip(workbook_path, sheet_name)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = render_html(title, sheet_name, start_day, end_day, days)
    output_path.write_text(document, encoding="utf-8")
    content_version = hashlib.sha256(document.encode("utf-8")).hexdigest()[:12]
    write_pwa_files(output_path.parent, title, output_path.name, content_version)
    item_count = sum(len(day.items) for day in days)
    print(f"已生成 PWA：{output_path}（{len(days)} 天，{item_count} 项）")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从旅行 Excel 新版模板生成手机端 HTML")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK, help="Excel 文件路径")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="要生成的 Sheet 名称")
    parser.add_argument("--output", help="输出 HTML 路径；默认写入 output/index.html，并生成同目录 PWA 文件")
    parser.add_argument("--all", action="store_true", help="生成工作簿内所有新版模板 Sheet")
    parser.add_argument("--list-sheets", action="store_true", help="列出可生成的新版模板 Sheet")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workbook_path = Path(args.workbook).resolve()
    if not workbook_path.is_file():
        print(f"错误：找不到 Excel 文件：{workbook_path}", file=sys.stderr)
        return 1
    try:
        sheets = compatible_sheets(workbook_path)
        if args.list_sheets:
            print("\n".join(sheets) if sheets else "没有找到新版模板 Sheet")
            return 0
        if args.all:
            if args.output:
                raise ValueError("--all 不能与 --output 同时使用")
            if not sheets:
                raise ValueError("工作簿内没有新版模板 Sheet")
            for sheet_name in sheets:
                destination = Path("output") / safe_filename(sheet_name)
                generate(workbook_path, sheet_name, destination / "index.html")
        else:
            output_path = Path(args.output) if args.output else Path("output") / "index.html"
            generate(workbook_path, args.sheet, output_path)
        return 0
    except (OSError, ValueError) as error:
        print(f"错误：{error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
