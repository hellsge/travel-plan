"""与模板无关的站点产物：PNG 图标、manifest、Service Worker。"""

from __future__ import annotations

import json
import struct
import zlib
from pathlib import Path

ICONS = {"icon-192.png": 192, "icon-512.png": 512, "apple-touch-icon.png": 180}


def _png_chunk(chunk_type: bytes, data: bytes) -> bytes:
    return (
        struct.pack(">I", len(data))
        + chunk_type
        + data
        + struct.pack(">I", zlib.crc32(chunk_type + data) & 0xFFFFFFFF)
    )


def _icon_png(size: int) -> bytes:
    """蓝底白圆的旅行图标，不依赖第三方图像库。"""
    rows = bytearray()
    center = (size - 1) / 2
    radius = size * 0.39
    for y in range(size):
        rows.append(0)
        for x in range(size):
            inside = ((x - center) ** 2 + (y - center) ** 2) ** 0.5 <= radius
            rows.extend((255, 255, 255, 255) if inside else (37, 99, 235, 255))
    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", header)
        + _png_chunk(b"IDAT", zlib.compress(bytes(rows), 9))
        + _png_chunk(b"IEND", b"")
    )


def write_icons(output: Path) -> list[str]:
    icons_dir = output / "icons"
    icons_dir.mkdir(parents=True, exist_ok=True)
    for name, size in ICONS.items():
        (icons_dir / name).write_bytes(_icon_png(size))
    return [f"icons/{name}" for name in ICONS]


def write_manifest(output: Path) -> str:
    manifest = {
        "name": "旅行计划",
        "short_name": "旅行计划",
        "description": "由 YAML 自动生成的离线旅行计划",
        "lang": "zh-CN",
        "start_url": "./",
        "scope": "./",
        "display": "standalone",
        "background_color": "#f4f7fb",
        "theme_color": "#2563eb",
        "icons": [
            {"src": "icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
        ],
    }
    (output / "manifest.webmanifest").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return "manifest.webmanifest"


def write_service_worker(output: Path, assets: list[str], version: str) -> None:
    """网络优先、离线回退；导航请求离线时回退到首页。"""
    urls = [f"./{path}" for path in assets]
    script = f"""const CACHE_NAME = {json.dumps(f"travel-{version}")};
const ASSETS = {json.dumps(urls, ensure_ascii=False, indent=2)};

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
  const navigate = event.request.mode === 'navigate';
  event.respondWith(fetch(event.request, navigate ? {{cache: 'reload'}} : {{}}).then(response => {{
    if (response.ok && new URL(event.request.url).origin === self.location.origin) {{
      const copy = response.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
    }}
    return response;
  }}).catch(() => caches.match(event.request, {{ignoreSearch: true}})
    .then(cached => cached || (navigate ? caches.match('./') : undefined))));
}});
"""
    (output / "service-worker.js").write_text(script, encoding="utf-8")
