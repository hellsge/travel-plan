# 旅行计划

以 [travel.xlsx](travel.xlsx) 为唯一数据源，自动发布“26暑期云南”行程到 GitHub Pages，并支持在 iPhone 上安装为可离线使用的 PWA。

## 在线地址

**[打开旅行计划](https://hellsge.github.io/travel-plan/)**

## iPhone 安装

1. 使用 Safari 打开上面的在线地址。
2. 等待页面完整加载。
3. 点击“分享”→“添加到主屏幕”。
4. 启用“作为网页 App 打开”。
5. 从主屏幕打开一次，完成离线缓存。

行程、预算和待办筛选支持离线使用；地图搜索仍需要网络。

## 更新行程

1. 修改 [travel.xlsx](travel.xlsx) 中的“26暑期云南”Sheet。
2. 提交并推送到 `main`：

```bash
git add travel.xlsx
git commit -m "Update summer Yunnan itinerary"
git push
```

1. GitHub Actions 自动读取 Excel、生成 PWA 并发布。
1. iPhone 联网打开网页一次，即可获取新版本。

可在仓库的 [Actions](https://github.com/hellsge/travel-plan/actions) 页面查看发布状态。

## 项目结构

```text
travel-plan/
├── travel.xlsx                      # 唯一可信数据源
├── scripts/
│   └── generate_pwa.py              # Excel → PWA 生成器
├── .github/workflows/
│   └── deploy-pages.yml             # GitHub Pages 自动发布
├── requirements.txt                 # Python 依赖
└── README.md
```

`output/` 是本地生成目录，已被 Git 忽略；线上构建也会临时生成该目录，不需要提交。

## 开发与验证

安装依赖并生成静态文件：

```bash
python -m pip install -r requirements.txt
python scripts/generate_pwa.py
```

生成结果位于 `output/`。如需发布其他旅行，可显式指定 Sheet：

```bash
python scripts/generate_pwa.py --sheet "模板验证-清明重庆"
```

直接修改生成的 HTML 没有意义，下次构建时会被 Excel 内容覆盖。
