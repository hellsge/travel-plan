# 旅行计划

以 [travel.xlsx](travel.xlsx) 为唯一数据源，自动发布“26暑期云南”行程为可离线使用的 PWA，供 iPhone 浏览。

## 在线地址

**[打开旅行计划](https://hellsge.github.io/travel-plan/)**

## iPhone 安装

1. 使用 Safari 打开上面的在线地址。
2. 等待页面完整加载。
3. 点击“分享”→“添加到主屏幕”。
4. 启用“作为网页 App 打开”。
5. 从主屏幕打开一次，完成离线缓存。

行程、预算和待办筛选支持离线使用；地图搜索仍需要网络。检测到新版时，网页 App 会显示“立即刷新”提示。

## 更新行程

1. 修改 [travel.xlsx](travel.xlsx) 中的“26暑期云南”Sheet。
2. 提交并推送到 `main`：

```bash
git add travel.xlsx
git commit -m "Update summer Yunnan itinerary"
git push
```

1. GitHub Actions 自动读取 Excel、生成 PWA 并发布。
1. iPhone 联网打开网页 App，看到“行程已有新版本”后点击“立即刷新”。

## 发布平台配置

发布目标由 GitHub 仓库变量 `DEPLOY_TARGET` 决定，在仓库 **Settings → Secrets and variables → Actions → Variables** 中设置（当前为 `github-pages`）：

- `github-pages`：发布到 GitHub Pages（当前默认）
- `tencent`：发布到腾讯云 CloudBase 静态托管（需按量计费环境）
- `none`：只构建、不发布

使用腾讯云时，还需在 **Secrets** 中配置：

| 名称 | 说明 |
| --- | --- |
| `TENCENT_SECRET_ID` | 腾讯云 CAM 密钥 SecretId |
| `TENCENT_SECRET_KEY` | 腾讯云 CAM 密钥 SecretKey |
| `TENCENT_ENV_ID` | CloudBase 静态托管环境 ID |

配置后，推送 [travel.xlsx](travel.xlsx) 会自动重新构建并发布。

## 地点和地图

新版模板的 L 列“地点 / 地图关键词”专用于地图搜索。只有填写该列的事项才显示地图操作，路线说明、营业时间等内容继续填写在 F 列“路线 / 备注”。

点击“打开地图”可选择高德地图、Apple 地图或百度地图，网页 App 会记住上次使用的地图。地点尽量填写可被地图准确识别的 POI，例如“昆明长水国际机场”或“大理古城”。

可在仓库的 [Actions](https://github.com/hellsge/travel-plan/actions) 页面查看发布状态。

## 项目结构

```text
travel-plan/
├── travel.xlsx                      # 唯一可信数据源
├── scripts/
│   └── generate_pwa.py              # Excel → PWA 生成器
├── .github/workflows/
│   └── deploy.yml                   # 自动构建并按 DEPLOY_TARGET 发布
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
