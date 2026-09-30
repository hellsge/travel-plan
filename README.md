# 旅行计划

每次旅行的数据保存为 `trips/` 下的一个 YAML 文件，这些文件是唯一可信的数据源。仓库会根据它们自动发布一个多旅行 PWA，可以在 iPhone 上离线浏览。

## 在线地址

**[打开旅行计划](https://hellsge.github.io/travel-plan/)**

首页列出所有旅行，进行中和即将出发的排在最前。点进任意一次旅行可以看到：

- **时间轴视图**：适合出行途中查看。
- **表格视图**：适合在电脑上查看全程或横向打印。
- **花销统计**：包括总额、人均、按类型汇总，以及已付和待付金额。

## iPhone 安装

1. 用 Safari 打开上面的在线地址。
2. 点击“分享”→“添加到主屏幕”，并开启“作为网页 App 打开”。
3. 从主屏幕打开一次，完成离线缓存。

之后行程和统计都可以离线查看，地图搜索仍然需要网络。联网打开时会自动获取最新内容。

## 修改行程

推荐让 AI 修改 `trips/*.yaml`。手动修改时，用 VS Code 打开仓库，并按提示安装 YAML 插件（`redhat.vscode-yaml`），插件会提供以下辅助：

- `type`、`status` 可以用 `Ctrl+Space` 从下拉候选中选择。
- 时间有每 10 分钟一个的候选值，也可以手动输入精确时间，例如 `"19:28"`。
- 在 `items:` 下输入 `-` 并按 `Ctrl+Space`，可以插入“新事项”模板。
- 字段写错或取值不合法时，会实时显示红色波浪线。

改完后先运行校验，终端会输出每次旅行的总额和待办数：

```bash
python scripts/generate_site.py --check
```

确认无误后提交并推送到 `main`，GitHub Actions 会自动校验、测试并发布：

```bash
git add trips/
git commit -m "chore: update yunnan itinerary"
git push
```

字段含义和取值规则见 [CLAUDE.md](CLAUDE.md)。

## 本地生成与测试

```bash
python -m pip install -r requirements-dev.txt
python scripts/generate_site.py      # 生成到 output/
python -m pytest -q
```

修改 `scripts/trip_model.py` 中的数据模型后，需要重新导出 schema：

```bash
python scripts/export_schema.py
```

## 发布平台配置

发布目标由仓库变量 `DEPLOY_TARGET` 决定，在 **Settings → Secrets and variables → Actions → Variables** 中设置。可选值：

- `github-pages`：发布到 GitHub Pages（当前默认）
- `tencent`：发布到腾讯云 CloudBase 静态托管，需要使用按量计费环境
- `none`：只构建，不发布

使用腾讯云时，还要在 **Secrets** 中配置以下三项：

- `TENCENT_SECRET_ID`
- `TENCENT_SECRET_KEY`
- `TENCENT_ENV_ID`

发布状态可以在 [Actions](https://github.com/hellsge/travel-plan/actions) 页面查看。

## 项目结构

```text
travel-plan/
├── trips/                  # 每次旅行一个 YAML（唯一可信数据源）
├── schema/                 # 导出的 JSON Schema，供 VS Code 使用
├── scripts/
│   ├── trip_model.py       # 数据模型、校验、统计
│   ├── generate_site.py    # YAML → 站点
│   ├── site_assets.py      # 图标、manifest、Service Worker
│   └── export_schema.py    # 导出 schema
├── templates/              # Jinja2 模板与 CSS/JS
├── tests/                  # pytest
├── docs/                   # 设计文档与实施计划
└── .github/workflows/      # 自动校验、构建与发布
```
