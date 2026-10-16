# 旅行计划

每次旅行的数据保存为 `trips/` 下的一个 YAML 文件，这些文件是唯一可信的数据源。每个 YAML 旁边还有一份同名的 xlsx，方便在 Excel 中预览和编辑，两边由脚本自动同步。仓库会根据 YAML 自动发布一个多旅行 PWA，可以在 iPhone 上离线浏览。

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

## 每台电脑的首次配置

在仓库根目录执行一次：

```bash
python -m pip install -r requirements-dev.txt
git config core.hooksPath .githooks
```

第二行启用 pre-commit hook：提交前自动同步 YAML 和 xlsx，并把同步写出的文件加入本次提交。遇到冲突、校验失败或文件被 Excel 占用时，会中止提交。这项配置只写入本仓库的 `.git/config`，因此每台电脑都要执行一次。

用 VS Code 打开仓库时，“Trip sync watch” 任务会在后台监听 `trips/`，保存 YAML 或 xlsx 后约 2 秒自动同步到另一边。第一次打开时，VS Code 会询问是否允许本文件夹的自动任务，选择允许即可；也可以从命令面板执行 **Tasks: Manage Automatic Tasks** 修改。不用 VS Code 时，可以在终端手动启动：

```bash
python scripts/trip_xlsx.py sync --watch
```

## 修改行程

有三种方式，改完都要同步。watch 运行时会自动同步，否则手动执行：

```bash
python scripts/trip_xlsx.py sync
```

- **在 Excel 中修改** `trips/<id>.xlsx`：类型、状态和时间都有下拉候选，时间也可以手动输入精确值，例如 `19:28`；顶部会即时汇总总额、已付、待付和待办数。浅蓝色日期分隔行下面的事项都归属这一天。要把事项挪到另一天，把整行剪切到那一天的分隔行下方即可。
- **让 AI 修改** `trips/*.yaml`。
- **在 VS Code 中修改 YAML**：按提示安装 YAML 插件（`redhat.vscode-yaml`）后，`type`、`status` 可以用 `Ctrl+Space` 从候选中选择；在 `items:` 下输入 `-` 并按 `Ctrl+Space` 可以插入“新事项”模板；字段写错时会实时显示红色波浪线。

新建一次旅行时，复制一份已有的 xlsx，按 id 规则重命名（只能用小写字母、数字和 `-`），在 Excel 中修改后运行 sync，就会生成对应的 YAML。

同步后运行校验，终端会输出每次旅行的总额和待办数：

```bash
python scripts/generate_site.py --check
python scripts/trip_xlsx.py sync --check
```

确认无误后提交并推送到 `main`。YAML 和 xlsx 要一起提交，GitHub Actions 校验两边一致、运行测试后再发布：

```bash
git add trips/
git commit -m "chore: update yunnan itinerary"
git push
```

### 同步规则

`sync` 会比较 YAML、xlsx，以及上次同步时记在 xlsx 隐藏 Sheet `_meta` 中的标记，从而判断改的是哪一边，再同步到另一边。

- 两边都改过时，判定为**冲突**，不会写任何文件，终端会输出两边的差异摘要。确认以哪边为准后，执行：

  ```bash
  python scripts/trip_xlsx.py sync <id> --prefer yaml   # 以 YAML 为准，覆盖 xlsx
  python scripts/trip_xlsx.py sync <id> --prefer xlsx   # 以 xlsx 为准，覆盖 YAML
  ```

- Excel 打开 xlsx 时，脚本无法写入这个文件，会提示“请先关闭 Excel”。watch 会每 5 秒重试一次，关闭 Excel 后自动写入。
- 从 xlsx 写回 YAML 时会按规范格式重新输出，YAML 中的手写注释会丢失。

字段含义和取值规则见 [CLAUDE.md](CLAUDE.md)。

### 用腾讯文档编辑

1. 把 `trips/<id>.xlsx` 上传到腾讯文档，在线编辑。
2. 编辑完下载为 xlsx，覆盖仓库中的 `trips/<id>.xlsx`，文件名保持不变。
3. 运行 `python scripts/trip_xlsx.py sync`。

腾讯文档导出的文件会丢掉隐藏 Sheet `_meta`，sync 无法判断改的是哪一边，因此只要内容和 YAML 不同就会判为冲突。确认这期间 YAML 没有被改过后，执行：

```bash
python scripts/trip_xlsx.py sync <id> --prefer xlsx
```

如果腾讯文档把公式转成了数值，没有关系：消费列等于单价 × 数量时会自动忽略；不相等时会报错，并指出是哪一行。

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
├── trips/                  # 每次旅行一个 YAML（唯一可信数据源）和一份同名 xlsx
├── schema/                 # 导出的 JSON Schema，供 VS Code 使用
├── scripts/
│   ├── trip_model.py       # 数据模型、校验、统计
│   ├── trip_xlsx.py        # YAML 与 xlsx 的导出、解析、同步
│   ├── generate_site.py    # YAML → 站点
│   ├── site_assets.py      # 图标、manifest、Service Worker
│   └── export_schema.py    # 导出 schema
├── templates/              # Jinja2 模板与 CSS/JS
├── tests/                  # pytest
├── docs/                   # 设计文档与实施计划
├── .githooks/              # pre-commit：提交前同步 YAML 与 xlsx
├── .vscode/                # YAML 插件配置和 watch 自动任务
└── .github/workflows/      # 自动校验、构建与发布
```
