# 旅行计划项目约定

## 项目目标

本项目用于维护个人旅行计划。每次旅行的数据保存为一个 YAML 文件，这些文件是唯一可信的数据源。脚本根据 YAML 自动生成多旅行 PWA，适合在 iPhone 上浏览和离线使用，并通过 GitHub Pages 发布。

生成的 HTML 不作为编辑入口。修改行程主要由 AI 修改 YAML 完成；人工编辑主要在 Excel 中修改同名的 xlsx，由 `sync` 同步回 YAML，偶尔也会在 VS Code 中直接修改 YAML。

## 目录

| 路径 | 用途 |
| --- | --- |
| `trips/<id>.yaml` | 每次旅行一个文件。文件名就是旅行 id 和网址路径，只能使用小写字母、数字和 `-`。 |
| `trips/<id>.xlsx` | 同一次旅行的 Excel 编辑视图，由 `sync` 生成并提交，带下拉、公式和条件格式。隐藏 Sheet `_meta` 记录同步标记。 |
| `scripts/trip_model.py` | 唯一的数据定义，包括 pydantic 模型、YAML 读写和花销统计。 |
| `scripts/trip_xlsx.py` | YAML 与 xlsx 的导出、解析和同步。`sync` 按两边内容和 `_meta` 自动判断方向；`--check` 只检查是否一致；`--watch` 常驻监听并自动同步。 |
| `.githooks/pre-commit` | 提交前执行 `sync`，并把写过的文件加入本次提交。 |
| `scripts/generate_site.py` | 由 YAML 生成站点。加 `--check` 时只做校验并输出摘要。 |
| `scripts/export_schema.py` | 导出 `schema/trip.schema.json`，供 VS Code 补全和校验使用。 |
| `templates/` | Jinja2 页面模板，以及 `assets/` 下的 CSS 和 JS。 |
| `tests/` | pytest 测试。 |
| `docs/specs/`、`docs/plans/` | 设计文档和实施计划。 |

`output/` 是生成产物，已被 Git 忽略，不提交，也不手动编辑。

## YAML 字段

旅行级字段：`title`、`start`、`end`（必填）；`travelers`、`notes`（可选）；`days`（必填）。

每天的字段：`date`（必填）；`label`（可选，例如“苏州 → 上海”）；`items`（必填，可以是空列表）。

事项字段：

| 字段 | 说明 |
| --- | --- |
| `title` | 必填，事项名称。 |
| `type` | 必填，取值为交通、住宿、游览、餐饮、日常、购物、其他之一。 |
| `start` / `end` | 时间，必须写成带引号的 `"HH:MM"`。`end` 早于 `start` 表示跨过午夜。 |
| `note` | 路线、车次、地址、营业时间等补充信息。 |
| `price` / `qty` | 单价和数量，消费金额由二者相乘得出。 |
| `cost` | 总价，仅在无法拆成单价 × 数量时填写。不能与 `price` / `qty` 同时出现。 |
| `booking_at` | 开售或预约时间，写作 `2026-08-01` 或 `2026-08-02 14:00`。 |
| `status` | 默认“无需预订”，其余可选值为待购买、待预约、已购买、已预约、已完成、已取消。 |
| `place` | 地图 POI 关键词。为空时页面不显示地图按钮。 |

规则：

- 禁止出现未知字段。每天的日期必须落在旅行起止日期之内，并且严格递增。
- 时间轴是行程的核心：事项按执行时间排列，不按类别拆开。
- 待办指状态为“待购买”或“待预约”的事项。已付指状态为已购买、已预约或已完成的事项。已取消的事项不计入任何统计。
- 金额一律按人民币计。
- 不要删除有实际意义的未定事项。只有 `0`、空值这类纯占位内容可以清理。
- `place` 要填写地图能准确识别的 POI。不确定时留空，不要凭名称猜测。

## 修改流程

1. 修改 `trips/<id>.yaml`，或在 Excel 中修改 `trips/<id>.xlsx`。改 YAML 时尽量保持原有字段顺序；新增事项时按时间插入到正确位置。
2. 运行 `python scripts/trip_xlsx.py sync`，把改动同步到另一边。本机运行 watch 时保存后会自动同步，但 AI 改完 YAML 后仍必须手动执行一次（重复执行不会改写文件）。
3. 运行 `python scripts/generate_site.py --check` 和 `python scripts/trip_xlsx.py sync --check`，确认都通过，并核对输出的总额和待办数是否符合预期。
4. 如果修改了 `trip_model.py` 中的模型，运行 `python scripts/export_schema.py` 并提交新导出的 schema。
5. 修改代码或模板后，运行 `python -m pytest -q`。
6. 提交并推送到 `main`。YAML 和 xlsx 要放在同一次提交里，pre-commit 会再同步一次作为兜底；GitHub Actions 校验两边一致、运行测试后发布。

AI 修改行程时的约定：

- 只修改 YAML，不用代码直接改写 xlsx，xlsx 一律由 `sync` 生成。
- `sync` 报告冲突（两边都改过）时，不要自行加 `--prefer` 选择一边，先把变更摘要告诉用户，由用户决定以哪边为准。
- 从 xlsx 写回 YAML 时按规范格式重新输出，YAML 中的手写注释会丢失。需要长期保留的说明写进 `note` 或 `notes` 字段。

## 发布

- 工作流为 `.github/workflows/deploy.yml`，发布目标由仓库变量 `DEPLOY_TARGET` 决定。
- 正式访问地址：<https://hellsge.github.io/travel-plan/>。首页是旅行列表，每次旅行的页面位于 `/<id>/`，表格视图位于 `/<id>/table.html`。
- 发布后要检查 GitHub Actions 的运行状态、线上页面、manifest 和 Service Worker。
- 仓库本地提交身份使用 `hellsge <hellsge@qq.com>`。

## 页面规范

- 时间轴页用于出行途中查看，功能包括：按天导航、标记当前和下一事项、在高德 / Apple / 百度地图之间选择、复制、深色模式、离线提示，以及可折叠的花销统计。
- 表格视图用于桌面浏览和横向打印。
- 类型颜色只作用于类型单元格，状态颜色只作用于状态单元格，不给整行着色。
- 移动端主要按 iPhone 16（`393 × 852`）和 iPhone 16 Pro（`402 × 874`）验证。需要满足：保留 `viewport-fit=cover`，适配 `env(safe-area-inset-*)`，触控区域不小于 44px，页面不出现横向溢出。
- 视觉效果不能只看构建成功与否，必须按目标尺寸截图检查，浅色和深色模式都要检查。
