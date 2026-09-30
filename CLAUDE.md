# 旅行计划项目约定

## 项目目标

本项目用于维护个人旅行计划。每次旅行的数据保存为一个 YAML 文件，这些文件是唯一可信的数据源。脚本根据 YAML 自动生成多旅行 PWA，适合在 iPhone 上浏览和离线使用，并通过 GitHub Pages 发布。

生成的 HTML 不作为编辑入口。修改行程主要由 AI 修改 YAML 完成，偶尔也会在 VS Code 中手动修改。

## 目录

| 路径 | 用途 |
| --- | --- |
| `trips/<id>.yaml` | 每次旅行一个文件。文件名就是旅行 id 和网址路径，只能使用小写字母、数字和 `-`。 |
| `scripts/trip_model.py` | 唯一的数据定义，包括 pydantic 模型、YAML 读写和花销统计。 |
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

1. 修改 `trips/*.yaml`。尽量保持原有字段顺序；新增事项时按时间插入到正确位置。
2. 运行 `python scripts/generate_site.py --check`，确认校验通过，并核对输出的总额和待办数是否符合预期。
3. 如果修改了 `trip_model.py` 中的模型，运行 `python scripts/export_schema.py` 并提交新导出的 schema。
4. 修改代码或模板后，运行 `python -m pytest -q`。
5. 提交并推送到 `main`，由 GitHub Actions 校验、测试并发布。

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
