# Excel → YAML 数据源迁移实施计划

**Goal:** 把旅行数据从 `travel.xlsx` 迁移为每次旅行一个 YAML 文件，并由这些 YAML 生成多旅行 PWA 网站。

**Architecture:**

- `scripts/trip_model.py` 是唯一的数据定义，包含 pydantic 模型、YAML 读写、错误汇总和花销统计。
- `scripts/generate_site.py` 用 Jinja2 模板把所有旅行渲染成静态站点。
- `scripts/migrate_xlsx.py` 是一次性迁移脚本，核对通过后与 `travel.xlsx` 一起删除。

**Tech Stack:** Python 3.10（本地）/ 3.12（CI）、pydantic 2.13、ruamel.yaml 0.19（YAML 1.2）、Jinja2 3.1、pytest 9。

**Spec:** [docs/specs/2026-09-30-yaml-data-source-design.md](../specs/2026-09-30-yaml-data-source-design.md)

## Global Constraints

- 代码必须同时兼容 Python 3.10 和 3.12。
- 运行依赖只有 pydantic、ruamel.yaml 和 jinja2。pytest 放在 `requirements-dev.txt`。openpyxl 只在迁移阶段临时使用。
- 字段名用英文，枚举值用中文。待办状态为“待购买”和“待预约”；已付状态为“已购买”“已预约”和“已完成”；已取消的事项不计入统计。
- 金额统一按人民币计，显示格式为 `¥1,234.56`。
- iPhone 页面的约束：设置 `viewport-fit=cover`，适配 safe-area，触控区域不小于 44px，不出现横向溢出。
- 类型的颜色只作用于类型单元格，状态的颜色只作用于状态单元格。
- commit message 用英文，遵循 Conventional Commits，提交身份为 `hellsge <hellsge@qq.com>`。提交和推送都要先经用户同意。

## Tasks

每个任务都是先写测试、再实现，完成后运行 `python -m pytest -q`，全部通过才算完成。

### Task 1：数据模型与读取 ✅

- 文件：`scripts/trip_model.py`、`tests/conftest.py`、`tests/test_model.py`
- 接口：
  - `load_trip(path) -> LoadedTrip`
  - `load_trips(dir) -> (list[LoadedTrip], list[str])`
  - `dump_trip(Trip) -> str`
  - `Item.amount`、`Item.is_todo`
  - `TripLoadError.messages`
  - 常量 `ITEM_TYPES`、`STATUSES`、`TODO_STATUSES`、`PAID_STATUSES`、`TIME_PATTERN`
- 测试覆盖：
  - 各类非法输入的反例：未知字段、数字时间、时间格式错误、非法枚举值、`cost` 与 `price` 互斥、日期超出范围、日期不递增、文件名不合法、YAML 语法错误
  - 多个文件的错误汇总
  - 写回后再读取，结果一致

### Task 2：花销统计 ✅

- 文件：`trip_model.py` 新增 `TripStats` 和 `compute_stats(trip) -> TripStats`；测试写在 `tests/test_stats.py`。
- `TripStats` 包含以下字段：
  - `total`
  - `per_person`（可为 None）
  - `by_type`：按金额降序排列的 `list[(type, amount, ratio)]`
  - `paid`
  - `unpaid`
  - `todo_count`
  - `day_totals`：`dict[date, float]`
- 已取消的事项既不计入金额，也不计入待办数。

### Task 3：JSON Schema 与 VS Code 配置 ✅

- 文件：`scripts/export_schema.py`、`schema/trip.schema.json`、`.vscode/extensions.json`、`.vscode/settings.json`，测试写在 `tests/test_schema.py`。
- Schema 的要求：
  - 每个字段都有中文 `description`；
  - 时间字段提供 10 分钟间隔的 `examples`；
  - 提供两个 `defaultSnippets`：新事项、新的一天。
- 测试内容：
  - 仓库里提交的 schema 与重新导出的结果完全一致；
  - 时间字段的 pattern 能接受 `19:28` 这样的非整点时间。

### Task 4：迁移脚本并生成 `trips/*.yaml` ✅

- 新增 `scripts/migrate_xlsx.py`，按 spec 第 2 节的规则迁移。
- 迁移时核对每次旅行的总消费和事项数。
- 清明这次旅行另外输出一份差异清单。
- 迁移完成后运行 `load_trips`，确认全部通过校验，再请用户抽查迁移结果。
- 新增 `tests/test_trips_data.py`，校验 `trips/` 下的全部数据。

### Task 5：站点生成器与模板 ✅

- 新增 `scripts/generate_site.py`，包含 CLI 和渲染逻辑；新增 `scripts/site_assets.py`，负责图标、manifest 和 Service Worker。
- 新增模板 `templates/` 和静态资源 `templates/assets/`，测试写在 `tests/test_site.py`。
- 生成的页面包括首页、时间轴页、表格视图和花销统计区块，现有时间轴页的功能全部保留。
- 命令行 `--check` 用于输出校验结果和统计摘要。
- 测试内容：
  - 首页卡片数量正确；
  - 页面之间的链接有效；
  - 内容做了 HTML 转义；
  - Service Worker 的缓存清单覆盖所有页面；
  - `--check` 在数据有误时返回非零退出码。

### Task 6：视觉验证与编辑体验验证（截图已完成；VS Code 下拉补全用户实测未生效，待排查）

- 在 393×852 和 402×874 两种尺寸下，分别截图浅色、深色两种模式；桌面宽度下截图表格视图，并查看打印预览。
- 在 VS Code 中实际验证下拉选择和代码片段。
- 截图用的临时文件在验证后删除。

### Task 7：CI、文档与清理（待提交、推送和线上验证）

- 修改 `.github/workflows/deploy.yml`，调整触发路径，并加入 check 和 pytest 步骤。
- 重写 `CLAUDE.md` 和 `README.md`。
- 删除 `scripts/generate_pwa.py`、`scripts/migrate_xlsx.py` 和 `travel.xlsx`，并从依赖中移除 openpyxl。
- 请用户确认后再提交和推送，推送后验证线上页面。
