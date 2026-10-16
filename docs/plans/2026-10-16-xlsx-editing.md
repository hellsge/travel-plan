# Excel 编辑视图与双向同步实施计划

## Context

迁移到 YAML 之后，人工编辑变难了：没有下拉、没有即时汇总，也不能再传到腾讯文档上编辑。设计文档 [docs/specs/2026-10-16-xlsx-editing-design.md](docs/specs/2026-10-16-xlsx-editing-design.md) 已经确认，做法如下：

- 每次旅行在 `trips/<id>.yaml` 旁边提交一份格式化的 `trips/<id>.xlsx`。
- 用 `sync` 命令按三方哈希自动判断同步方向。
- 本机运行 watch，保存后自动同步；pre-commit 和 CI 再各做一层兜底。
- YAML 仍是唯一可信数据源。

批准后，第一步是把本计划另存为 `docs/plans/2026-10-16-xlsx-editing.md`，与现有计划文档放在一起，风格保持一致，然后按任务顺序执行。

**Tech Stack:** Python 3.10（本地）/ 3.12（CI）、openpyxl 3.1.5（本地已安装）、watchdog（新增并固定版本）、现有的 pydantic、ruamel.yaml 和 pytest。

## Global Constraints

- 同时兼容 Python 3.10 和 3.12，不新增其他工具。
- 复用 `scripts/trip_model.py` 中已有的接口，不重复定义规则：
  - 模型和常量：`Trip`/`Day`/`Item`、`ITEM_TYPES`、`STATUSES`、`DEFAULT_STATUS`、`TODO_STATUSES`、`PAID_STATUSES`、`CANCELLED_STATUS`、`TRIP_ID_PATTERN`、`RESERVED_IDS`
  - 读写与统计：`load_trip`、`dump_trip`、`compute_stats`、`TripLoadError`、`_describe`（报错文案）
- 颜色取自 `templates/assets/` 的 CSS：类型用 `.type-交通{--type:#2673d9}` 等，状态用 `.status-待购买{background:#fef0b8}` 等。只给类型单元格和状态单元格着色。
- 每个任务都先写测试，再写实现，以 `python -m pytest -q` 全部通过为完成标准。
- 未经用户同意不 commit 或 push。commit message 用英文，遵循 Conventional Commits。
- 工作区里已有的 `docs/plans/2026-09-30-yaml-data-source.md` 改动和未跟踪的 `trips/test-template.yaml`（该文件本身不合法）都不动。sync 会把非法 YAML 当作这次旅行的错误报出来，执行过程中如果它挡住了验证，再问用户怎么处理。

## Tasks

### Task 1：导出 Trip → xlsx

- 新文件：`scripts/trip_xlsx.py`（导出、解析和同步都放在这里，每部分的函数各自独立）、`tests/test_xlsx_export.py`
- 接口：
  - `digest(trip) -> str`：`sha256(dump_trip(trip))`
  - `build_workbook(trip, digest) -> Workbook`
  - `write_workbook(wb, path)`：先写同目录的临时文件，再 `os.replace`；遇到 `PermissionError` 时转成 `FileLockedError`
- 版式按 spec 第 2 节：
  - 顶部摘要区：可编辑字段；总额、已付、待付、待办用 `SUMPRODUCT`/`COUNTIFS` 公式计算，并排除“已取消”。
  - 日期分隔行：A 列为日期值，B 至 L 列合并，显示 label。
  - 事项行：A–L 列。
  - 隐藏 Sheet `_lists`：时间下拉 00:00–23:50，以及类型、状态的选项。下拉的 `showErrorMessage=False`，允许手动输入 `19:28`。
  - 隐藏 Sheet `_meta`：存放 digest。
  - D、K 列按“单元格值等于”设置条件格式。
  - 冻结窗格，设置列宽，打印区域设为横向。
  - 把 `wb.properties.created/modified` 固定为常量，减少无意义的元数据变化。
- 测试：
  - 用 openpyxl 读回导出文件，断言 D、K、B、C 列有数据验证。
  - I 列是 `=G*H` 公式；`cost` 事项的 I 列是数值。
  - 条件格式规则数量正确，`_meta` 中是 digest。
  - 空的一天也有分隔行。

### Task 2：解析 xlsx → Trip

- 接口：
  - `parse_workbook(path) -> Trip`
  - 出错时抛 `XlsxParseError(messages)`，每条信息形如“第 N 行 D 列 类型：…”
- 规则按 spec 第 3 节：
  - 区分行类型：分隔行、空行、事项行。
  - 事项归属以上方最近的分隔行为准。A 列日期不一致时，通过 `warnings` 列表返回提示。
  - 金额：G 或 H 有值时按单价×数量读取，否则读 I 列作为 `cost`。
  - 时间：接受 `datetime.time`、`datetime`、Excel 浮点数和 `HH:MM` 文本，四舍五入到分钟。
  - `booking_at` 时间部分为 00:00 时按纯日期处理。
  - 状态留空表示无需预订。
  - 交给 `Trip.model_validate` 校验，再把 pydantic 的 `loc`（`days[i].items[j].type`）反查成 Excel 行号和列名，记录方式是解析时维护一张 `(day_idx, item_idx) -> row` 映射表。
- 测试 `tests/test_xlsx_parse.py`：
  - **往返测试**：对 `trips/` 下每个合法 YAML 执行 export 后再 parse，digest 必须相同。这是核心用例。
  - 事项挪到另一天、公式被转成数值、只填总价、文本格式的时间、空行、第一条分隔行之前出现事项（报错）。
  - 非法类型、单价与总价同时填写、日期越界，报错中的行号正确。

### Task 3：同步状态机与 CLI

- 接口：
  - `sync_trip(trip_id, trips_dir, prefer=None, check=False) -> SyncResult`：返回 `action`、`changes` 摘要、`written` 文件列表、`error`
  - `sync_all(...)`
  - `main(argv)`：参数为 `sync [id] [--check] [--prefer yaml|xlsx] [--watch]`，`--trips` 默认指向 `ROOT/"trips"`，写法与 `generate_site.parse_args` 一致
- 按 spec 第 4 节的状态表实现：
  - 只有 xlsx：生成 YAML。只有 YAML：导出 xlsx。
  - 发生冲突时，不写任何文件。
  - 变更摘要列出新增、删除、修改的事项数，以及总额和待办数的前后变化，数字来自 `compute_stats`。
  - 有任何一次旅行失败，退出码为 1，但已成功的旅行照常写入。
  - `--check`：任意一次旅行 X≠Y 或缺少某一边，退出码为 1，不写文件。
- 测试 `tests/test_xlsx_sync.py`：在 `tmp_path` 中构造每一种状态，至少各一个用例，包括冲突时文件保持不变、`--prefer` 两个方向、`FileLockedError` 的提示文案、非法文件名。

### Task 4：watch 模式

- 把调度逻辑写成不依赖 watchdog 的 `Debouncer`：
  - `notify(trip_id, now)` 记录事件；`due(now)` 返回静默满 1.5 秒的 trip_id。
  - 文件被占用的 trip 每 5 秒重试一次。
- 路径到 trip_id 的映射忽略 `~$*`、临时文件和非 `.yaml`/`.xlsx` 文件。
- `run_watch()`：用 watchdog 的 Observer 加上 `time.sleep(0.2)` 循环。启动时先完整执行一次 sync。出错时打印信息，不退出。自身写入会再触发事件，下一轮按 digest 判定为已同步。
- 每次写入后打印一行，包括时间、id、同步方向和变更摘要。
- `requirements.txt` 加入固定版本的 `openpyxl==3.1.5` 和 `watchdog`（安装时确认版本号）。
- 测试 `tests/test_xlsx_watch.py`：
  - 多次事件合并成一次。
  - 静默不足时不触发。
  - 占用后重试。
  - 忽略锁文件。
  - 出错后后续事件仍会处理。

### Task 5：hook、VS Code task、CI

- `.githooks/pre-commit`：用 POSIX sh，Windows 上由 Git Bash 执行。
  - 先 `export PYTHONIOENCODING=utf-8`：Git Bash 下 Python 的 stdout 默认是 gbk，中文报错会乱码。
  - 执行 `python scripts/trip_xlsx.py sync`，失败时 `exit 1`。
  - 成功后用 `git add` 加入 sync 写过的文件。CLI 用 `--print-written` 输出文件路径，每行一个。
  - 在终端列出这些文件。
- `.vscode/tasks.json`：任务名 “Trip sync watch”，配置 `runOptions.runOn: folderOpen`、`isBackground: true`，使用专用面板。
- `.gitignore` 加入 `~$*.xlsx`。
- `.github/workflows/deploy.yml`：在 “Validate trips” 之后加一步 `python scripts/trip_xlsx.py sync --check`。
- 执行 `python scripts/trip_xlsx.py sync`，为现有 4 次旅行生成 xlsx。

### Task 6：文档

- `CLAUDE.md`：
  - 目录表加入 `trips/<id>.xlsx` 和 `scripts/trip_xlsx.py`。
  - 修改流程改为：修改 YAML 或 xlsx → sync（watch 运行时自动执行）→ `--check` → pytest → 提交。
  - 写明 AI 只修改 YAML，改完必须执行 sync。
- `README.md`：
  - 每台 PC 各执行一次 `git config core.hooksPath .githooks`。
  - 在 VS Code 中允许自动任务。
  - 腾讯文档往返流程：上传 → 编辑 → 下载并覆盖 `trips/<id>.xlsx` → sync。下载的文件缺少 `_meta`，如果两边内容不同会判为冲突，这时用 `--prefer xlsx` 指定以表格为准。
- 更新 spec 中第 1 节的依赖说明，`openpyxl` 和 `watchdog` 都要写上。

## Verification

1. 执行 `python -m pytest -q`，全部通过，包括原有测试和新增的 4 个测试文件。
2. 执行 `python scripts/generate_site.py --check` 和 `python scripts/trip_xlsx.py sync --check`，都通过；总额和待办数与修改前一致。
3. 手工 watch：运行 VS Code task，然后：
   - 用 Excel 修改一处状态并保存，YAML 应在约 2 秒内更新；
   - 修改 YAML，xlsx 应随之刷新；
   - Excel 开着文件时修改 YAML，应提示文件被占用并重试，关闭 Excel 后自动写入。
4. 截图：用 Excel 打开 `2026-summer-yunnan.xlsx`，按 `gui-screenshot-verify` 的流程用 PrintWindow 截图，检查下拉、颜色、公式汇总和冻结窗格。临时截图文件验证后删除。
5. pre-commit：只修改 xlsx 后直接运行 `sh .githooks/pre-commit`，用 `git diff --cached --name-only` 确认 YAML 已被加入暂存区，然后执行 `git restore --staged` 恢复。不执行真正的 commit。
6. 腾讯文档往返由用户手工验证一次：下载后执行 `sync --prefer xlsx`，内容能写回。
