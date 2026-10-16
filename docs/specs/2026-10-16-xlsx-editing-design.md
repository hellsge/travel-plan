# Excel 编辑视图与双向同步设计

## 背景与目标

迁移到 YAML 后（见 [2026-09-30-yaml-data-source-design.md](2026-09-30-yaml-data-source-design.md)），AI 修改数据变得可靠了，但人工编辑和预览变难了：

- 类型、状态、时间没有下拉可选；
- 金额无法即时汇总，必须运行脚本或打开网页才能看到；
- 无法像以前那样把表格上传到腾讯文档编辑。

使用习惯：在电脑上用 Excel 预览和编辑，只有出门在外时才用手机网页。两台 PC 通过 Git 同步仓库，偶尔把 Excel 上传到腾讯文档编辑后再拿回来。

目标：

1. 每次旅行在 YAML 之外再提交一份格式化的 `.xlsx`，供人预览和编辑。表格提供下拉、公式和条件格式，版式沿用迁移前的旧模板。
2. 用一条 `sync` 命令自动判断同步方向，YAML 和 xlsx 两边都可以作为编辑入口。
3. 在本机保存 YAML 或 xlsx 后自动同步（watch 模式），提交前再同步一次作为兜底。CI 检查两边是否一致，不一致就不发布。
4. YAML 仍然是**唯一可信数据源**：网站生成、校验和统计都只读 YAML，xlsx 中的修改必须先同步回 YAML 才算生效。

## 不在本次范围

- 接入腾讯文档 API 自动拉取或推送；
- 在 PWA 内直接编辑行程；
- 保留 YAML 中的手写注释（从 xlsx 写回 YAML 时会按规范格式重新输出）；
- 对一次旅行中的不同字段做细粒度的三方合并（两边都改过时，整份文件选一边为准）。

## 1. 文件与命令

```text
trips/
  2026-summer-yunnan.yaml    # 唯一可信数据源
  2026-summer-yunnan.xlsx    # 编辑和预览视图，同样提交
scripts/
  trip_xlsx.py               # 新增：导出、解析、同步
.githooks/
  pre-commit                 # 新增：提交前执行 sync
```

- `python scripts/trip_xlsx.py sync`：同步 `trips/` 下所有旅行。
- `python scripts/trip_xlsx.py sync --check`：只检查两边是否一致，不写文件；不一致时返回非 0。供 CI 使用。
- `python scripts/trip_xlsx.py sync <id> --prefer yaml|xlsx`：解决冲突时指定以哪一边为准。
- `python scripts/trip_xlsx.py sync --watch`：常驻监听，见第 6 节。
- `python scripts/trip_xlsx.py sync --print-written`：stdout 只输出写过的文件路径，每行一个，其余提示改到 stderr。供 pre-commit 使用。
- 依赖：`requirements.txt` 新增 `openpyxl==3.1.5` 和 `watchdog==6.0.0`。watchdog 只在 `--watch` 时才导入，未安装时其他命令照常可用。
- `.gitignore` 新增 `~$*.xlsx`（Excel 锁文件）。

xlsx 的文件名规则与 YAML 相同，必须符合 `TRIP_ID_PATTERN`。`trips/` 中出现不符合规则的 xlsx 时报错。

## 2. 表格版式

整次旅行放在一张大 Sheet 里，Sheet 名为旅行标题。另有两个隐藏 Sheet：`_meta` 存放同步哈希，`_lists` 存放下拉选项。

### 顶部摘要区

| 内容 | 可编辑 | 说明 |
| --- | --- | --- |
| 标题、开始日期、结束日期 | 是 | 对应 `title` / `start` / `end` |
| 同行人数、备注 | 是 | 对应 `travelers` / `notes`，可以留空 |
| 预计总消费、已付、待付、待办数 | 否（公式） | 口径与 `compute_stats` 一致，排除“已取消” |

公式计算的是 Excel 中的即时数字。如果和 `--check` 输出不一致，以 `--check` 为准。测试中会比较两者是否相同。

### 日期分隔行

每天一行，空的一天也保留。A 列为真正的 Excel 日期值，B 至 L 列合并，填写当天的 `label`。显示样式醒目，例如“07-20 周一”加 label。日期和 label 都可以编辑：复制一行分隔行并改掉日期，就是新增一天。

### 事项行（A–L 列沿用旧模板）

| 列 | 字段 | 输入方式 |
| --- | --- | --- |
| A 日期 | 显示用 | 导出时填上所属日期，样式弱化 |
| B 开始 / C 结束 | `start` / `end` | Excel 时间值，显示为 `hh:mm`；提供 10 分钟间隔的下拉，也可以手动输入 `19:28` |
| D 类型 | `type` | 下拉选择，条件格式只给这一格着色 |
| E 项目 | `title` | 文本 |
| F 路线 / 备注 | `note` | 文本，自动换行 |
| G 单价 / H 数量 | `price` / `qty` | 数值 |
| I 消费 | 由公式计算或填 `cost` | 默认写成 `=G*H` 公式；只有总价时，G 和 H 留空，直接填数 |
| J 开售 / 预约 | `booking_at` | Excel 日期或日期时间值 |
| K 状态 | `status` | 下拉选择，条件格式只给这一格着色；留空表示“无需预订” |
| L 地点 | `place` | 文本 |

冻结表头和摘要区，设置列宽，打印区域设为横向。类型和状态的颜色与网页一致。

## 3. 解析规则（xlsx → Trip）

- **区分行类型**：A 列为日期、且 D 至 L 列全部为空的行是分隔行；B 至 L 列全部为空的行是空行，跳过；其余都是事项行。
- **事项属于哪一天**：以它上方最近的分隔行为准，所以剪切粘贴到其他日期下面就能挪动事项。A 列日期和所属日期不一致时只给出警告，下次导出时会被覆盖。第一条分隔行之前出现的事项行直接报错。
- **金额**：G 或 H 有值时，按 `price` / `qty` 读取，忽略 I 列，因为腾讯文档可能把公式转成数值。G 和 H 都为空、I 有值时，按 `cost` 读取。
- **时间**：接受 Excel 时间值（四舍五入到分钟）和 `HH:MM` 文本。`booking_at` 的时间部分为 `00:00` 时按纯日期处理。
- **校验**：解析结果统一交给 `trip_model.Trip.model_validate`，报错信息换算成 **Excel 行号和列名**，例如“第 23 行 D 列 类型：无效值 '吃饭'”。只要有任何错误，就不写任何文件。

## 4. 同步算法

**规范哈希**：`digest(trip) = sha256(dump_trip(trip))`。基于语义计算，所以调整列宽、颜色或 YAML 缩进都不算修改。

每次导出 xlsx，都把当时的 digest 记为 `H`，写进 `_meta`。每次同步时：

- `Y` = digest(加载 YAML)，`X` = digest(解析 xlsx)

| 情况 | 动作 |
| --- | --- |
| X = Y，且 `_meta` 中的 H 等于这个值 | 已同步，不写文件 |
| X = Y，但 H 缺失或不同（例如经过了腾讯文档） | 重新导出 xlsx，恢复 `_meta` |
| X ≠ Y 且 Y = H（只改了 xlsx） | 用 `dump_trip` 写回 YAML，再重新导出 xlsx 刷新格式 |
| X ≠ Y 且 X = H（只改了 YAML，比如 AI 修改） | 重新导出 xlsx |
| X ≠ Y，且 H 与两边都不同或缺失 | **冲突，中止**。输出两边的差异摘要，等用户用 `--prefer` 指定以哪边为准 |
| 只有 xlsx | 当作新建旅行，生成 YAML（复制一份已有的 xlsx 再修改，就能开新行程） |
| 只有 YAML | 导出 xlsx |

- 只要有一边改过，就输出变更摘要：新增、删除、修改了多少条事项，总额和待办数的前后变化。
- 只在内容真的变化时才重写 xlsx，避免每次提交都带上一个无意义的二进制 diff。
- 写入时先写临时文件再替换原文件。目标文件被 Excel 占用时，提示“请先关闭 Excel：<文件>”并返回非 0，不会留下写了一半的文件。
- 任何一次旅行的同步出错都会让整条命令返回非 0，但已成功的旅行照常写入。

## 5. 提交前自动同步与 CI

- `.githooks/pre-commit` 执行 `python scripts/trip_xlsx.py sync --print-written`：
  - 成功后，把 sync 写入的文件执行 `git add`，加入本次提交，并在终端列出这些文件；
  - 遇到冲突、校验失败或文件被占用时中止提交。
- 每台 PC 需要各执行一次 `git config core.hooksPath .githooks`，只修改仓库本地配置。README 中补充这一步。
- `deploy.yml` 在 “Validate trips” 之后增加一步 `python scripts/trip_xlsx.py sync --check`。即使有人跳过 hook 直接提交，也会被 CI 拦下，不会发布出与 xlsx 不一致的网页。

## 6. 本机保存后自动同步（watch 模式）

- `python scripts/trip_xlsx.py sync --watch`：常驻进程，用 `watchdog` 监听 `trips/` 下的 `*.yaml` 和 `*.xlsx`。启动时先完整执行一次 `sync`。
- **防抖**：一次旅行的文件最后一次变化后静默 1.5 秒，才对这次旅行执行 sync。Excel 保存时会出现临时文件、重命名等多个事件，这样合并成一次。`~$*.xlsx` 和其他临时文件忽略。
- **忽略自身写入**：sync 写出的文件会再触发事件，按写入后的 digest 判定为“已同步”，不会循环触发。没有写文件的 sync 会记下两个文件的 `(mtime, size)`；防抖到期时，如果两个文件都没变，就跳过这次 sync。这样可以避免 Windows 上读文件也触发事件、导致反复解析 xlsx。
- **出错不退出**：遇到校验失败、冲突或文件被占用时，在终端打印错误，保持文件原样，继续监听。文件被占用时，每 5 秒重试一次，直到 Excel 关闭；重试期间同一条提示只打印一次。冲突仍然需要手动执行 `sync <id> --prefer ...`，处理完后 watch 会自动恢复。
- **成功提示**：每次写入后打印一行，包括时间、旅行 id、同步方向和变更摘要。
- **启动方式**：`.vscode/tasks.json` 提供 “Trip sync watch” 任务，设置 `runOptions.runOn: folderOpen`，打开仓库后自动在终端面板运行。第一次需要在 VS Code 中允许自动任务。也可以从命令面板手动启动。
- pre-commit 和 CI 的检查保留不变，作为 watch 没在运行时的兜底。
- 依赖：`requirements.txt` 新增 `watchdog`，固定版本号。CI 不运行 watch。

已知限制：Excel 打开 xlsx 时，watch 无法写入这个文件。如果这期间 YAML 被修改（例如由 AI 修改），xlsx 要等 Excel 关闭后才会刷新。如果在刷新前又在 Excel 中保存了修改，会被判定为冲突，需要手动选择以哪一边为准。

## 7. 测试

- **往返测试**：对仓库中的每个 trip 执行 export 后再 parse，结果的 digest 必须与原 YAML 相同。
- **同步状态机**：第 4 节表格中每一行至少一个用例，包括冲突时不写任何文件。
- **解析规则**：把事项挪到另一天、公式被转成数值、只填总价、文本格式的时间、空行、第一条分隔行之前出现事项。
- **报错定位**：非法类型、单价与总价同时填写、日期越界，断言报错中包含正确的 Excel 行号。
- **导出结构**：用 openpyxl 读取导出文件，断言数据验证（下拉）、I 列公式、条件格式和 `_meta` 都存在；摘要区公式的计算结果与 `compute_stats` 一致（用 LibreOffice 或手工核对一次即可，不放进 CI）。
- **watch 模式**：把防抖和调度逻辑与 `watchdog` 解耦后单独测试，包括多次事件合并为一次 sync、自身写入不再次触发、出错后继续处理后续事件、文件被占用时重试。真实的文件监听只做一次手工验证。
- **人工验证**：用 Excel 打开并截图检查版式；上传到腾讯文档编辑一处，下载后执行 sync，确认能识别并写回。

## 8. 文档更新

- `CLAUDE.md`：目录表加入 xlsx 和 `trip_xlsx.py`。修改流程改为“修改 YAML 或 xlsx → `sync`（watch 运行时自动执行）→ `--check` → 提交”。规定 AI 只修改 YAML，修改后必须执行 `sync`。
- `README.md`：说明首次在一台 PC 上配置 hook 和允许 VS Code 自动任务的方法，以及腾讯文档的往返流程。
