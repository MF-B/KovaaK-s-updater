# Aim Updater – 阶段 2：配置化 + 监控

该项目提供两个工具：

1. `update_scores.py` – 读取 KovaaK 的 `stats.db` 或 CSV 导出并把最高分同步到 `Viscose Benchmarks Beta.xlsx`。
2. `watcher.pyw` – 监听 KovaaK 目录，一旦有新的成绩文件就自动调用 `update_scores` 完成同步。

两者都围绕统一的 `config.json` 运行，可通过命令行参数进行临时覆盖。

## 环境准备

- Python 3.11 或以上（在 Linux/Steam Deck + Windows 下测试）。
- 安装依赖：

  ```bash
  python3 -m pip install -r requirements.txt
  ```

  - `openpyxl` 负责读写 Excel
  - `pandas` 用于漂亮的终端摘要（可选）
  - `watchdog` 为 watcher 提供跨平台文件系统事件

## 配置文件（config.json）

1. 复制示例：`cp config.example.json config.json`
2. 根据实际环境修改路径。相对路径以 `config.json` 所在目录为基准。

关键字段：

| 字段 | 说明 |
| --- | --- |
| `excel` | 目标 Excel 文件。默认 `Viscose Benchmarks Beta.xlsx` 可放在仓库根目录。|
| `stats_db` | KovaaK 的 `stats.db` 路径。缺失时会尝试 CSV。|
| `csv_paths` / `csv_pattern` / `csv_delimiter` | 当没有数据库或希望改用 CSV 时，列出 CSV 文件或目录。|
| `sheets` | 需要更新的工作表名称（可选，默认匹配包含 `Scenarios` 的表）。|
| `table` / `name_column` / `score_column` | 如自动探测数据库失败，可在此强制指定。|
| `output` | 写入新的 Excel 文件而非覆盖原文件（可选）。|
| `watch` | Watcher 配置，包含 `paths`（监控目录/文件）、`debounce_seconds`（防抖）、`run_on_start`。|

> `config.json` 不是必需的：没有配置文件时脚本会回退到旧的命令行行为，并尝试自动检测常见的 KovaaK 安装目录。

## 手动同步（CLI）

```bash
python3 update_scores.py --config config.json
```

- 如果没有 `--config`，脚本会在当前目录查找 `config.json`，找不到则使用内置默认值。
- 任意命令行参数都会覆盖配置文件中的值，例如：

  ```bash
  python3 update_scores.py --config config.json --excel "~/Documents/Viscose.xlsx" --dry-run
  ```

### CLI 参数速查

| 参数 | 作用 |
| ---- | ---- |
| `--config` | 指定配置文件（默认寻找 `./config.json`）。|
| `--excel` / `--db` / `--csv` 等 | 与配置字段同名，可即时覆盖。|
| `--output` | 输出到新的 Excel。|
| `--dry-run` | 不落盘，仅输出日志。|
| `--verbose` | 输出调试日志。|

脚本会自动：

1. 选取数据库或 CSV 作为数据源（优先数据库）。
2. 为每个场景挑选最高分，匹配 Excel 中的 `Scenario` 列。
3. 更新 `High Score` 列，并在终端打印每个工作表的更新统计。

## 自动同步（watcher）

Watcher 会监视 `watch.paths` 中的目录/文件，只要有修改就调用一次 `update_scores`。

```bash
python3 watcher.pyw --config config.json
```

可选参数：

| 参数 | 作用 |
| --- | --- |
| `--debounce` | 临时覆盖配置中的防抖秒数。|
| `--skip-initial` | 启动时不立刻跑第一次同步。|
| `--log-level` / `--log-file` | 控制 watcher 的日志输出。|

> Windows 下可以用 `pythonw watcher.pyw --config config.json` 在后台运行，并配合 `--log-file` 保存日志。

## 常见路径参考

- Windows：`C:/Program Files (x86)/Steam/steamapps/common/KovaaKs FPS Aim Trainer/stats/stats.db`
- Linux/Steam Deck：`~/.steam/steam/steamapps/common/KovaaKs FPS Aim Trainer/stats/stats.db`
- Proton/FPSAimTrainer：`~/.steam/steam/steamapps/common/FPSAimTrainer/FPSAimTrainer/stats`

如果既找不到数据库也没有 CSV，脚本会提示需要手动指定路径。

## 当前进度

- [x] 阶段 1：手动脚本（MVP）
- [x] 阶段 2：JSON 配置、命令行覆盖、watcher 自动同步
- [ ] 阶段 3：更友好的 UI/打包（后续计划）

欢迎根据自己的习惯扩展配置，或在 `tests/` 中添加新的单元测试来覆盖更多场景。