# Aim Updater – 阶段 1 (MVP)

该项目实现了一个跨平台的 Python 脚本 `update_scores.py`，用于从 KovaaK 的 `stats.db` 数据库读取每个场景的历史最高分，并把分数写入到 `Viscose Benchmarks Beta.xlsx` 中的 "High Score" 列。

## 环境准备

1. 建议使用 Python 3.11 及以上版本（脚本已在 Linux 上测试，Windows 同样适用）。
2. 安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

> `sqlite3` 是 Python 标准库（不需要额外安装），`openpyxl` 用于读写 Excel，`pandas` 用于在终端展示更新摘要（没有也可以运行脚本）。

## KovaaK 数据源路径

- Windows 默认路径：`C:\Program Files (x86)\Steam\steamapps\common\KovaaKs FPS Aim Trainer\stats\stats.db`
- Linux/Steam Deck 常见路径：`~/.steam/steam/steamapps/common/KovaaKs FPS Aim Trainer/stats/stats.db`
- Linux Proton (FPSAimTrainer) 路径：`~/.steam/steam/steamapps/common/FPSAimTrainer/FPSAimTrainer/stats`

如果你的安装目录中不存在 `stats.db`（例如 Linux Proton 环境只保存 CSV），脚本会自动在同一 `stats` 目录下搜索 `*.csv` 文件并改用 CSV 数据源。也可以手动指定：

```bash
python3 update_scores.py --csv "/path/to/stats/"
```

脚本会尝试自动检测这些常见路径。若自动探测失败，请通过 `--db` 手动指定。

## 使用方法

1. **确保 Excel 文件在仓库根目录**（或使用 `--excel` 指定其它路径）。
2. 运行脚本：

```bash
python3 update_scores.py --db "/path/to/stats.db" --excel "Viscose Benchmarks Beta.xlsx"
```

若使用 CSV 数据源（单个文件或目录均可）：

```bash
python3 update_scores.py --csv "/path/to/stats/*.csv" --excel "Viscose Benchmarks Beta.xlsx"
```

脚本会：

- 连接 SQLite 数据库，自动检测合适的数据表/列，按场景聚合最高分。
- 遍历包含 "Scenarios" 的工作表（Easier/Medium/Hard），匹配 `Scenario` 列中的名称，与数据库中的场景名称进行大小写不敏感匹配。
- 将匹配成功的分数写入 "High Score" 列。
- 在终端输出每个工作表的更新统计以及无法匹配的场景。

### 常用参数

| 参数 | 说明 |
| ---- | ---- |
| `--db` | stats.db 路径。若省略则尝试自动探测常见路径。 |
| `--csv` | 一个或多个 CSV 文件/目录路径（支持通配符所在目录），在缺少 `stats.db` 时作为数据源。 |
| `--csv-pattern` | 当 `--csv` 指向目录时用于匹配文件的 glob，默认 `*.csv`。 |
| `--csv-delimiter` | CSV 的分隔符，默认逗号。 |
| `--excel` | Excel 文件路径，默认 `Viscose Benchmarks Beta.xlsx`。 |
| `--output` | 将结果写入新的文件而不是覆盖原文件。 |
| `--sheets` | 指定需要更新的工作表名称，默认匹配名称里包含 `Scenarios` 的工作表。 |
| `--table` / `--name-column` / `--score-column` | 如果自动检测数据库结构失败，可用此三项明确指定 SQL 表与列。 |
| `--dry-run` | 只查看匹配结果，不写入 Excel。 |
| `--verbose` | 输出调试日志，方便定位问题。 |

### 示例

```bash
python3 update_scores.py \
  --db "C:/Program Files (x86)/Steam/steamapps/common/KovaaKs FPS Aim Trainer/stats/stats.db" \
  --excel "Viscose Benchmarks Beta.xlsx" \
  --output "Viscose Benchmarks Updated.xlsx"
```

### 运行机制概述

- **自动 SQL 检测**：脚本会扫描数据库的所有表，优先选择同时包含场景名称与分数字段的表。必要时可以用 `--table` 等参数强制指定。
- **场景匹配**：将数据库与 Excel 中的场景名称进行大小写与空白符无关的匹配，例如 "Air Angelic"、"air  angelic" 都视为同一键。
- **写回 Excel**：只会覆盖 "High Score" 列，其他列（公式、格式、排名阈值等）保持不变。

## 已完成的阶段 1 任务

- [x] 安装并记录所需依赖（sqlite3、openpyxl、pandas）。
- [x] 连接 KovaaK `stats.db` 并查询每个场景的历史最高分。
- [x] 读取 `Viscose Benchmarks Beta.xlsx` 的场景工作表。
- [x] 将场景名称与数据库结果匹配，并写入对应的 "High Score" 单元格。
- [x] 保存更新后的 Excel 文件（可选输出路径）。

接下来（阶段 2/3）将基于当前脚本继续封装配置、加入文件监控、以及启动器等功能。