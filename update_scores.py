#!/usr/bin/env python3
"""Utility for syncing KovaaK stats.db high scores into the Viscose Excel tracker.

Phase 1 deliverable: a manually executed script that reads KovaaK's SQLite database,
extracts per-scenario high scores, and writes them into the "High Score" column
inside the provided workbook.
"""
from __future__ import annotations

import argparse
import csv
import io
import logging
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import pandas as pd  # Optional, used for pretty console summaries.
except ImportError:  # pragma: no cover - pandas is optional but recommended.
    pd = None  # type: ignore

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

LOGGER = logging.getLogger("kovaak_updater")
NAME_KEYWORDS = ["scenario", "drill", "map", "task", "playlist", "name", "title"]
SCORE_KEYWORDS = ["score", "result", "value", "points"]
NUMERIC_HINTS = ["int", "real", "num", "double", "float"]
DEFAULT_SCENARIO_SHEET_FRAGMENT = "scenario"
HOME = Path.home()
DEFAULT_STATS_DIRS = [
    Path("C:/Program Files (x86)/Steam/steamapps/common/KovaaKs FPS Aim Trainer/stats"),
    Path("C:/Program Files/Steam/steamapps/common/KovaaKs FPS Aim Trainer/stats"),
    HOME / ".steam/steam/steamapps/common/KovaaKs FPS Aim Trainer/stats",
    HOME / ".local/share/Steam/steamapps/common/KovaaKs FPS Aim Trainer/stats",
    HOME / "Steam/steamapps/common/KovaaKs FPS Aim Trainer/stats",
    HOME / ".steam/steam/steamapps/common/FPSAimTrainer/FPSAimTrainer/stats",
]


def try_parse_float(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    cleaned = str(value).replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


@dataclass
class TableSelection:
    table: str
    name_column: str
    score_column: str


@dataclass
class SheetHeader:
    header_row: int
    scenario_col: int
    high_score_col: int


@dataclass
class SheetUpdateStats:
    sheet: str
    updated: int
    missing: List[str]


def normalize_key(value: str) -> str:
    clean = " ".join(value.split())
    return clean.strip().lower()


QUOTED_TEXT_PATTERN = re.compile(r'"([^\"]+)"')


def extract_header_text(value) -> Optional[str]:
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    if text.startswith("="):
        match = re.match(r'^="([^"]+)"$', text)
        if match:
            text = match.group(1)
        else:
            quoted = QUOTED_TEXT_PATTERN.search(text)
            if quoted:
                text = quoted.group(1)
            else:
                text = text.lstrip("=")
    text = text.strip()
    return text.lower() if text else None


def quote_identifier(identifier: str) -> str:
    escaped = identifier.replace('"', '""')
    return f'"{escaped}"'


def list_existing_stats_dirs() -> List[Path]:
    return [path for path in DEFAULT_STATS_DIRS if path.exists()]


def detect_table(conn: sqlite3.Connection) -> Optional[TableSelection]:
    cursor = conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
    best_score = -1.0
    best_selection: Optional[TableSelection] = None

    for (table_name,) in cursor.fetchall():
        table_info = conn.execute(f"PRAGMA table_info({quote_identifier(table_name)})").fetchall()
        if not table_info:
            continue

        name_candidates = []
        score_candidates = []
        for column in table_info:
            col_name = column[1]
            col_type = (column[2] or '').lower()
            lower_name = col_name.lower()
            name_weight = max((len(kw) for kw in NAME_KEYWORDS if kw in lower_name), default=0)
            score_weight = max((len(kw) for kw in SCORE_KEYWORDS if kw in lower_name), default=0)
            numeric_hint = any(hint in col_type for hint in NUMERIC_HINTS)

            if name_weight:
                name_candidates.append((col_name, name_weight))
            if score_weight or numeric_hint:
                bonus = 2 if score_weight else 1
                score_candidates.append((col_name, score_weight + bonus))

        if not name_candidates or not score_candidates:
            continue

        table_weight = (max(n[1] for n in name_candidates) + max(s[1] for s in score_candidates))
        if table_weight > best_score:
            best_score = table_weight
            best_selection = TableSelection(
                table=table_name,
                name_column=max(name_candidates, key=lambda x: x[1])[0],
                score_column=max(score_candidates, key=lambda x: x[1])[0],
            )

    return best_selection


def fetch_scores_from_db(
    db_path: Path,
    table: Optional[str] = None,
    name_column: Optional[str] = None,
    score_column: Optional[str] = None,
) -> Tuple[Dict[str, float], TableSelection]:
    conn = sqlite3.connect(db_path)
    try:
        conn.row_factory = sqlite3.Row
        selection = TableSelection(table, name_column, score_column) if table and name_column and score_column else None
        if selection is None:
            selection = detect_table(conn)
        if selection is None:
            raise RuntimeError(
                "无法根据 stats.db 自动识别场景得分表。请使用 --table/--name-column/--score-column 参数指定。"
            )

        LOGGER.info(
            "使用数据表 %s (名称列=%s, 分数列=%s)",
            selection.table,
            selection.name_column,
            selection.score_column,
        )
        name_sql = quote_identifier(selection.name_column)
        score_sql = quote_identifier(selection.score_column)
        table_sql = quote_identifier(selection.table)
        query = f"""
            SELECT {name_sql} AS scenario_name,
                   MAX(CAST({score_sql} AS REAL)) AS high_score
            FROM {table_sql}
            WHERE {name_sql} IS NOT NULL AND TRIM({name_sql}) <> ''
              AND {score_sql} IS NOT NULL
            GROUP BY {name_sql}
        """
        result: Dict[str, float] = {}
        for row in conn.execute(query):
            name_val = row["scenario_name"]
            score_val = row["high_score"]
            if name_val is None or score_val is None:
                continue
            normalized = normalize_key(str(name_val))
            result[normalized] = float(score_val)

        if not result:
            raise RuntimeError("无法从 stats.db 获取任何分数，请确认数据库内包含游戏数据。")

        return result, selection
    finally:
        conn.close()


def detect_csv_columns(headers: Sequence[Optional[str]]) -> Tuple[Optional[str], Optional[str]]:
    scenario_col = None
    scenario_weight = 0
    score_col = None
    score_weight = 0
    for header in headers:
        if not header:
            continue
        text = header.strip().lower()
        n_weight = max((len(kw) for kw in NAME_KEYWORDS if kw in text), default=0)
        s_weight = max((len(kw) for kw in SCORE_KEYWORDS if kw in text), default=0)
        if n_weight > scenario_weight:
            scenario_col = header
            scenario_weight = n_weight
        if s_weight > score_weight:
            score_col = header
            score_weight = s_weight
    return scenario_col, score_col


def select_numeric_column(rows: List[Dict[str, str]], headers: Sequence[Optional[str]]) -> Optional[str]:
    for header in headers:
        if not header:
            continue
        numeric_values = 0
        for row in rows:
            value = row.get(header)
            parsed = try_parse_float(value) if value is not None else None
            if parsed is None:
                numeric_values = 0
                break
            numeric_values += 1
        if numeric_values:
            return header
    return None

def extract_key_value_score(rows: Sequence[Sequence[str]], fallback_name: str) -> Optional[Tuple[str, float]]:
    scenario_name: Optional[str] = None
    best_score: Optional[float] = None

    for row in rows:
        trimmed = [cell.strip() for cell in row if cell and cell.strip()]
        if not trimmed:
            continue
        idx = 0
        while idx < len(trimmed):
            cell = trimmed[idx]
            key = None
            value = None
            if ":" in cell:
                key_part, _, remainder = cell.partition(":")
                key = key_part.strip()
                value = remainder.strip()
                if not value and idx + 1 < len(trimmed):
                    value = trimmed[idx + 1]
                    idx += 1
            elif cell.endswith(":") and idx + 1 < len(trimmed):
                key = cell.rstrip(":").strip()
                value = trimmed[idx + 1]
                idx += 1
            else:
                idx += 1
                continue

            idx += 1
            if not key:
                continue

            key_lower = key.lower()
            normalized_value = value.strip() if value else ""
            if key_lower == "scenario" or any(kw in key_lower for kw in NAME_KEYWORDS):
                scenario_name = normalized_value or scenario_name
            if "score" in key_lower:
                parsed = try_parse_float(normalized_value)
                if parsed is not None and (best_score is None or parsed > best_score):
                    best_score = parsed

    if best_score is None:
        return None

    scenario = scenario_name or fallback_name
    return scenario, best_score


def fetch_scores_from_csv(
    csv_files: Sequence[Path],
    delimiter: str = ",",
) -> Dict[str, float]:
    if not csv_files:
        raise RuntimeError("未提供任何 CSV 文件。")

    aggregated: Dict[str, float] = {}
    for csv_path in csv_files:
        if not csv_path.exists():
            LOGGER.warning("CSV 文件不存在: %s", csv_path)
            continue
        raw_text = csv_path.read_text(encoding="utf-8-sig")
        dict_reader = csv.DictReader(io.StringIO(raw_text), delimiter=delimiter)
        rows = [row for row in dict_reader]
        headers = dict_reader.fieldnames or []
        plain_rows = list(csv.reader(io.StringIO(raw_text), delimiter=delimiter))

        fallback_name = csv_path.stem
        updated_rows = 0

        scenario_col, score_col = detect_csv_columns(headers)
        if score_col is None and rows:
            score_col = select_numeric_column(rows, headers)

        if score_col is not None and rows:
            for row in rows:
                scenario_value = row.get(scenario_col) if scenario_col else fallback_name
                if scenario_value is None:
                    scenario_value = fallback_name
                scenario_name = str(scenario_value).strip() or fallback_name

                raw_score = row.get(score_col)
                score_value = try_parse_float(raw_score) if raw_score is not None else None
                if score_value is None:
                    continue

                key = normalize_key(scenario_name)
                previous = aggregated.get(key)
                if previous is None or score_value > previous:
                    aggregated[key] = score_value
                    updated_rows += 1
            if updated_rows == 0:
                kv_result = extract_key_value_score(plain_rows, fallback_name)
                if kv_result:
                    scenario_name, score_value = kv_result
                    key = normalize_key(scenario_name)
                    previous = aggregated.get(key)
                    if previous is None or score_value > previous:
                        aggregated[key] = score_value
                        updated_rows += 1
        else:
            kv_result = extract_key_value_score(plain_rows, fallback_name)
            if kv_result:
                scenario_name, score_value = kv_result
                key = normalize_key(scenario_name)
                previous = aggregated.get(key)
                if previous is None or score_value > previous:
                    aggregated[key] = score_value
                    updated_rows += 1

        if updated_rows:
            LOGGER.debug("解析 CSV %s -> 记录 %d 条", csv_path.name, updated_rows)
        else:
            LOGGER.warning("CSV %s 未找到有效的场景/分数记录，已跳过。", csv_path)

    if not aggregated:
        raise RuntimeError("CSV 文件中未找到可用的场景分数。")

    return aggregated


def detect_header(ws: Worksheet) -> Optional[SheetHeader]:
    max_row = min(ws.max_row, 30)
    scenario_col = None
    high_score_col = None
    header_row = None

    for row in range(1, max_row + 1):
        for cell in ws[row]:
            value = cell.value
            text = extract_header_text(value)
            if text is None:
                continue
            if text == "scenario":
                scenario_col = cell.column
                header_row = cell.row
            elif text == "high score":
                high_score_col = cell.column
                header_row = header_row or cell.row
        if scenario_col and high_score_col:
            return SheetHeader(header_row=header_row, scenario_col=scenario_col, high_score_col=high_score_col)

    return None


def update_sheet(ws: Worksheet, scores: Dict[str, float]) -> SheetUpdateStats:
    header = detect_header(ws)
    if header is None:
        LOGGER.warning("工作表 %s 未找到 'Scenario'/'High Score' 标题，已跳过。", ws.title)
        return SheetUpdateStats(sheet=ws.title, updated=0, missing=[])

    updated = 0
    missing: List[str] = []
    for row_idx in range(header.header_row + 1, ws.max_row + 1):
        scenario_cell = ws.cell(row=row_idx, column=header.scenario_col)
        value = scenario_cell.value
        if value is None:
            continue
        scenario_name = str(value).strip()
        if not scenario_name:
            continue
        normalized = normalize_key(scenario_name)
        if normalized not in scores:
            missing.append(scenario_name)
            continue
        score_value = scores[normalized]
        target_cell = ws.cell(row=row_idx, column=header.high_score_col)
        target_cell.value = round(score_value, 2)
        updated += 1

    return SheetUpdateStats(sheet=ws.title, updated=updated, missing=missing)


def select_default_sheets(sheet_names: Sequence[str]) -> List[str]:
    candidates = [name for name in sheet_names if DEFAULT_SCENARIO_SHEET_FRAGMENT in name.lower()]
    return candidates or list(sheet_names)


def collect_csv_files(entries: Sequence[Path], pattern: str = "*.csv") -> List[Path]:
    files: List[Path] = []
    for entry in entries:
        resolved = entry.expanduser().resolve()
        if resolved.is_file() and resolved.suffix.lower() == ".csv":
            files.append(resolved)
        elif resolved.is_dir():
            files.extend(sorted(resolved.glob(pattern)))
        else:
            LOGGER.warning("未找到 CSV 路径: %s", resolved)
    return files


def guess_default_csv_files(pattern: str = "*.csv") -> List[Path]:
    files: List[Path] = []
    for directory in DEFAULT_STATS_DIRS:
        stats_dir = directory.expanduser()
        if stats_dir.is_dir():
            files.extend(sorted(stats_dir.glob(pattern)))
    return files


def update_workbook(
    workbook_path: Path,
    scores: Dict[str, float],
    sheets: Optional[Sequence[str]] = None,
) -> Tuple[List[SheetUpdateStats], Workbook]:
    wb = load_workbook(workbook_path)
    target_sheets = sheets or select_default_sheets(wb.sheetnames)
    stats: List[SheetUpdateStats] = []
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            LOGGER.warning("工作表 %s 不存在，已跳过。", sheet_name)
            continue
        ws = wb[sheet_name]
        stats.append(update_sheet(ws, scores))

    return stats, wb


def guess_default_db_path() -> Optional[Path]:
    for stats_dir in DEFAULT_STATS_DIRS:
        candidate = stats_dir / "stats.db"
        if candidate.exists():
            return candidate
    return None


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="同步 KovaaK 高分到 Excel")
    parser.add_argument("--db", type=Path, default=None, help="stats.db 路径 (默认自动探测常见路径)")
    parser.add_argument("--excel", type=Path, default=Path("Viscose Benchmarks Beta.xlsx"), help="Excel 文件路径")
    parser.add_argument("--output", type=Path, default=None, help="保存为新文件的路径 (默认原地覆盖)")
    parser.add_argument("--sheets", nargs="*", help="需要更新的工作表名称 (默认自动选择包含 Scenarios 的工作表)")
    parser.add_argument("--csv", nargs="*", type=Path, help="CSV 文件或目录，用于在缺少 stats.db 时提供分数数据")
    parser.add_argument("--csv-pattern", default="*.csv", help="当 --csv 指向目录时使用的 glob 模式")
    parser.add_argument("--csv-delimiter", default=",", help="CSV 文件使用的分隔符 (默认逗号)")
    parser.add_argument("--table", help="数据库中包含分数的表名")
    parser.add_argument("--name-column", dest="name_column", help="场景名称列")
    parser.add_argument("--score-column", dest="score_column", help="分数列")
    parser.add_argument("--dry-run", action="store_true", help="仅输出日志，不写回 Excel")
    parser.add_argument("--verbose", action="store_true", help="输出调试日志")
    return parser


def render_summary(stats: Iterable[SheetUpdateStats]):
    rows = [
        {
            "Sheet": item.sheet,
            "Updated": item.updated,
        }
        for item in stats
        if item.updated > 0
    ]

    if not rows:
        print("没有任何得分被写入 (0 条更新)。")
        return

    if pd is not None:
        df = pd.DataFrame(rows)
        print(df.to_string(index=False))
    else:
        for row in rows:
            print(f"工作表 {row['Sheet']} -> 更新 {row['Updated']} 条")


def main():
    parser = build_parser()
    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="[%(levelname)s] %(message)s")

    db_path: Optional[Path] = None
    if args.db:
        candidate = args.db.expanduser().resolve()
        if not candidate.exists():
            parser.error(f"数据库路径不存在: {candidate}")
        db_path = candidate
    else:
        guessed_db = guess_default_db_path()
        if guessed_db:
            db_path = guessed_db.expanduser().resolve()
            LOGGER.info("自动检测到 stats.db: %s", db_path)

    csv_files: List[Path] = []
    csv_requested = bool(args.csv)
    if args.csv:
        csv_files = collect_csv_files(args.csv, pattern=args.csv_pattern)
        if not csv_files:
            parser.error("在 --csv 指定的路径中未找到任何 CSV 文件")

    if db_path is None and not csv_files:
        guessed_csv = guess_default_csv_files(pattern=args.csv_pattern)
        if guessed_csv:
            csv_files = guessed_csv
            LOGGER.info("未找到 stats.db，改用 %d 个 CSV 文件", len(csv_files))

    data_source = None
    if csv_requested and csv_files:
        data_source = "csv"
    elif db_path is not None:
        data_source = "db"
    elif csv_files:
        data_source = "csv"
    else:
        parser.error("未能找到 stats.db，也没有可用的 CSV 文件。请使用 --db 或 --csv 指定数据来源。")

    excel_path = args.excel.expanduser().resolve()
    if not excel_path.exists():
        parser.error(f"未找到 Excel 文件: {excel_path}")

    if data_source == "db":
        assert db_path is not None
        scores, selection = fetch_scores_from_db(
            db_path=db_path,
            table=args.table,
            name_column=args.name_column,
            score_column=args.score_column,
        )
        LOGGER.info("共加载 %d 条场景分数记录 (来自 SQLite)", len(scores))
    else:
        scores = fetch_scores_from_csv(csv_files, delimiter=args.csv_delimiter)
        selection = None
        LOGGER.info("共加载 %d 条场景分数记录 (来自 CSV)", len(scores))

    workbook_stats, workbook = update_workbook(
        workbook_path=excel_path,
        scores=scores,
        sheets=args.sheets,
    )

    render_summary(workbook_stats)

    if args.dry_run:
        LOGGER.info("已启用 --dry-run，未对 Excel 进行写入。")
        return

    output_path = args.output.expanduser().resolve() if args.output else excel_path
    if output_path == excel_path:
        LOGGER.info("保存修改 -> %s", output_path)
    else:
        LOGGER.info("保存到新文件 -> %s", output_path)
    workbook.save(output_path)


if __name__ == "__main__":
    main()
