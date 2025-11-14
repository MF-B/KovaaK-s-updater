#!/usr/bin/env python3
"""Background watcher that keeps KovaaK scores synced whenever stats files change."""
from __future__ import annotations

import argparse
import logging
import threading
import time
from pathlib import Path
from typing import Iterable, Optional, Set

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

from update_scores import (
	UpdateSettings,
	WatchSettings,
	discover_config_path,
	load_settings_from_config,
	prepare_watch_settings,
	render_summary,
	run_update,
)

LOGGER = logging.getLogger("kovaak_watcher")
UPDATE_LOCK = threading.Lock()


def _setup_logging(level: str, log_file: Optional[Path]) -> None:
	log_level = getattr(logging, level.upper(), logging.INFO)
	logging_kwargs = {
		"level": log_level,
		"format": "%(asctime)s [%(levelname)s] %(message)s",
	}
	if log_file:
		log_file = log_file.expanduser().resolve()
		logging_kwargs["filename"] = str(log_file)
	logging.basicConfig(**logging_kwargs)


class DebouncedRunner:
	def __init__(self, callback, debounce_seconds: float) -> None:
		self.callback = callback
		self.debounce_seconds = max(0.0, debounce_seconds)
		self._timer: Optional[threading.Timer] = None
		self._lock = threading.Lock()

	def trigger(self) -> None:
		if self.debounce_seconds <= 0:
			self.callback()
			return

		with self._lock:
			if self._timer:
				self._timer.cancel()
			self._timer = threading.Timer(self.debounce_seconds, self._fire)
			self._timer.daemon = True
			self._timer.start()

	def _fire(self) -> None:
		try:
			self.callback()
		finally:
			with self._lock:
				self._timer = None


def _path_is_within(path: Path, parents: Set[Path]) -> bool:
	for parent in parents:
		try:
			path.relative_to(parent)
			return True
		except ValueError:
			continue
	return False


class StatsEventHandler(FileSystemEventHandler):
	def __init__(
		self,
		runner: DebouncedRunner,
		file_targets: Set[Path],
		directory_targets: Set[Path],
	) -> None:
		super().__init__()
		self.runner = runner
		self.file_targets = {path.resolve() for path in file_targets}
		self.directory_targets = {path.resolve() for path in directory_targets}

	def on_created(self, event: FileSystemEvent) -> None:  # type: ignore[override]
		self._handle_event(event)

	def on_modified(self, event: FileSystemEvent) -> None:  # type: ignore[override]
		self._handle_event(event)

	def _handle_event(self, event: FileSystemEvent) -> None:
		if event.is_directory:
			return
		path = Path(event.src_path).resolve()
		if self._is_relevant(path):
			LOGGER.debug("检测到文件变更: %s", path)
			self.runner.trigger()

	def _is_relevant(self, path: Path) -> bool:
		if not self.file_targets and not self.directory_targets:
			return True
		if path in self.file_targets:
			return True
		return _path_is_within(path, self.directory_targets)


def _determine_schedule_paths(paths: Iterable[Path]) -> tuple[Set[Path], Set[Path], Set[Path]]:
	files: Set[Path] = set()
	directories: Set[Path] = set()
	schedule_dirs: Set[Path] = set()

	def find_existing_dir(path: Path) -> Path:
		candidate = path
		while not candidate.exists():
			if candidate.parent == candidate:
				break
			candidate = candidate.parent
		return candidate if candidate.exists() else candidate.parent

	for raw in paths:
		resolved = raw.expanduser().resolve()
		if resolved.is_dir():
			directories.add(resolved)
			schedule_dirs.add(resolved)
		else:
			files.add(resolved)
			parent = resolved.parent if resolved.parent != resolved else resolved
			existing = find_existing_dir(parent) if not parent.exists() else parent
			if existing.exists():
				schedule_dirs.add(existing)

	if not schedule_dirs:
		schedule_dirs.add(Path.cwd())

	return files, directories, schedule_dirs


def _perform_update(settings: UpdateSettings) -> None:
	with UPDATE_LOCK:
		try:
			stats, output_path = run_update(settings)
			total = sum(item.updated for item in stats)
			if total:
				LOGGER.info("已写入 %d 条记录 -> %s", total, output_path)
			else:
				LOGGER.info("本次触发未产生新的得分。")
			render_summary(stats)
		except Exception:  # pragma: no cover - watcher should keep running
			LOGGER.exception("自动更新失败")


def main() -> None:
	parser = argparse.ArgumentParser(description="监视 KovaaK stats 文件并自动更新 Excel")
	parser.add_argument("--config", type=Path, default=None, help="配置文件路径 (默认当前目录 config.json)")
	parser.add_argument("--debounce", type=float, default=None, help="覆盖配置中的防抖秒数")
	parser.add_argument("--skip-initial", action="store_true", help="启动时跳过第一次同步")
	parser.add_argument("--log-level", default="INFO", help="日志级别 (INFO/DEBUG/...)")
	parser.add_argument("--log-file", type=Path, default=None, help="可选日志文件 (推荐配合 pythonw 使用)")
	args = parser.parse_args()

	_setup_logging(args.log_level, args.log_file)

	config_path = discover_config_path(args.config)
	if not config_path:
		parser.error("未提供 --config，且当前目录没有 config.json")

	update_settings, watch_settings = load_settings_from_config(config_path)
	if update_settings is None:
		parser.error("配置文件中缺少必要的 Excel/路径设置")

	watch_settings = prepare_watch_settings(update_settings, watch_settings)
	if args.debounce is not None:
		watch_settings.debounce_seconds = max(0.0, args.debounce)
	if args.skip_initial:
		watch_settings.run_on_start = False

	if not watch_settings.paths:
		parser.error("无法确定需要监视的路径，请在 config.json 的 watch.paths 中指定")

	LOGGER.info("监视目录/文件: %s", ", ".join(str(p) for p in watch_settings.paths))
	runner = DebouncedRunner(lambda: _perform_update(update_settings), watch_settings.debounce_seconds)

	file_targets, dir_targets, schedule_dirs = _determine_schedule_paths(watch_settings.paths)
	handler = StatsEventHandler(runner, file_targets, dir_targets)

	observer = Observer()
	for directory in sorted(schedule_dirs):
		observer.schedule(handler, str(directory), recursive=True)

	observer.start()
	LOGGER.info("监视器已启动 (防抖 %.1f 秒)", watch_settings.debounce_seconds)

	if watch_settings.run_on_start:
		runner.trigger()

	try:
		while True:
			time.sleep(1)
	except KeyboardInterrupt:
		LOGGER.info("收到退出指令，正在停止监视器...")
	finally:
		observer.stop()
		observer.join()


if __name__ == "__main__":
	main()
