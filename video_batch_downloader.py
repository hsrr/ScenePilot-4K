#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import logging
import math
import random
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from shutil import which
from typing import Any, Iterable

import pandas as pd


COLUMN_ALIASES = {
    "folder": [
        "folder",
        "folder_name",
        "category",
        "一级目录",
        "一级文件夹",
        "文件夹",
        "文件夹名称",
        "目录",
        "分类",
    ],
    "slice": [
        "slice",
        "slice_folder",
        "slice_name",
        "subfolder",
        "二级目录",
        "二级文件夹",
        "切片",
        "切片文件夹",
        "切片文件夹名称",
        "切片名称",
        "子文件夹",
    ],
    "url": [
        "url",
        "link",
        "video_url",
        "video_link",
        "播放链接",
        "视频链接",
        "网络链接",
        "链接",
        "网址",
        "视频网址",
    ],
    "name": [
        "name",
        "video_name",
        "filename",
        "index",
        "order",
        "编号",
        "序号",
        "视频名称",
        "文件名",
    ],
}

TEMP_FILE_SUFFIXES = {".part", ".ytdl", ".temp"}
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)


@dataclass
class VideoTask:
    row_number: int
    folder_name: str
    slice_name: str
    url: str
    base_name: str
    output_dir: Path

    @property
    def output_stem(self) -> Path:
        return self.output_dir / self.base_name


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read an Excel/CSV manifest and download videos with yt-dlp.",
    )
    parser.add_argument("--input", required=True, help="Path to the Excel or CSV file.")
    parser.add_argument(
        "--output-root",
        default="downloads",
        help="Root output directory. Default: %(default)s",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Excel sheet name or zero-based sheet index. Defaults to the first sheet.",
    )
    parser.add_argument(
        "--folder-column",
        default=None,
        help="Override the column name for the first-level folder.",
    )
    parser.add_argument(
        "--slice-column",
        default=None,
        help="Override the column name for the slice/sub-folder.",
    )
    parser.add_argument(
        "--url-column",
        default=None,
        help="Override the column name for the video URL.",
    )
    parser.add_argument(
        "--name-column",
        default=None,
        help="Optional override for the sequence/name column.",
    )
    parser.add_argument(
        "--plan-only",
        action="store_true",
        help="Only parse the sheet and print the planned output paths.",
    )
    parser.add_argument(
        "--allow-playlist",
        action="store_true",
        help="Allow playlist-style URLs. Disabled by default to avoid accidental batch downloads.",
    )
    parser.add_argument(
        "--format",
        default="bv*+ba/best",
        help="yt-dlp format selector. Default: %(default)s",
    )
    parser.add_argument(
        "--merge-output-format",
        default="mp4",
        help="Merged output format when ffmpeg is available. Default: %(default)s",
    )
    parser.add_argument(
        "--cookies-file",
        default=None,
        help="Path to an exported Netscape cookies.txt file.",
    )
    parser.add_argument(
        "--cookies-browser",
        default=None,
        help="Browser name for --cookies-from-browser, e.g. chrome/firefox/edge.",
    )
    parser.add_argument(
        "--cookies-profile",
        default=None,
        help="Optional browser profile for --cookies-from-browser.",
    )
    parser.add_argument(
        "--proxy",
        default=None,
        help="Optional HTTP/SOCKS proxy passed to yt-dlp, e.g. socks5://127.0.0.1:7890",
    )
    parser.add_argument(
        "--user-agent",
        default=DEFAULT_USER_AGENT,
        help="User-Agent header used by yt-dlp. Default is a recent Chrome UA.",
    )
    parser.add_argument(
        "--sleep-interval",
        type=float,
        default=2.0,
        help="Base wait before each download, in seconds. Default: %(default)s",
    )
    parser.add_argument(
        "--max-sleep-interval",
        type=float,
        default=6.0,
        help="Random upper bound for yt-dlp sleep interval. Default: %(default)s",
    )
    parser.add_argument(
        "--sleep-requests",
        type=float,
        default=1.0,
        help="Wait between individual HTTP requests, in seconds. Default: %(default)s",
    )
    parser.add_argument(
        "--task-retries",
        type=int,
        default=2,
        help="How many times to retry a failed row after yt-dlp itself exhausts retries.",
    )
    parser.add_argument(
        "--retry-backoff",
        type=float,
        default=8.0,
        help="Initial task retry backoff in seconds. Default: %(default)s",
    )
    parser.add_argument(
        "--retry-jitter",
        type=float,
        default=2.0,
        help="Additional random jitter added to task retry backoff.",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=10,
        help="yt-dlp extractor/network retries per task. Default: %(default)s",
    )
    parser.add_argument(
        "--fragment-retries",
        type=int,
        default=10,
        help="yt-dlp fragment retries per task. Default: %(default)s",
    )
    parser.add_argument(
        "--socket-timeout",
        type=float,
        default=30.0,
        help="Socket timeout in seconds. Default: %(default)s",
    )
    parser.add_argument(
        "--limit-rate",
        default=None,
        help="Optional limit rate passed to yt-dlp, e.g. 2M or 800K.",
    )
    parser.add_argument(
        "--report-file",
        default=None,
        help="Optional CSV report path. Defaults to <output-root>/download_report.csv",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print the full yt-dlp command output.",
    )
    return parser.parse_args()


def configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def parse_sheet(sheet: str | None) -> int | str | None:
    if sheet is None:
        return None
    return int(sheet) if sheet.isdigit() else sheet


def normalize_column_name(name: Any) -> str:
    value = str(name).strip().lower()
    return re.sub(r"[\s_\-./()]+", "", value)


def resolve_column(
    frame: pd.DataFrame,
    override: str | None,
    logical_name: str,
    required: bool,
) -> str | None:
    if override:
        if override not in frame.columns:
            raise ValueError(
                f"Column '{override}' was not found. Available columns: {list(frame.columns)}"
            )
        return override

    normalized_to_original = {
        normalize_column_name(column): column for column in frame.columns
    }
    for alias in COLUMN_ALIASES[logical_name]:
        original = normalized_to_original.get(normalize_column_name(alias))
        if original is not None:
            return original

    alias_tokens = [normalize_column_name(alias) for alias in COLUMN_ALIASES[logical_name]]
    partial_matches: list[str] = []
    for normalized_name, original_name in normalized_to_original.items():
        if any(
            alias_token and (
                alias_token in normalized_name or normalized_name in alias_token
            )
            for alias_token in alias_tokens
        ):
            partial_matches.append(original_name)

    if len(partial_matches) == 1:
        return partial_matches[0]

    if required:
        raise ValueError(
            f"Could not detect the {logical_name} column automatically. "
            f"Available columns: {list(frame.columns)}. "
            f"You can also pass --{logical_name}-column explicitly."
        )
    return None


def sanitize_path_component(value: Any, fallback: str) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        text = ""
    else:
        text = str(value).strip()
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or fallback


def value_is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def format_base_name(raw_value: Any, default_index: int, width: int) -> str:
    if value_is_missing(raw_value):
        return f"{default_index:0{width}d}"

    if isinstance(raw_value, int):
        return f"{raw_value:0{width}d}"
    if isinstance(raw_value, float) and raw_value.is_integer():
        return f"{int(raw_value):0{width}d}"

    text = str(raw_value).strip()
    digits_only = re.fullmatch(r"\d+(?:\.0+)?", text)
    if digits_only:
        return f"{int(float(text)):0{width}d}"

    cleaned = sanitize_path_component(text, fallback=f"{default_index:0{width}d}")
    return cleaned


def deduplicate_name(
    used_names: dict[tuple[str, str], set[str]],
    group_key: tuple[str, str],
    base_name: str,
) -> str:
    existing_names = used_names.setdefault(group_key, set())
    if base_name not in existing_names:
        existing_names.add(base_name)
        return base_name

    suffix = 2
    candidate = f"{base_name}_{suffix:02d}"
    while candidate in existing_names:
        suffix += 1
        candidate = f"{base_name}_{suffix:02d}"
    existing_names.add(candidate)
    return candidate


def load_manifest(path: Path, sheet: int | str | None) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, sheet_name=0 if sheet is None else sheet)

    raise ValueError("Only .xlsx/.xlsm/.xls/.csv files are supported.")


def build_tasks(frame: pd.DataFrame, args: argparse.Namespace, output_root: Path) -> list[VideoTask]:
    folder_column = resolve_column(frame, args.folder_column, "folder", required=True)
    slice_column = resolve_column(frame, args.slice_column, "slice", required=True)
    url_column = resolve_column(frame, args.url_column, "url", required=True)
    name_column = resolve_column(frame, args.name_column, "name", required=False)

    working_frame = frame.copy()
    working_frame["_folder_name"] = working_frame[folder_column].map(
        lambda value: sanitize_path_component(value, fallback="未分类")
    )
    working_frame["_slice_name"] = working_frame[slice_column].map(
        lambda value: sanitize_path_component(value, fallback="默认切片")
    )

    group_totals = (
        working_frame.groupby(["_folder_name", "_slice_name"], dropna=False)
        .size()
        .to_dict()
    )
    group_counters: dict[tuple[str, str], int] = {}
    used_names: dict[tuple[str, str], set[str]] = {}
    tasks: list[VideoTask] = []

    for row_index, row in working_frame.iterrows():
        url = "" if value_is_missing(row[url_column]) else str(row[url_column]).strip()
        if not url:
            logging.warning("Skip empty URL at Excel row %s", row_index + 2)
            continue

        folder_name = row["_folder_name"]
        slice_name = row["_slice_name"]
        group_key = (folder_name, slice_name)
        group_counters[group_key] = group_counters.get(group_key, 0) + 1
        width = max(2, len(str(group_totals[group_key])))
        raw_name = row[name_column] if name_column else None
        base_name = format_base_name(
            raw_value=raw_name,
            default_index=group_counters[group_key],
            width=width,
        )
        base_name = deduplicate_name(used_names, group_key, base_name)
        tasks.append(
            VideoTask(
                row_number=row_index + 2,
                folder_name=folder_name,
                slice_name=slice_name,
                url=url,
                base_name=base_name,
                output_dir=output_root / folder_name / slice_name,
            )
        )

    return tasks


def find_existing_download(output_stem: Path) -> Path | None:
    for candidate in sorted(output_stem.parent.glob(f"{output_stem.name}.*")):
        if candidate.suffix.lower() in TEMP_FILE_SUFFIXES:
            continue
        if candidate.is_file():
            return candidate
    return None


def build_yt_dlp_command(task: VideoTask, args: argparse.Namespace) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "yt_dlp",
        "--newline",
        "--continue",
        "--no-overwrites",
        "--retries",
        str(args.retries),
        "--fragment-retries",
        str(args.fragment_retries),
        "--socket-timeout",
        str(args.socket_timeout),
        "--sleep-interval",
        str(args.sleep_interval),
        "--max-sleep-interval",
        str(args.max_sleep_interval),
        "--sleep-requests",
        str(args.sleep_requests),
        "--concurrent-fragments",
        "1",
        "--user-agent",
        args.user_agent,
        "--add-header",
        "Accept-Language:zh-CN,zh;q=0.9,en;q=0.8",
        "--format",
        args.format,
        "--output",
        str(task.output_dir / f"{task.base_name}.%(ext)s"),
    ]

    if args.verbose:
        command.append("--verbose")
    else:
        command.extend(["--no-warnings", "--progress"])

    if not args.allow_playlist:
        command.append("--no-playlist")
    if args.limit_rate:
        command.extend(["--limit-rate", args.limit_rate])
    if args.proxy:
        command.extend(["--proxy", args.proxy])
    if args.cookies_file:
        command.extend(["--cookies", args.cookies_file])
    elif args.cookies_browser:
        browser_spec = args.cookies_browser
        if args.cookies_profile:
            browser_spec = f"{browser_spec}:{args.cookies_profile}"
        command.extend(["--cookies-from-browser", browser_spec])
    if which("ffmpeg"):
        command.extend(["--merge-output-format", args.merge_output_format])
    else:
        logging.warning("ffmpeg not found, keeping the original container format.")

    command.append(task.url)
    return command


def download_task(task: VideoTask, args: argparse.Namespace) -> dict[str, Any]:
    task.output_dir.mkdir(parents=True, exist_ok=True)
    existing = find_existing_download(task.output_stem)
    if existing:
        logging.info("Skip existing file for row %s: %s", task.row_number, existing)
        return build_report_row(task, "skipped_existing", existing, "")

    max_attempts = args.task_retries + 1
    for attempt in range(1, max_attempts + 1):
        command = build_yt_dlp_command(task, args)
        logging.info(
            "Downloading row %s -> %s/%s/%s (attempt %s/%s)",
            task.row_number,
            task.folder_name,
            task.slice_name,
            task.base_name,
            attempt,
            max_attempts,
        )
        process = subprocess.run(command, check=False)
        if process.returncode == 0:
            final_file = find_existing_download(task.output_stem)
            return build_report_row(task, "downloaded", final_file, "")

        if attempt == max_attempts:
            break

        sleep_seconds = args.retry_backoff * (2 ** (attempt - 1)) + random.uniform(
            0, args.retry_jitter
        )
        logging.warning(
            "Row %s failed with exit code %s, retrying in %.1f seconds.",
            task.row_number,
            process.returncode,
            sleep_seconds,
        )
        time.sleep(sleep_seconds)

    return build_report_row(
        task,
        "failed",
        None,
        "yt-dlp exited with a non-zero status after all retries.",
    )


def build_report_row(
    task: VideoTask,
    status: str,
    final_file: Path | None,
    message: str,
) -> dict[str, Any]:
    return {
        "row_number": task.row_number,
        "folder_name": task.folder_name,
        "slice_name": task.slice_name,
        "video_name": task.base_name,
        "url": task.url,
        "target_path": str(final_file or task.output_stem),
        "status": status,
        "message": message,
    }


def write_report(report_path: Path, rows: Iterable[dict[str, Any]]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    rows = list(rows)
    with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "row_number",
                "folder_name",
                "slice_name",
                "video_name",
                "url",
                "target_path",
                "status",
                "message",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)


def print_plan(tasks: list[VideoTask]) -> None:
    for task in tasks:
        print(f"[row {task.row_number}] {task.url}")
        print(f"  -> {task.output_dir / task.base_name}")


def main() -> int:
    args = parse_args()
    configure_logging(args.verbose)

    input_path = Path(args.input).expanduser().resolve()
    output_root = Path(args.output_root).expanduser().resolve()
    report_path = (
        Path(args.report_file).expanduser().resolve()
        if args.report_file
        else output_root / "download_report.csv"
    )

    frame = load_manifest(input_path, parse_sheet(args.sheet))
    tasks = build_tasks(frame, args, output_root)

    if not tasks:
        logging.error("No valid tasks found in the input file.")
        return 1

    logging.info("Prepared %s download tasks.", len(tasks))
    print_plan(tasks)
    if args.plan_only:
        plan_rows = [build_report_row(task, "planned", None, "") for task in tasks]
        write_report(report_path, plan_rows)
        logging.info("Plan saved to %s", report_path)
        return 0

    results = [download_task(task, args) for task in tasks]
    write_report(report_path, results)

    failed_count = sum(result["status"] == "failed" for result in results)
    skipped_count = sum(result["status"] == "skipped_existing" for result in results)
    downloaded_count = sum(result["status"] == "downloaded" for result in results)
    logging.info(
        "Finished. downloaded=%s skipped=%s failed=%s report=%s",
        downloaded_count,
        skipped_count,
        failed_count,
        report_path,
    )
    return 1 if failed_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
