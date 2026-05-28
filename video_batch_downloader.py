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
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from shutil import which
from threading import Lock, Semaphore
from typing import Any, Iterable
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

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
MEDIA_SUFFIXES = {
    ".mp4",
    ".mkv",
    ".mov",
    ".flv",
    ".avi",
    ".webm",
    ".m4v",
    ".ts",
}
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)
_FFMPEG_WARNING_EMITTED = False
_OUTPUT_LOCK = Lock()


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


@dataclass
class OutputInspection:
    valid_file: Path | None
    temp_files: list[Path]
    zero_byte_files: list[Path]


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
        "--delete-incomplete-in-output-root",
        action="store_true",
        help=(
            "Delete temp files, zero-byte files, and empty task folders from "
            "--output-root before retrying incomplete tasks in place."
        ),
    )
    parser.add_argument(
        "--existing-output-root",
        default=None,
        help=(
            "Optional old output root to treat as already-downloaded reference files. "
            "Useful when moving unfinished work to another disk."
        ),
    )
    parser.add_argument(
        "--delete-incomplete-from-existing",
        action="store_true",
        help=(
            "Delete temp files, zero-byte files, and empty task folders from "
            "--existing-output-root before re-downloading missing items to --output-root."
        ),
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
        "--youtube-format-fallback",
        default="best",
        help=(
            "Fallback format used for YouTube when the primary format is unavailable. "
            "Default: %(default)s"
        ),
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
        "--youtube-cookies-file",
        default=None,
        help="Optional cookies.txt file used only for YouTube URLs.",
    )
    parser.add_argument(
        "--bilibili-cookies-file",
        default=None,
        help="Optional cookies.txt file used only for Bilibili URLs.",
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
        "--youtube-proxy",
        default=None,
        help=(
            "Proxy used only for YouTube URLs, e.g. http://127.0.0.1:7897. "
            "When set, Bilibili still uses a direct connection by default."
        ),
    )
    parser.add_argument(
        "--youtube-proxy-port",
        type=int,
        default=None,
        help=(
            "Convenience option for a local YouTube proxy port, e.g. 7897. "
            "Equivalent to --youtube-proxy http://127.0.0.1:<port>."
        ),
    )
    parser.add_argument(
        "--user-agent",
        default=DEFAULT_USER_AGENT,
        help="User-Agent header used by yt-dlp. Default is a recent Chrome UA.",
    )
    parser.add_argument(
        "--impersonate",
        default=None,
        help=(
            "Optional yt-dlp client impersonation target, e.g. chrome. "
            "Works best when curl-cffi is installed."
        ),
    )
    parser.add_argument(
        "--sleep-interval",
        type=float,
        default=2.0,
        help="Base wait between yt-dlp requests, in seconds. Default: %(default)s",
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
        "--task-sleep-min",
        type=float,
        default=1.0,
        help=(
            "Random sleep lower bound before each video task, in seconds. "
            "Default: %(default)s"
        ),
    )
    parser.add_argument(
        "--task-sleep-max",
        type=float,
        default=3.0,
        help=(
            "Random sleep upper bound before each video task, in seconds. "
            "Default: %(default)s"
        ),
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=3,
        help="Total concurrent download workers. Default: %(default)s",
    )
    parser.add_argument(
        "--bilibili-workers",
        type=int,
        default=1,
        help="Maximum concurrent Bilibili downloads. Default: %(default)s",
    )
    parser.add_argument(
        "--youtube-workers",
        type=int,
        default=3,
        help="Maximum concurrent YouTube downloads. Default: %(default)s",
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


def strip_media_extension(value: str) -> str:
    suffix = Path(value).suffix.lower()
    if suffix in MEDIA_SUFFIXES:
        return value[: -len(suffix)].rstrip()
    return value


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

    cleaned = sanitize_path_component(
        strip_media_extension(text),
        fallback=f"{default_index:0{width}d}",
    )
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


def task_output_stem_for_root(task: VideoTask, root: Path) -> Path:
    return root / task.folder_name / task.slice_name / task.base_name


def inspect_output_stem(output_stem: Path) -> OutputInspection:
    temp_candidates: list[Path] = []
    valid_candidates: list[Path] = []
    zero_byte_candidates: list[Path] = []

    for candidate in sorted(output_stem.parent.glob(f"{output_stem.name}.*")):
        if not candidate.is_file():
            continue
        if candidate.suffix.lower() in TEMP_FILE_SUFFIXES:
            temp_candidates.append(candidate)
            continue
        if candidate.stat().st_size <= 0:
            zero_byte_candidates.append(candidate)
            continue
        valid_candidates.append(candidate)

    return OutputInspection(
        valid_file=valid_candidates[0] if valid_candidates else None,
        temp_files=temp_candidates,
        zero_byte_files=zero_byte_candidates,
    )


def log_incomplete_artifacts(output_stem: Path, inspection: OutputInspection) -> None:
    if inspection.temp_files:
        logging.info(
            "Detected unfinished download artifacts for %s: %s",
            output_stem.name,
            ", ".join(str(path.name) for path in inspection.temp_files),
        )

    if inspection.zero_byte_files:
        logging.warning(
            "Ignoring zero-byte output files for %s: %s",
            output_stem.name,
            ", ".join(str(path.name) for path in inspection.zero_byte_files),
        )



def find_existing_download(output_stem: Path) -> Path | None:
    inspection = inspect_output_stem(output_stem)
    log_incomplete_artifacts(output_stem, inspection)
    return inspection.valid_file


def remove_empty_parents(path: Path, stop_at: Path) -> None:
    stop_at = stop_at.resolve()
    current = path.resolve()
    while current != stop_at:
        try:
            current.rmdir()
        except OSError:
            break
        current = current.parent


def cleanup_incomplete_outputs(output_stem: Path, root_to_prune: Path) -> list[Path]:
    inspection = inspect_output_stem(output_stem)
    deleted_paths: list[Path] = []

    for path in inspection.temp_files + inspection.zero_byte_files:
        path.unlink(missing_ok=True)
        deleted_paths.append(path)

    if output_stem.parent.exists():
        remove_empty_parents(output_stem.parent, root_to_prune)

    return deleted_paths


def is_bilibili_url(url: str) -> bool:
    hostname = urlparse(url).netloc.lower()
    return "bilibili.com" in hostname or hostname.endswith("b23.tv")


def is_youtube_url(url: str) -> bool:
    hostname = urlparse(url).netloc.lower()
    return (
        "youtube.com" in hostname
        or hostname.endswith("youtu.be")
        or hostname.endswith("youtube-nocookie.com")
    )


def task_platform(task: VideoTask) -> str:
    if is_bilibili_url(task.url):
        return "bilibili"
    if is_youtube_url(task.url):
        return "youtube"
    return "other"


def normalize_download_url(url: str) -> str:
    parsed = urlparse(url.strip())
    hostname = parsed.netloc.lower()
    if not parsed.scheme or not hostname:
        return url.strip()

    if "bilibili.com" in hostname or hostname.endswith("b23.tv"):
        kept_pairs = [
            (key, value)
            for key, value in parse_qsl(parsed.query, keep_blank_values=True)
            if key.lower() in {"p", "t"}
        ]
        return urlunparse(parsed._replace(query=urlencode(kept_pairs), fragment=""))

    return url.strip()


def build_site_headers(url: str) -> list[str]:
    if is_bilibili_url(url):
        return [
            "Accept-Language:zh-CN,zh;q=0.9,en;q=0.8",
            "Referer:https://www.bilibili.com/",
            "Origin:https://www.bilibili.com",
        ]

    return ["Accept-Language:zh-CN,zh;q=0.9,en;q=0.8"]


def resolve_youtube_proxy(args: argparse.Namespace) -> str | None:
    if args.youtube_proxy:
        return args.youtube_proxy
    if args.youtube_proxy_port:
        return f"http://127.0.0.1:{args.youtube_proxy_port}"
    return None


def resolve_proxy_for_task(task: VideoTask, args: argparse.Namespace) -> str | None:
    youtube_proxy = resolve_youtube_proxy(args)
    if is_youtube_url(task.url) and youtube_proxy:
        return youtube_proxy
    return args.proxy


def resolve_cookies_file_for_task(task: VideoTask, args: argparse.Namespace) -> str | None:
    if is_youtube_url(task.url) and args.youtube_cookies_file:
        return args.youtube_cookies_file
    if is_bilibili_url(task.url) and args.bilibili_cookies_file:
        return args.bilibili_cookies_file
    return args.cookies_file


@lru_cache(maxsize=1)
def ffmpeg_available() -> bool:
    return which("ffmpeg") is not None


def warn_ffmpeg_missing_once() -> None:
    global _FFMPEG_WARNING_EMITTED
    if _FFMPEG_WARNING_EMITTED or ffmpeg_available():
        return
    logging.warning("ffmpeg not found, keeping the original container format.")
    _FFMPEG_WARNING_EMITTED = True


def build_format_candidates(task: VideoTask, args: argparse.Namespace) -> list[str]:
    candidates = [args.format]
    if is_youtube_url(task.url) and args.youtube_format_fallback:
        candidates.append(args.youtube_format_fallback)

    unique_candidates: list[str] = []
    for candidate in candidates:
        if candidate and candidate not in unique_candidates:
            unique_candidates.append(candidate)
    return unique_candidates


def build_yt_dlp_command(
    task: VideoTask,
    args: argparse.Namespace,
    format_selector: str | None = None,
) -> list[str]:
    normalized_url = normalize_download_url(task.url)
    task_proxy = resolve_proxy_for_task(task, args)
    task_cookies_file = resolve_cookies_file_for_task(task, args)
    active_format = format_selector or args.format
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
        "--format",
        active_format,
        "--output",
        str(task.output_dir / f"{task.base_name}.%(ext)s"),
    ]

    for header in build_site_headers(normalized_url):
        command.extend(["--add-header", header])

    if args.verbose:
        command.append("--verbose")
    else:
        command.extend(["--no-warnings", "--progress"])

    if not args.allow_playlist:
        command.append("--no-playlist")
    if args.limit_rate:
        command.extend(["--limit-rate", args.limit_rate])
    if task_proxy:
        command.extend(["--proxy", task_proxy])
    if args.impersonate:
        command.extend(["--impersonate", args.impersonate])
    if task_cookies_file:
        command.extend(["--cookies", task_cookies_file])
    elif args.cookies_browser:
        browser_spec = args.cookies_browser
        if args.cookies_profile:
            browser_spec = f"{browser_spec}:{args.cookies_profile}"
        command.extend(["--cookies-from-browser", browser_spec])
    if ffmpeg_available():
        command.extend(["--merge-output-format", args.merge_output_format])
    else:
        warn_ffmpeg_missing_once()

    command.append(normalized_url)
    return command


def run_command_with_live_output(command: list[str]) -> tuple[int, str]:
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert process.stdout is not None

    captured_lines: list[str] = []
    for line in process.stdout:
        with _OUTPUT_LOCK:
            print(line, end="")
        captured_lines.append(line)
    return process.wait(), "".join(captured_lines)


def sleep_before_task(task: VideoTask, args: argparse.Namespace) -> None:
    lower = max(0.0, min(args.task_sleep_min, args.task_sleep_max))
    upper = max(0.0, max(args.task_sleep_min, args.task_sleep_max))
    if upper <= 0:
        return

    wait_seconds = random.uniform(lower, upper)
    logging.info(
        "Random wait %.1f seconds before row %s (%s/%s/%s).",
        wait_seconds,
        task.row_number,
        task.folder_name,
        task.slice_name,
        task.base_name,
    )
    time.sleep(wait_seconds)


def format_unavailable(output: str) -> bool:
    return "Requested format is not available" in output


def youtube_bot_verification_required(output: str) -> bool:
    lowered = output.lower()
    return "sign in to confirm" in lowered and "not a bot" in lowered


def build_failure_message(task: VideoTask, args: argparse.Namespace, output: str) -> str:
    if "Could not copy Chrome cookie database" in output:
        return (
            "yt-dlp could not read Chrome cookies because the browser cookie database "
            "is locked on Windows. Fully exit Chrome/Edge first (check Task Manager), "
            "then rerun. If you do not want to close the browser, export cookies to "
            "cookies.txt and use --cookies-file instead."
        )

    if is_youtube_url(task.url) and youtube_bot_verification_required(output):
        return (
            "YouTube requested login verification. Export a YouTube account cookies.txt "
            "from youtube.com and pass it with --youtube-cookies-file. Reusing a "
            "Bilibili cookies file will not help for YouTube."
        )

    if is_youtube_url(task.url) and format_unavailable(output):
        return (
            "The requested YouTube format was unavailable for this video. The downloader "
            "can retry with --youtube-format-fallback, or you can choose a more permissive "
            "main format selector."
        )

    if is_bilibili_url(task.url) and (
        "HTTP Error 412" in output or "Request is blocked by server (412)" in output
    ):
        if not resolve_cookies_file_for_task(task, args) and not args.cookies_browser:
            return (
                "Bilibili returned HTTP 412 (anti-bot). Open the same video in a normal "
                "browser first, then rerun with --cookies-browser chrome/edge/firefox "
                "or --cookies-file cookies.txt. If the IP is from a VPN or data center, "
                "switch to a residential/home network or proxy."
            )
        return (
            "Bilibili still returned HTTP 412. Refresh Bilibili in the same browser "
            "to obtain fresh cookies, confirm the exact link opens normally in that "
            "browser, then retry. If it still fails, the IP is likely being challenged; "
            "switch to a residential/home network or proxy, and optionally try "
            "--impersonate chrome after installing curl-cffi."
        )

    return "yt-dlp exited with a non-zero status after all retries."


def is_retryable_failure(task: VideoTask, output: str) -> bool:
    if "Could not copy Chrome cookie database" in output:
        return False
    if is_youtube_url(task.url) and youtube_bot_verification_required(output):
        return False
    if is_youtube_url(task.url) and format_unavailable(output):
        return False
    if is_bilibili_url(task.url) and (
        "HTTP Error 412" in output or "Request is blocked by server (412)" in output
    ):
        return False
    return True


def download_task(task: VideoTask, args: argparse.Namespace) -> dict[str, Any]:
    task.output_dir.mkdir(parents=True, exist_ok=True)
    if args.delete_incomplete_in_output_root:
        deleted_paths = cleanup_incomplete_outputs(task.output_stem, Path(args.output_root))
        if deleted_paths:
            logging.info(
                "Deleted incomplete artifacts for row %s from current output root: %s",
                task.row_number,
                ", ".join(str(path.name) for path in deleted_paths),
            )
    existing = find_existing_download(task.output_stem)
    if existing:
        logging.info("Skip existing file for row %s: %s", task.row_number, existing)
        return build_report_row(task, "skipped_existing", existing, "")

    existing_root = Path(args.existing_output_root) if args.existing_output_root else None
    if existing_root is not None:
        old_output_stem = task_output_stem_for_root(task, existing_root)
        old_existing = find_existing_download(old_output_stem)
        if old_existing:
            logging.info(
                "Skip existing file for row %s from old output root: %s",
                task.row_number,
                old_existing,
            )
            return build_report_row(task, "skipped_existing", old_existing, "")

        if args.delete_incomplete_from_existing:
            deleted_paths = cleanup_incomplete_outputs(old_output_stem, existing_root)
            if deleted_paths:
                logging.info(
                    "Deleted incomplete artifacts for row %s from old output root: %s",
                    task.row_number,
                    ", ".join(str(path.name) for path in deleted_paths),
                )

    max_attempts = args.task_retries + 1
    for attempt in range(1, max_attempts + 1):
        if attempt == 1:
            sleep_before_task(task, args)
        active_proxy = resolve_proxy_for_task(task, args)
        if active_proxy:
            logging.info(
                "Row %s routing via proxy %s",
                task.row_number,
                active_proxy,
            )
        elif is_bilibili_url(task.url) and resolve_youtube_proxy(args):
            logging.info(
                "Row %s uses direct connection for Bilibili.",
                task.row_number,
            )
        format_candidates = build_format_candidates(task, args)
        failure_message = "yt-dlp exited with a non-zero status after all retries."
        command_output = ""
        return_code = 1
        for format_index, format_candidate in enumerate(format_candidates, start=1):
            if len(format_candidates) > 1:
                logging.info(
                    "Row %s using format candidate %s/%s: %s",
                    task.row_number,
                    format_index,
                    len(format_candidates),
                    format_candidate,
                )
            logging.info(
                "Downloading row %s -> %s/%s/%s (attempt %s/%s)",
                task.row_number,
                task.folder_name,
                task.slice_name,
                task.base_name,
                attempt,
                max_attempts,
            )
            command = build_yt_dlp_command(task, args, format_selector=format_candidate)
            return_code, command_output = run_command_with_live_output(command)
            if return_code == 0:
                final_file = find_existing_download(task.output_stem)
                return build_report_row(task, "downloaded", final_file, "")

            if (
                is_youtube_url(task.url)
                and format_unavailable(command_output)
                and format_index < len(format_candidates)
            ):
                logging.warning(
                    "Row %s format '%s' unavailable, trying fallback format '%s'.",
                    task.row_number,
                    format_candidate,
                    format_candidates[format_index],
                )
                continue

            break

        failure_message = build_failure_message(task, args, command_output)
        if failure_message != "yt-dlp exited with a non-zero status after all retries.":
            logging.warning("Row %s hint: %s", task.row_number, failure_message)

        if not is_retryable_failure(task, command_output):
            logging.warning(
                "Row %s hit a known non-retryable error; skipping remaining retries.",
                task.row_number,
            )
            break

        if attempt == max_attempts:
            break

        sleep_seconds = args.retry_backoff * (2 ** (attempt - 1)) + random.uniform(
            0, args.retry_jitter
        )
        logging.warning(
            "Row %s failed with exit code %s, retrying in %.1f seconds.",
            task.row_number,
            return_code,
            sleep_seconds,
        )
        time.sleep(sleep_seconds)

    return build_report_row(
        task,
        "failed",
        None,
        failure_message,
    )


def build_semaphores(args: argparse.Namespace) -> dict[str, Semaphore]:
    return {
        "bilibili": Semaphore(max(1, args.bilibili_workers)),
        "youtube": Semaphore(max(1, args.youtube_workers)),
    }


def run_task_with_limits(
    task: VideoTask,
    args: argparse.Namespace,
    semaphores: dict[str, Semaphore],
) -> dict[str, Any]:
    platform = task_platform(task)
    semaphore = semaphores.get(platform)
    if semaphore is None:
        return download_task(task, args)

    with semaphore:
        return download_task(task, args)


def run_downloads(tasks: list[VideoTask], args: argparse.Namespace) -> list[dict[str, Any]]:
    worker_count = max(1, args.workers)
    if worker_count == 1:
        return [download_task(task, args) for task in tasks]

    semaphores = build_semaphores(args)
    results_by_index: dict[int, dict[str, Any]] = {}
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        future_to_index = {
            executor.submit(run_task_with_limits, task, args, semaphores): index
            for index, task in enumerate(tasks)
        }
        for future in as_completed(future_to_index):
            index = future_to_index[future]
            results_by_index[index] = future.result()

    return [results_by_index[index] for index in range(len(tasks))]


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
    args.output_root = str(output_root)
    if args.existing_output_root:
        args.existing_output_root = str(
            Path(args.existing_output_root).expanduser().resolve()
        )
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

    logging.info(
        "Starting downloads with workers=%s bilibili_workers=%s youtube_workers=%s",
        max(1, args.workers),
        max(1, args.bilibili_workers),
        max(1, args.youtube_workers),
    )
    results = run_downloads(tasks, args)
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
