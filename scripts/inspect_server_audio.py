#!/usr/bin/env python3
"""Inspect server-side AudioRecord files and compare them with a manifest."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
import shutil
import subprocess
from collections import Counter, defaultdict
from pathlib import Path


SOURCE_RE = re.compile(r"([A-Za-z]+)(\d{8})-(\d{6})\.(\w+)$")
MISSING_RE = re.compile(r"Missing source recording:\s*(.+)$")
COMMAND_INPUT_RE = re.compile(r"\s-i\s+([A-Za-z]:\\.*?\.m4a)\s+-ss\s+")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect D:\\AudioRecord and summarize files needed by extraction_manifest.csv."
    )
    parser.add_argument("--audio-root", type=Path, default=Path(r"D:\AudioRecord"))
    parser.add_argument("--manifest", type=Path, default=None)
    parser.add_argument("--output-dir", type=Path, default=Path(r"D:\audio_extractor\server_inspection"))
    parser.add_argument("--station", default="LX")
    parser.add_argument("--source-ext", default="m4a")
    parser.add_argument("--dates", nargs="*", default=None, help="YYYYMMDD dates to inspect.")
    parser.add_argument("--ffprobe", default="ffprobe")
    parser.add_argument("--skip-ffprobe", action="store_true")
    return parser.parse_args()


def parse_source_time(path: Path) -> dt.datetime | None:
    match = SOURCE_RE.match(path.name)
    if not match:
        return None
    _, day, clock, _ = match.groups()
    return dt.datetime.strptime(day + clock, "%Y%m%d%H%M%S")


def ffprobe_duration(ffprobe: str, path: Path) -> tuple[float | None, str]:
    command = [
        ffprobe,
        "-v",
        "error",
        "-show_entries",
        "format=duration",
        "-of",
        "json",
        str(path),
    ]
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        return None, result.stderr.strip()
    try:
        payload = json.loads(result.stdout)
        duration = float(payload["format"]["duration"])
        return duration, ""
    except Exception as exc:
        return None, f"Could not parse ffprobe output: {exc}"


def read_manifest_sources(manifest: Path | None) -> tuple[set[Path], Counter, list[dict[str, str]]]:
    needed: set[Path] = set()
    status_counts: Counter = Counter()
    failed_rows: list[dict[str, str]] = []
    if not manifest or not manifest.exists():
        return needed, status_counts, failed_rows

    with manifest.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            status = row.get("status", "")
            status_counts[status] += 1
            source_files = row.get("source_files", "")
            if source_files:
                for item in source_files.split(";"):
                    if item:
                        needed.add(Path(item))
            message = row.get("message", "")
            missing = MISSING_RE.search(message)
            if missing:
                needed.add(Path(missing.group(1)))
            command_input = COMMAND_INPUT_RE.search(message)
            if command_input:
                needed.add(Path(command_input.group(1)))
            if status == "failed":
                failed_rows.append(row)
    return needed, status_counts, failed_rows


def infer_dates(audio_root: Path, station: str, dates: list[str] | None, needed: set[Path]) -> list[str]:
    if dates:
        return sorted(dates)
    inferred = set()
    for path in needed:
        match = re.search(rf"{re.escape(station)}(\d{{8}})", str(path))
        if match:
            inferred.add(match.group(1))
    if inferred:
        return sorted(inferred)
    return []


def list_audio_files(audio_root: Path, station: str, source_ext: str, dates: list[str]) -> list[Path]:
    files: list[Path] = []
    for day in dates:
        day_dir = audio_root / f"{station}{day}"
        if not day_dir.exists():
            continue
        files.extend(sorted(day_dir.glob(f"{station}{day}-*.{source_ext.lstrip('.')}")))
    return files


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    needed, status_counts, failed_rows = read_manifest_sources(args.manifest)
    dates = infer_dates(args.audio_root, args.station, args.dates, needed)
    if not dates:
        date_dirs = sorted(args.audio_root.glob(f"{args.station}[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]"))
        dates = [path.name.replace(args.station, "", 1) for path in date_dirs]

    ffprobe_available = shutil.which(args.ffprobe) is not None
    do_ffprobe = (not args.skip_ffprobe) and ffprobe_available

    inventory_rows: list[dict[str, object]] = []
    files_by_day: defaultdict[str, int] = defaultdict(int)
    files = list_audio_files(args.audio_root, args.station, args.source_ext, dates)

    for file_path in files:
        start = parse_source_time(file_path)
        duration, probe_error = (None, "")
        if do_ffprobe:
            duration, probe_error = ffprobe_duration(args.ffprobe, file_path)
        end = start + dt.timedelta(seconds=duration) if start and duration is not None else None
        files_by_day[file_path.parent.name] += 1
        stat = file_path.stat()
        inventory_rows.append(
            {
                "path": str(file_path),
                "day_dir": file_path.parent.name,
                "name": file_path.name,
                "start_time": start.isoformat(sep=" ") if start else "",
                "duration_sec": f"{duration:.3f}" if duration is not None else "",
                "end_time": end.isoformat(sep=" ") if end else "",
                "size_bytes": stat.st_size,
                "modified_time": dt.datetime.fromtimestamp(stat.st_mtime).isoformat(sep=" "),
                "ffprobe_error": probe_error,
            }
        )

    needed_rows: list[dict[str, object]] = []
    for path in sorted(needed, key=str):
        needed_rows.append(
            {
                "path": str(path),
                "exists": path.exists(),
                "size_bytes": path.stat().st_size if path.exists() else "",
            }
        )

    failed_type_counts = Counter()
    for row in failed_rows:
        message = row.get("message", "")
        if message.startswith("Missing source recording:"):
            failed_type_counts["missing_source"] += 1
        elif message.startswith("Command failed:"):
            failed_type_counts["ffmpeg_failed"] += 1
        else:
            failed_type_counts["other_failed"] += 1

    write_csv(
        args.output_dir / "audio_inventory.csv",
        inventory_rows,
        [
            "path",
            "day_dir",
            "name",
            "start_time",
            "duration_sec",
            "end_time",
            "size_bytes",
            "modified_time",
            "ffprobe_error",
        ],
    )
    write_csv(args.output_dir / "needed_sources.csv", needed_rows, ["path", "exists", "size_bytes"])

    report_lines = [
        "Server Audio Inspection Report",
        f"audio_root: {args.audio_root}",
        f"manifest: {args.manifest or ''}",
        f"output_dir: {args.output_dir}",
        f"dates: {', '.join(dates)}",
        f"ffprobe_available: {ffprobe_available}",
        f"ffprobe_used: {do_ffprobe}",
        "",
        "manifest_status_counts:",
    ]
    for key, count in sorted(status_counts.items()):
        report_lines.append(f"  {key}: {count}")
    report_lines.append("")
    report_lines.append("failed_type_counts:")
    for key, count in sorted(failed_type_counts.items()):
        report_lines.append(f"  {key}: {count}")
    report_lines.append("")
    report_lines.append("files_by_day:")
    for key in sorted(files_by_day):
        report_lines.append(f"  {key}: {files_by_day[key]}")
    report_lines.append("")
    report_lines.append("missing_needed_sources:")
    missing_needed = [row["path"] for row in needed_rows if not row["exists"]]
    for path in missing_needed[:200]:
        report_lines.append(f"  {path}")
    if len(missing_needed) > 200:
        report_lines.append(f"  ... {len(missing_needed) - 200} more")

    (args.output_dir / "summary.txt").write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    print(f"Wrote inspection report to: {args.output_dir}")
    print(f"Inventory rows: {len(inventory_rows)}")
    print(f"Needed sources: {len(needed_rows)}")
    print(f"Missing needed sources: {len([row for row in needed_rows if not row['exists']])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
