#!/usr/bin/env python3
r"""统一生成数据集构建前后的轻量检查报告。

默认围绕当前 Git 追踪的新版 Excel 标注表工作：

    data\annotations\260311-260430.LX事件.lgm已核实.xlsx

报告输出到 reports/，包括标注质量、服务器录音库存、行级覆盖率和数据集类别统计。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path


DEFAULT_EXCEL = Path(r"data\annotations\260311-260430.LX事件.lgm已核实.xlsx")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="一键生成标注、录音覆盖和数据集统计报告。")
    parser.add_argument("--excel", type=Path, default=DEFAULT_EXCEL)
    parser.add_argument("--password", default=None)
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--audio-root", type=Path, default=Path(r"D:\AudioRecord"))
    parser.add_argument("--manifest", type=Path, default=Path(r"D:\SatelliteAudio\extraction_manifest.csv"))
    parser.add_argument("--raw-root", type=Path, default=Path(r"D:\SatelliteAudio\raw"))
    parser.add_argument("--output-dir", type=Path, default=Path("reports"))
    parser.add_argument("--clear-output-dir", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--station", default="LX")
    parser.add_argument("--source-ext", default="m4a")
    parser.add_argument("--start-col", default="F")
    parser.add_argument("--end-col", default="G")
    parser.add_argument("--label-col", default="T")
    parser.add_argument("--invalid-col", default="N")
    parser.add_argument("--invalid-marker", default="无效信号")
    parser.add_argument("--first-data-row", type=int, default=2)
    parser.add_argument("--skip-audio-inspection", action="store_true")
    parser.add_argument("--use-existing-inventory", action="store_true")
    parser.add_argument("--skip-dataset-counts", action="store_true")
    parser.add_argument("--no-progress", action="store_true")
    parser.add_argument("--progress-interval-sec", type=float, default=0.2)
    return parser.parse_args()


def column_index(column: str) -> int:
    result = 0
    for char in column.strip().upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid Excel column: {column}")
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def row_value(row: tuple[object, ...], index: int | None):
    if index is None or len(row) < index:
        return None
    return row[index - 1]


def normalize_cell_text(value) -> str:
    return re.sub(r"\s+", "", str(value or "").strip()).lower()


def load_workbook(path: Path, password: str | None):
    from openpyxl import load_workbook

    try:
        return load_workbook(path, data_only=True, read_only=True)
    except Exception:
        if not password:
            raise
        import msoffcrypto

        decrypted = io.BytesIO()
        with path.open("rb") as handle:
            office_file = msoffcrypto.OfficeFile(handle)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        decrypted.seek(0)
        return load_workbook(decrypted, data_only=True, read_only=True)


def parse_excel_datetime(value) -> dt.datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time.min)
    if isinstance(value, (int, float)):
        from openpyxl.utils.datetime import from_excel

        parsed = from_excel(value)
        if isinstance(parsed, dt.datetime):
            return parsed.replace(tzinfo=None)
        if isinstance(parsed, dt.date):
            return dt.datetime.combine(parsed, dt.time.min)
        return None

    text = str(value).strip()
    if not text:
        return None
    for fmt in (
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
        "%Y%m%d-%H%M%S",
        "%Y%m%d-%H%M",
    ):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unsupported datetime value: {value!r}")


class ProgressReporter:
    def __init__(self, enabled: bool, total: int | None, interval_sec: float, stage: str) -> None:
        self.enabled = enabled
        self.total = total if total and total > 0 else None
        self.interval_sec = max(interval_sec, 0.0)
        self.stage = stage
        self.last_update = 0.0
        self.started_at = time.monotonic()

    def update(self, current: int, force: bool = False) -> None:
        if not self.enabled:
            return
        now = time.monotonic()
        if not force and now - self.last_update < self.interval_sec:
            return
        self.last_update = now
        elapsed = format_duration(now - self.started_at)
        if self.total:
            ratio = min(current / self.total, 1.0)
            width = 30
            filled = int(width * ratio)
            bar = "#" * filled + "-" * (width - filled)
            message = f"\r{self.stage} [{bar}] {current}/{self.total} {ratio * 100:5.1f}% elapsed={elapsed}"
        else:
            message = f"\r{self.stage} processed={current} elapsed={elapsed}"
        sys.stderr.write(message[:180].ljust(180))
        sys.stderr.flush()

    def finish(self) -> None:
        if self.enabled:
            sys.stderr.write("\n")
            sys.stderr.flush()


def format_duration(seconds: float) -> str:
    seconds_int = int(seconds)
    hours, remainder = divmod(seconds_int, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def safe_clear_report_dir(output_dir: Path) -> None:
    resolved = output_dir.resolve()
    if resolved.anchor == str(resolved):
        raise SystemExit(f"Refusing to clear drive root: {resolved}")
    if resolved.name.lower() != "reports":
        raise SystemExit(f"Refusing to clear {resolved}: report output dir must be named 'reports'.")
    if not resolved.exists():
        return
    if not resolved.is_dir():
        raise SystemExit(f"Refusing to clear non-directory report path: {resolved}")
    for child in resolved.iterdir():
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()


def generate_annotation_quality_report(args: argparse.Namespace) -> None:
    workbook = load_workbook(args.excel, args.password)
    worksheet = workbook[args.sheet] if args.sheet else workbook.active
    start_idx = column_index(args.start_col)
    end_idx = column_index(args.end_col)
    label_idx = column_index(args.label_col)
    invalid_idx = None if args.invalid_col.lower() in {"", "none"} else column_index(args.invalid_col)
    marker_text = normalize_cell_text(args.invalid_marker)

    total_rows = max((getattr(worksheet, "max_row", 0) or 0) - args.first_data_row + 1, 0)
    progress = ProgressReporter(not args.no_progress, total_rows, args.progress_interval_sec, "Annotation")
    rows: list[dict[str, object]] = []
    counts: dict[str, int] = {}
    unexpected_invalid_values: dict[str, int] = {}
    processed = 0

    try:
        progress.update(0, force=True)
        for row_index, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            if row_index < args.first_data_row:
                continue
            processed += 1
            progress.update(processed)
            raw_start = row_value(row, start_idx)
            raw_end = row_value(row, end_idx)
            label = str(row_value(row, label_idx) or "").strip()
            invalid_value = str(row_value(row, invalid_idx) or "").strip()
            has_any_key_value = any(value not in (None, "") for value in (raw_start, raw_end, label, invalid_value))
            if not has_any_key_value:
                continue

            invalid_norm = normalize_cell_text(invalid_value)
            is_invalid_signal = invalid_norm == marker_text
            unexpected_invalid = bool(invalid_norm and invalid_norm != marker_text)
            if unexpected_invalid:
                unexpected_invalid_values[invalid_value] = unexpected_invalid_values.get(invalid_value, 0) + 1

            try:
                start = parse_excel_datetime(raw_start)
                end = parse_excel_datetime(raw_end)
                time_ok = bool(start and end and end > start)
                time_message = "" if time_ok else "missing_or_invalid_time"
            except Exception as exc:
                start = None
                end = None
                time_ok = False
                time_message = str(exc)

            label_missing = not label or label in {"?", "？"}
            if unexpected_invalid:
                status = "unexpected_invalid_value"
                message = f"invalid column contains {invalid_value!r}"
            elif is_invalid_signal:
                status = "invalid_signal"
                message = "marked invalid signal"
            elif label_missing:
                status = "missing_label"
                message = "missing satellite label"
            elif not time_ok:
                status = "invalid_time"
                message = time_message
            else:
                status = "ready_for_extraction"
                message = ""

            counts[status] = counts.get(status, 0) + 1
            rows.append(
                {
                    "row_index": row_index,
                    "label": label,
                    "invalid_signal": invalid_value,
                    "start_raw": raw_start,
                    "end_raw": raw_end,
                    "start_time": start.isoformat(sep=" ") if start else "",
                    "end_time": end.isoformat(sep=" ") if end else "",
                    "annotation_status": status,
                    "message": message,
                }
            )
    finally:
        progress.update(processed, force=True)
        progress.finish()

    report_path = args.output_dir / "annotation_quality_report.csv"
    summary_path = args.output_dir / "annotation_quality_summary.txt"
    write_csv(
        report_path,
        rows,
        [
            "row_index",
            "label",
            "invalid_signal",
            "start_raw",
            "end_raw",
            "start_time",
            "end_time",
            "annotation_status",
            "message",
        ],
    )

    lines = [
        "Annotation Quality Report",
        f"excel: {args.excel}",
        f"sheet: {worksheet.title}",
        f"columns: start={args.start_col}, end={args.end_col}, label={args.label_col}, invalid={args.invalid_col}",
        f"rows_checked: {len(rows)}",
        "",
        "annotation_status_counts:",
    ]
    for key in sorted(counts):
        lines.append(f"  {key}: {counts[key]}")
    lines.append("")
    lines.append("unexpected_invalid_values:")
    if unexpected_invalid_values:
        for value, count in sorted(unexpected_invalid_values.items()):
            lines.append(f"  {value}: {count}")
    else:
        lines.append("  none")
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def run_child(command: list[str]) -> None:
    result = subprocess.run(command)
    if result.returncode != 0:
        raise SystemExit(result.returncode)


def main() -> int:
    args = parse_args()
    if args.clear_output_dir:
        safe_clear_report_dir(args.output_dir)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    generate_annotation_quality_report(args)

    server_report_dir = args.output_dir / "server_inspection"
    ran_audio_inspection = False
    if not args.skip_audio_inspection:
        inspect_command = [
            sys.executable,
            str(Path(__file__).with_name("inspect_server_audio.py")),
            "--audio-root",
            str(args.audio_root),
            "--output-dir",
            str(server_report_dir),
            "--station",
            args.station,
            "--source-ext",
            args.source_ext,
        ]
        if args.manifest.exists():
            inspect_command.extend(["--manifest", str(args.manifest)])
        if args.no_progress:
            inspect_command.append("--no-progress")
        run_child(inspect_command)
        ran_audio_inspection = True

    inventory = server_report_dir / "audio_inventory.csv"
    if inventory.exists() and (ran_audio_inspection or args.use_existing_inventory):
        coverage_command = [
            sys.executable,
            str(Path(__file__).with_name("report_uncovered_rows.py")),
            "--excel",
            str(args.excel),
            "--inventory",
            str(inventory),
            "--output",
            str(server_report_dir / "row_coverage_report.csv"),
            "--summary",
            str(server_report_dir / "row_coverage_summary.txt"),
            "--start-col",
            args.start_col,
            "--end-col",
            args.end_col,
            "--label-col",
            args.label_col,
            "--invalid-col",
            args.invalid_col,
            "--invalid-marker",
            args.invalid_marker,
        ]
        if args.password:
            coverage_command.extend(["--password", args.password])
        if args.sheet:
            coverage_command.extend(["--sheet", args.sheet])
        if args.no_progress:
            coverage_command.append("--no-progress")
        run_child(coverage_command)

    if not args.skip_dataset_counts:
        dataset_command = [
            sys.executable,
            str(Path(__file__).with_name("report_dataset_counts.py")),
            "--output",
            str(args.output_dir / "dataset_counts.csv"),
            "--summary",
            str(args.output_dir / "dataset_counts_summary.txt"),
        ]
        if args.manifest.exists():
            dataset_command.extend(["--manifest", str(args.manifest)])
            run_child(dataset_command)
        elif args.raw_root.exists():
            dataset_command.extend(["--raw-root", str(args.raw_root)])
            run_child(dataset_command)

    print(f"Reports written to: {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
