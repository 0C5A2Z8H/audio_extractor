#!/usr/bin/env python3
"""Report labelled Excel rows that are not fully covered by real recordings."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare annotation rows with server audio coverage intervals."
    )
    parser.add_argument("--excel", type=Path, required=True)
    parser.add_argument("--password", default=None)
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--inventory", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--summary", type=Path, default=None)
    parser.add_argument("--start-col", default="B")
    parser.add_argument("--end-col", default="C")
    parser.add_argument("--label-col", default="AK")
    parser.add_argument("--first-data-row", type=int, default=2)
    parser.add_argument("--ignore-dates", nargs="*", default=["20260311", "20260312", "20260313"])
    return parser.parse_args()


def column_index(column: str) -> int:
    result = 0
    for char in column.strip().upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid Excel column: {column}")
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


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
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unsupported datetime value: {value!r}")


def sanitize_label(value) -> str:
    return str(value).strip() if value is not None else ""


def read_inventory(path: Path) -> list[dict[str, object]]:
    intervals: list[dict[str, object]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not row.get("start_time") or not row.get("end_time"):
                continue
            if row.get("ffprobe_error"):
                continue
            intervals.append(
                {
                    "path": row["path"],
                    "start": dt.datetime.fromisoformat(row["start_time"]),
                    "end": dt.datetime.fromisoformat(row["end_time"]),
                }
            )
    intervals.sort(key=lambda item: item["start"])
    return intervals


def coverage_parts(start: dt.datetime, end: dt.datetime, intervals: list[dict[str, object]]) -> list[dict[str, object]]:
    parts = []
    for item in intervals:
        item_start = item["start"]
        item_end = item["end"]
        if item_end <= start:
            continue
        if item_start >= end:
            break
        overlap_start = max(start, item_start)
        overlap_end = min(end, item_end)
        if overlap_start < overlap_end:
            parts.append({"start": overlap_start, "end": overlap_end, "path": item["path"]})
    return parts


def classify_coverage(start: dt.datetime, end: dt.datetime, parts: list[dict[str, object]]) -> tuple[str, str, float]:
    if not parts:
        return "no_coverage", "no recording overlaps this interval", 0.0

    cursor = start
    covered_seconds = 0.0
    gaps: list[str] = []
    for part in parts:
        part_start = part["start"]
        part_end = part["end"]
        if part_start > cursor:
            gaps.append(f"{cursor:%Y-%m-%d %H:%M:%S} -> {part_start:%Y-%m-%d %H:%M:%S}")
        if part_end > cursor:
            covered_seconds += (part_end - max(cursor, part_start)).total_seconds()
            cursor = part_end
    if cursor < end:
        gaps.append(f"{cursor:%Y-%m-%d %H:%M:%S} -> {end:%Y-%m-%d %H:%M:%S}")

    total_seconds = (end - start).total_seconds()
    ratio = covered_seconds / total_seconds if total_seconds > 0 else 0.0
    if not gaps:
        status = "fully_covered"
        message = ""
    elif ratio > 0:
        status = "partial_coverage"
        message = "; ".join(gaps)
    else:
        status = "no_coverage"
        message = "no recording overlaps this interval"
    return status, message, ratio


def main() -> int:
    args = parse_args()
    workbook = load_workbook(args.excel, args.password)
    worksheet = workbook[args.sheet] if args.sheet else workbook.active
    intervals = read_inventory(args.inventory)
    ignore_dates = set(args.ignore_dates)

    start_idx = column_index(args.start_col)
    end_idx = column_index(args.end_col)
    label_idx = column_index(args.label_col)

    rows: list[dict[str, object]] = []
    counts: dict[str, int] = {}

    for row_index, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        if row_index < args.first_data_row:
            continue
        label_raw = row[label_idx - 1] if len(row) >= label_idx else None
        label = sanitize_label(label_raw)
        if not label or label in {"?", "？"}:
            continue
        try:
            start = parse_excel_datetime(row[start_idx - 1] if len(row) >= start_idx else None)
            end = parse_excel_datetime(row[end_idx - 1] if len(row) >= end_idx else None)
            if start is None or end is None or end <= start:
                status = "invalid_time"
                message = "missing or invalid start/end time"
                ratio = 0.0
                parts: list[dict[str, object]] = []
            elif start.strftime("%Y%m%d") in ignore_dates:
                status = "ignored_deleted_date"
                message = "date intentionally removed"
                ratio = 0.0
                parts = []
            else:
                parts = coverage_parts(start, end, intervals)
                status, message, ratio = classify_coverage(start, end, parts)

            counts[status] = counts.get(status, 0) + 1
            rows.append(
                {
                    "row_index": row_index,
                    "label": label,
                    "start_time": start.isoformat(sep=" ") if start else "",
                    "end_time": end.isoformat(sep=" ") if end else "",
                    "duration_sec": f"{(end - start).total_seconds():.3f}" if start and end and end > start else "",
                    "coverage_status": status,
                    "coverage_ratio": f"{ratio:.6f}",
                    "message": message,
                    "source_files": ";".join(str(part["path"]) for part in parts),
                }
            )
        except Exception as exc:
            counts["error"] = counts.get("error", 0) + 1
            rows.append(
                {
                    "row_index": row_index,
                    "label": label,
                    "start_time": "",
                    "end_time": "",
                    "duration_sec": "",
                    "coverage_status": "error",
                    "coverage_ratio": "0.000000",
                    "message": str(exc),
                    "source_files": "",
                }
            )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "row_index",
                "label",
                "start_time",
                "end_time",
                "duration_sec",
                "coverage_status",
                "coverage_ratio",
                "message",
                "source_files",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)

    if args.summary:
        lines = ["coverage_status_counts:"]
        for key in sorted(counts):
            lines.append(f"  {key}: {counts[key]}")
        impossible = [row for row in rows if row["coverage_status"] in {"no_coverage", "partial_coverage", "invalid_time", "error"}]
        lines.append("")
        lines.append(f"impossible_or_incomplete_rows: {len(impossible)}")
        lines.append(",".join(str(row["row_index"]) for row in impossible))
        args.summary.parent.mkdir(parents=True, exist_ok=True)
        args.summary.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"Wrote {len(rows)} labelled rows to {args.output}")
    for key in sorted(counts):
        print(f"{key}: {counts[key]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
