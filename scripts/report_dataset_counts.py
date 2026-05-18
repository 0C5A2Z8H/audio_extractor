#!/usr/bin/env python3
"""按卫星或类别统计已提取数据集的音频数量。"""

from __future__ import annotations

import argparse
import csv
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path


AUDIO_EXTS = {".wav", ".flac", ".mp3", ".m4a", ".aac", ".ogg"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="统计每个卫星或类别下已有多少条已提取音频。"
    )
    parser.add_argument("--raw-root", type=Path, default=None, help="Directory like D:\\SatelliteAudio\\raw.")
    parser.add_argument("--manifest", type=Path, default=None, help="Extraction manifest CSV.")
    parser.add_argument("--output", type=Path, required=True, help="Output per-class CSV report.")
    parser.add_argument("--summary", type=Path, default=None, help="Optional text summary path.")
    parser.add_argument("--no-progress", action="store_true")
    parser.add_argument("--progress-interval-sec", type=float, default=0.2)
    return parser.parse_args()


def raise_csv_field_limit() -> None:
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            return
        except OverflowError:
            limit = int(limit / 10)


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


def rows_from_raw_root(raw_root: Path, progress: ProgressReporter | None = None) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    audio_paths = [
        audio_path
        for class_dir in sorted(path for path in raw_root.iterdir() if path.is_dir())
        for audio_path in sorted(path for path in class_dir.rglob("*") if path.is_file())
        if audio_path.suffix.lower() in AUDIO_EXTS
    ]
    if progress:
        progress.total = len(audio_paths)
        progress.update(0, force=True)
    for index, audio_path in enumerate(audio_paths, start=1):
        rows.append(
            {
                "label": audio_path.parent.name,
                "path": str(audio_path),
                "decode_status": "recover" if "_recover" in audio_path.stem else "primary",
                "exists": True,
            }
        )
        if progress:
            progress.update(index)
    if progress:
        progress.update(len(audio_paths), force=True)
        progress.finish()
    return rows


def manifest_data_row_count(manifest: Path) -> int:
    with manifest.open("r", encoding="utf-8-sig", newline="") as handle:
        return max(sum(1 for _ in handle) - 1, 0)


def rows_from_manifest(manifest: Path, progress: ProgressReporter | None = None) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    total = manifest_data_row_count(manifest)
    if progress:
        progress.total = total
        progress.update(0, force=True)
    with manifest.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for index, row in enumerate(reader, start=1):
            if row.get("status") == "ok":
                output_path = row.get("output_path", "")
                rows.append(
                    {
                        "label": row.get("label", ""),
                        "path": output_path,
                        "decode_status": row.get("decode_status", "") or ("recover" if "_recover" in Path(output_path).stem else "primary"),
                        "exists": Path(output_path).exists() if output_path else False,
                    }
                )
            if progress:
                progress.update(index)
    if progress:
        progress.update(total, force=True)
        progress.finish()
    return rows


def summarize(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        label = str(row.get("label") or "").strip()
        if label:
            grouped[label].append(row)

    report_rows: list[dict[str, object]] = []
    for label in sorted(grouped):
        items = grouped[label]
        decode_counts = Counter(str(item.get("decode_status") or "unknown") for item in items)
        report_rows.append(
            {
                "label": label,
                "count": len(items),
                "primary_count": decode_counts.get("primary", 0),
                "recover_count": decode_counts.get("recover", 0),
                "other_decode_count": len(items) - decode_counts.get("primary", 0) - decode_counts.get("recover", 0),
            }
        )
    report_rows.sort(key=lambda item: (-int(item["count"]), str(item["label"])))
    return report_rows


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["label", "count", "primary_count", "recover_count", "other_decode_count"],
        )
        writer.writeheader()
        writer.writerows(rows)


def write_summary(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    total = sum(int(row["count"]) for row in rows)
    recover_total = sum(int(row["recover_count"]) for row in rows)
    lines = [
        "Dataset Count Report",
        f"satellite_count: {len(rows)}",
        f"audio_count: {total}",
        f"recover_count: {recover_total}",
        "",
        "per_satellite:",
    ]
    for row in rows:
        lines.append(
            f"  {row['label']}: {row['count']} "
            f"(primary={row['primary_count']}, recover={row['recover_count']}, other={row['other_decode_count']})"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    raise_csv_field_limit()
    if bool(args.raw_root) == bool(args.manifest):
        raise SystemExit("Pass exactly one of --raw-root or --manifest.")

    progress = ProgressReporter(not args.no_progress, None, args.progress_interval_sec, "Dataset counts")
    if args.raw_root:
        rows = rows_from_raw_root(args.raw_root, progress)
    else:
        rows = rows_from_manifest(args.manifest, progress)

    report_rows = summarize(rows)
    write_csv(args.output, report_rows)
    if args.summary:
        write_summary(args.summary, report_rows)

    print(f"satellite_count={len(report_rows)}")
    print(f"audio_count={sum(int(row['count']) for row in report_rows)}")
    print(f"recover_count={sum(int(row['recover_count']) for row in report_rows)}")
    print(f"wrote={args.output}")
    if args.summary:
        print(f"summary={args.summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
