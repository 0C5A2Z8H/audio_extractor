#!/usr/bin/env python3
"""Summarize extracted dataset counts by satellite/category."""

from __future__ import annotations

import argparse
import csv
from collections import Counter, defaultdict
from pathlib import Path


AUDIO_EXTS = {".wav", ".flac", ".mp3", ".m4a", ".aac", ".ogg"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Report how many extracted audio clips exist for each satellite/category."
    )
    parser.add_argument("--raw-root", type=Path, default=None, help="Directory like D:\\SatelliteAudio\\raw.")
    parser.add_argument("--manifest", type=Path, default=None, help="Extraction manifest CSV.")
    parser.add_argument("--output", type=Path, required=True, help="Output per-class CSV report.")
    parser.add_argument("--summary", type=Path, default=None, help="Optional text summary path.")
    return parser.parse_args()


def rows_from_raw_root(raw_root: Path) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for class_dir in sorted(path for path in raw_root.iterdir() if path.is_dir()):
        for audio_path in sorted(path for path in class_dir.rglob("*") if path.is_file()):
            if audio_path.suffix.lower() not in AUDIO_EXTS:
                continue
            rows.append(
                {
                    "label": class_dir.name,
                    "path": str(audio_path),
                    "decode_status": "recover" if "_recover" in audio_path.stem else "primary",
                    "exists": True,
                }
            )
    return rows


def rows_from_manifest(manifest: Path) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    with manifest.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if row.get("status") != "ok":
                continue
            output_path = row.get("output_path", "")
            rows.append(
                {
                    "label": row.get("label", ""),
                    "path": output_path,
                    "decode_status": row.get("decode_status", "") or ("recover" if "_recover" in Path(output_path).stem else "primary"),
                    "exists": Path(output_path).exists() if output_path else False,
                }
            )
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
    if bool(args.raw_root) == bool(args.manifest):
        raise SystemExit("Pass exactly one of --raw-root or --manifest.")

    if args.raw_root:
        rows = rows_from_raw_root(args.raw_root)
    else:
        rows = rows_from_manifest(args.manifest)

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
