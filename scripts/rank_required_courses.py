#!/usr/bin/env python3
"""Build ranked required-course ratings from the 2024 MAcc exit survey."""

from __future__ import annotations

import csv
import math
import re
import struct
import zlib
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

INPUT_PATH = Path("data/grad_exit_survey_2024.xlsx")
FALLBACK_INPUT_PATH = Path("Grad Program Exit Survey Data 2024 (1).xlsx")
OUTPUT_CSV = Path("outputs/rank_order.csv")
OUTPUT_PNG = Path("outputs/rank_order.png")

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - 64)
    return value - 1


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("x:si", NS):
        parts = [node.text or "" for node in si.findall(".//x:t", NS)]
        values.append("".join(parts))
    return values


def read_sheet_rows(path: Path) -> list[dict[int, str]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_xml = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

    rows: list[dict[int, str]] = []
    for row in sheet_xml.find("x:sheetData", NS).findall("x:row", NS):
        row_values: dict[int, str] = {}
        for cell in row.findall("x:c", NS):
            idx = col_index(cell.attrib["r"])
            raw = cell.find("x:v", NS)
            if raw is None:
                row_values[idx] = ""
                continue
            t = cell.attrib.get("t")
            if t == "s":
                row_values[idx] = shared_strings[int(raw.text)]
            else:
                row_values[idx] = raw.text or ""
        rows.append(row_values)
    return rows


def normalize_course_name(label: str) -> str:
    part = label.split(" - ")[-1].strip()
    return re.sub(r"\s+", " ", part)


def parse_number(value: str) -> float | None:
    text = str(value).strip()
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if math.isnan(number):
        return None
    return number


def extract_core_course_columns(question_row: dict[int, str]) -> dict[int, str]:
    course_columns: dict[int, str] = {}
    for idx, label in question_row.items():
        lowered = label.lower()
        if "macc core course" in lowered and "- acc" in lowered:
            course_columns[idx] = normalize_course_name(label)
    if not course_columns:
        raise ValueError("No required-core course columns were detected in the workbook.")
    return course_columns


def aggregate_course_ratings(rows: list[dict[int, str]]) -> list[dict[str, object]]:
    if len(rows) < 4:
        raise ValueError("Workbook does not contain expected Qualtrics header structure.")

    question_row = rows[1]
    data_rows = rows[3:]
    course_columns = extract_core_course_columns(question_row)

    ratings_by_course: dict[str, list[float]] = defaultdict(list)

    for row in data_rows:
        for idx, course in course_columns.items():
            rating = parse_number(row.get(idx, ""))
            if rating is None:
                continue
            ratings_by_course[course].append(rating)

    results: list[dict[str, object]] = []
    for course, values in ratings_by_course.items():
        if not values:
            continue
        mean_value = sum(values) / len(values)
        results.append(
            {
                "course": course,
                "mean_rating": mean_value,
                "response_count": len(values),
            }
        )

    results.sort(key=lambda item: (-item["mean_rating"], item["course"]))
    for i, item in enumerate(results, start=1):
        item["rank"] = i
    return results


def write_csv(rows: list[dict[str, object]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["rank", "course", "mean_rating", "response_count"],
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "rank": row["rank"],
                    "course": row["course"],
                    "mean_rating": f"{row['mean_rating']:.4f}",
                    "response_count": row["response_count"],
                }
            )


def make_png(width: int, height: int, rgb_rows: list[bytes], path: Path) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack("!I", len(data))
            + tag
            + data
            + struct.pack("!I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    raw = b"".join(b"\x00" + row for row in rgb_rows)
    ihdr = struct.pack("!IIBBBBB", width, height, 8, 2, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(png)


def draw_bar_chart(rows: list[dict[str, object]], path: Path) -> None:
    if not rows:
        raise ValueError("No data available for chart generation.")

    width = 1200
    top = 40
    bottom = 30
    left = 200
    right = 40
    bar_h = 36
    gap = 16
    chart_h = len(rows) * (bar_h + gap) - gap
    height = top + chart_h + bottom

    background = (250, 250, 250)
    bar_color = (67, 97, 238)
    axis_color = (80, 80, 80)

    canvas = [[background[0], background[1], background[2]] * width for _ in range(height)]

    max_rating = max(float(row["mean_rating"]) for row in rows)
    scale = (width - left - right) / max_rating if max_rating else 1.0

    # Axis line
    for y in range(top - 10, top + chart_h + 1):
        idx = y
        if 0 <= idx < height:
            offset = (left - 1) * 3
            canvas[idx][offset : offset + 3] = list(axis_color)

    for i, row in enumerate(rows):
        mean_rating = float(row["mean_rating"])
        bar_width = max(1, int(mean_rating * scale))
        y0 = top + i * (bar_h + gap)
        y1 = y0 + bar_h
        x0 = left
        x1 = min(width - right, x0 + bar_width)
        for y in range(y0, y1):
            row_px = canvas[y]
            for x in range(x0, x1):
                p = x * 3
                row_px[p : p + 3] = list(bar_color)

    rgb_rows = [bytes(row) for row in canvas]
    make_png(width, height, rgb_rows, path)


def resolve_input_path() -> Path:
    if INPUT_PATH.exists():
        return INPUT_PATH
    if FALLBACK_INPUT_PATH.exists():
        print(f"Warning: {INPUT_PATH} not found; using fallback {FALLBACK_INPUT_PATH}.")
        return FALLBACK_INPUT_PATH
    raise FileNotFoundError(f"Missing input workbook at {INPUT_PATH}")


def main() -> None:
    rows = read_sheet_rows(resolve_input_path())
    ranked = aggregate_course_ratings(rows)
    write_csv(ranked, OUTPUT_CSV)
    draw_bar_chart(ranked, OUTPUT_PNG)
    print(f"Wrote {OUTPUT_CSV} and {OUTPUT_PNG}")


if __name__ == "__main__":
    main()
