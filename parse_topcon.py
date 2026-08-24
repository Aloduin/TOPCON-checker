"""Parse the TOPCON patient export and fixed-width HIST0010 image index.

The parser is based on the structure observed in the files copied from the
slit-lamp workstation.  It does not attempt to recover image pixels: HIST0010
contains metadata and image filenames, while the JPG/PNG files live elsewhere.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import struct
import sys
from collections import Counter
from dataclasses import asdict, dataclass, fields
from datetime import date, datetime
from pathlib import Path
from typing import Iterable


RECORD_SIZE = 820
TEXT_ENCODING = "gb18030"
IMAGE_FILENAME_RE = re.compile(r"^IM(?P<sequence>[0-9]{6})\.(?P<ext>JPG|PNG)$", re.I)
IMAGE_SUFFIXES = {".JPG", ".JPEG", ".PNG", ".BMP", ".TIF", ".TIFF"}


class TopconFormatError(ValueError):
    """Raised when an input file does not match the observed TOPCON format."""


@dataclass(frozen=True, slots=True)
class PatientRecord:
    source_row: int
    patient_id: str
    patient_name: str
    age_text: str
    birth_date: str
    sex: str
    registered_date: str


@dataclass(frozen=True, slots=True)
class ImageRecord:
    record_index: int
    patient_id: str
    patient_name: str
    age_or_legacy_birth_date: str
    sex: str
    birth_date: str
    registered_date: str
    last_capture_date: str
    repeated_patient_id: str
    capture_date: str
    capture_time: str
    capture_source: str
    image_number_in_exam: str
    image_filename: str
    image_sequence: int | None
    image_extension: str
    unknown_uint32_at_772: int
    unknown_uint32_at_776: int
    unknown_flag_at_790: int
    unknown_tail_hex: str


def _decode_field(record: bytes, start: int, end: int) -> str:
    raw = record[start:end].split(b"\0", 1)[0].rstrip()
    try:
        return raw.decode(TEXT_ENCODING)
    except UnicodeDecodeError as exc:
        raise TopconFormatError(
            f"字段 {start}:{end} 无法用 {TEXT_ENCODING} 解码"
        ) from exc


def parse_patients(path: str | Path) -> list[PatientRecord]:
    """Parse the six-column, GB18030 TOPCON patient TXT export."""

    input_path = Path(path)
    records: list[PatientRecord] = []
    try:
        handle = input_path.open("r", encoding=TEXT_ENCODING, newline="")
    except UnicodeError as exc:
        raise TopconFormatError(f"患者文件不是有效的 {TEXT_ENCODING} 文本") from exc

    try:
        with handle:
            for source_row, row in enumerate(csv.reader(handle), start=1):
                if not row or all(not value.strip() for value in row):
                    continue
                if len(row) != 6:
                    raise TopconFormatError(
                        f"患者文件第 {source_row} 行有 {len(row)} 列，预期为 6 列"
                    )
                values = [value.strip() for value in row]
                values[0] = values[0].lstrip("\ufeff")
                records.append(PatientRecord(source_row, *values))
    except UnicodeDecodeError as exc:
        raise TopconFormatError(f"患者文件不是有效的 {TEXT_ENCODING} 文本") from exc

    return records


def _parse_image_sequence(filename: str) -> tuple[int | None, str]:
    match = IMAGE_FILENAME_RE.fullmatch(filename)
    if match is None:
        suffix = Path(filename).suffix.lstrip(".").upper()
        return None, suffix
    return int(match.group("sequence")), match.group("ext").upper()


def parse_history(path: str | Path) -> list[ImageRecord]:
    """Parse fixed-width image index entries from a HIST0010 file."""

    input_path = Path(path)
    data = input_path.read_bytes()
    if not data:
        raise TopconFormatError("HIST0010 为空")
    if len(data) % RECORD_SIZE != 0:
        raise TopconFormatError(
            f"HIST0010 大小为 {len(data)} 字节，不是 {RECORD_SIZE} 字节记录的整数倍"
        )

    records: list[ImageRecord] = []
    for record_index, offset in enumerate(range(0, len(data), RECORD_SIZE)):
        raw = data[offset : offset + RECORD_SIZE]
        if raw[0:4] != b"*IM\0":
            raise TopconFormatError(
                f"HIST0010 第 {record_index} 条记录没有 *IM 标记"
            )

        patient_id = _decode_field(raw, 4, 24)
        repeated_patient_id = _decode_field(raw, 620, 640)
        if patient_id != repeated_patient_id:
            raise TopconFormatError(
                f"HIST0010 第 {record_index} 条记录的两个患者号字段不一致"
            )

        image_filename = _decode_field(raw, 756, 772)
        image_sequence, image_extension = _parse_image_sequence(image_filename)
        records.append(
            ImageRecord(
                record_index=record_index,
                patient_id=patient_id,
                patient_name=_decode_field(raw, 24, 48),
                age_or_legacy_birth_date=_decode_field(raw, 48, 76),
                sex=_decode_field(raw, 76, 80),
                birth_date=_decode_field(raw, 80, 132),
                registered_date=_decode_field(raw, 132, 160),
                last_capture_date=_decode_field(raw, 160, 188),
                repeated_patient_id=repeated_patient_id,
                capture_date=_decode_field(raw, 640, 652),
                capture_time=_decode_field(raw, 652, 664),
                capture_source=_decode_field(raw, 664, 684),
                image_number_in_exam=_decode_field(raw, 684, 756),
                image_filename=image_filename,
                image_sequence=image_sequence,
                image_extension=image_extension,
                unknown_uint32_at_772=struct.unpack_from("<I", raw, 772)[0],
                unknown_uint32_at_776=struct.unpack_from("<I", raw, 776)[0],
                unknown_flag_at_790=struct.unpack_from("<I", raw, 790)[0],
                unknown_tail_hex=raw[772:820].hex(),
            )
        )

    return records


def _write_records(path: Path, records: Iterable[object], record_type: type) -> None:
    field_names = [field.name for field in fields(record_type)]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=field_names)
        writer.writeheader()
        for record in records:
            writer.writerow(asdict(record))


def _resolve_input(data_dir: Path, value: str | Path | None, default: str) -> Path:
    if value is None:
        return data_dir / default
    path = Path(value)
    return path if path.is_absolute() else data_dir / path


def _find_patient_file(data_dir: Path, value: str | Path | None) -> Path:
    if value is not None:
        return _resolve_input(data_dir, value, "")
    candidates = sorted(
        path for path in data_dir.iterdir() if path.is_file() and path.suffix.lower() == ".txt"
    )
    if len(candidates) != 1:
        raise TopconFormatError(
            f"data 目录中找到 {len(candidates)} 个 TXT 文件；请用 --patients-file 指定一个"
        )
    return candidates[0]


def _valid_iso_date(value: str) -> bool:
    if not value:
        return True
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def _valid_time(value: str) -> bool:
    try:
        datetime.strptime(value, "%H:%M:%S")
    except ValueError:
        return False
    return True


def export_dataset(
    data_dir: str | Path,
    output_dir: str | Path,
    *,
    patients_file: str | Path | None = None,
    history_file: str | Path | None = None,
    label_file: str | Path | None = None,
) -> dict[str, object]:
    """Parse a copied TOPCON dataset and write CSV/JSON outputs."""

    source_dir = Path(data_dir)
    destination = Path(output_dir)
    if not source_dir.is_dir():
        raise TopconFormatError(f"数据目录不存在：{source_dir}")

    patient_path = _find_patient_file(source_dir, patients_file)
    history_path = _resolve_input(source_dir, history_file, "HIST0010")
    label_path = _resolve_input(source_dir, label_file, "LABEL")
    for path in (patient_path, history_path, label_path):
        if not path.is_file():
            raise TopconFormatError(f"输入文件不存在：{path}")

    try:
        label = label_path.read_text(encoding="ascii").strip()
    except UnicodeDecodeError as exc:
        raise TopconFormatError("LABEL 不是 ASCII 文本") from exc

    patients = parse_patients(patient_path)
    images = parse_history(history_path)
    destination.mkdir(parents=True, exist_ok=True)
    _write_records(destination / "patients.csv", patients, PatientRecord)
    _write_records(destination / "image_index.csv", images, ImageRecord)

    patient_ids = {record.patient_id for record in patients}
    history_patient_ids = {record.patient_id for record in images}
    existing_image_names = {
        path.name.upper()
        for path in source_dir.rglob("*")
        if path.is_file() and path.suffix.upper() in IMAGE_SUFFIXES
    }
    extension_counts = Counter(record.image_extension for record in images)
    source_counts = Counter(record.capture_source for record in images)

    report: dict[str, object] = {
        "format": {
            "label": label,
            "text_encoding": TEXT_ENCODING,
            "history_record_size": RECORD_SIZE,
        },
        "patients": {
            "source_file": patient_path.name,
            "row_count": len(patients),
            "unique_patient_ids": len(patient_ids),
            "duplicate_patient_rows": len(patients) - len(patient_ids),
        },
        "history": {
            "source_file": history_path.name,
            "byte_count": history_path.stat().st_size,
            "record_count": len(images),
            "unique_patient_ids": len(history_patient_ids),
        },
        "images": {
            "unique_filenames": len({record.image_filename for record in images}),
            "valid_im_filenames": sum(
                record.image_sequence is not None for record in images
            ),
            "extensions": dict(sorted(extension_counts.items())),
            "capture_sources": dict(sorted(source_counts.items())),
            "files_found_under_data": sum(
                record.image_filename.upper() in existing_image_names for record in images
            ),
            "files_missing_under_data": sum(
                record.image_filename.upper() not in existing_image_names for record in images
            ),
        },
        "cross_reference": {
            "history_patient_ids_found_in_patient_export": len(
                history_patient_ids & patient_ids
            ),
            "history_patient_ids_missing_from_patient_export": len(
                history_patient_ids - patient_ids
            ),
        },
        "validation": {
            "invalid_birth_dates": sum(
                not _valid_iso_date(record.birth_date) for record in images
            ),
            "invalid_registered_dates": sum(
                not _valid_iso_date(record.registered_date) for record in images
            ),
            "invalid_last_capture_dates": sum(
                not _valid_iso_date(record.last_capture_date) for record in images
            ),
            "invalid_capture_dates": sum(
                not _valid_iso_date(record.capture_date) for record in images
            ),
            "invalid_capture_times": sum(
                not _valid_time(record.capture_time) for record in images
            ),
        },
        "outputs": ["patients.csv", "image_index.csv", "parse_report.json"],
    }
    (destination / "parse_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return report


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="解析 TOPCON 裂隙灯患者导出和 HIST0010 图片索引"
    )
    parser.add_argument(
        "data_dir",
        nargs="?",
        default="data",
        help="包含 LABEL、HIST0010 和患者 TXT 的目录（默认：data）",
    )
    parser.add_argument(
        "--output-dir", default="output", help="CSV/JSON 输出目录（默认：output）"
    )
    parser.add_argument(
        "--patients-file",
        help="患者 TXT 文件名；相对路径按 data_dir 解析",
    )
    parser.add_argument(
        "--history-file",
        default="HIST0010",
        help="历史索引文件名（默认：HIST0010）",
    )
    parser.add_argument(
        "--label-file", default="LABEL", help="标签文件名（默认：LABEL）"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    try:
        report = export_dataset(
            args.data_dir,
            args.output_dir,
            patients_file=args.patients_file,
            history_file=args.history_file,
            label_file=args.label_file,
        )
    except (OSError, TopconFormatError) as exc:
        print(f"解析失败：{exc}", file=sys.stderr)
        return 1

    print(
        json.dumps(
            {
                "patient_rows": report["patients"]["row_count"],
                "image_records": report["history"]["record_count"],
                "image_files_found": report["images"]["files_found_under_data"],
                "output_dir": str(Path(args.output_dir)),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
