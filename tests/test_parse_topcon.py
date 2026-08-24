import csv
import json
import struct
import tempfile
import unittest
from pathlib import Path

from parse_topcon import (
    TopconFormatError,
    export_dataset,
    parse_history,
    parse_patients,
)


def _write_text_field(
    record: bytearray,
    start: int,
    end: int,
    value: str,
    *,
    encoding: str = "ascii",
) -> None:
    encoded = value.encode(encoding)
    if len(encoded) >= end - start:
        raise ValueError("test field does not leave room for a NUL terminator")
    record[start : start + len(encoded)] = encoded


def _make_history_record() -> bytes:
    record = bytearray(820)
    record[0:4] = b"*IM\0"
    _write_text_field(record, 4, 24, "P001")
    _write_text_field(record, 24, 48, "测试患者", encoding="gb18030")
    _write_text_field(record, 48, 76, "8岁", encoding="gb18030")
    _write_text_field(record, 76, 80, "F")
    _write_text_field(record, 80, 132, "2016-01-02")
    _write_text_field(record, 132, 160, "2024-03-04")
    _write_text_field(record, 160, 188, "2025-06-07")
    _write_text_field(record, 620, 640, "P001")
    _write_text_field(record, 640, 652, "2025-06-07")
    _write_text_field(record, 652, 664, "08:09:10")
    _write_text_field(record, 664, 684, "DC3Capture")
    _write_text_field(record, 684, 756, "#2")
    _write_text_field(record, 756, 772, "IM000042.JPG")
    struct.pack_into("<I", record, 776, 12345)
    struct.pack_into("<I", record, 790, 2)
    record[816:820] = b"END!"
    return bytes(record)


class ParseTopconTests(unittest.TestCase):
    def test_parse_patients_preserves_six_exported_columns(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "patients.TXT"
            path.write_bytes(
                "P001,测试患者,8岁,2016-01-02,Female,2024-03-04\r\n".encode(
                    "gb18030"
                )
            )

            patients = parse_patients(path)

        self.assertEqual(len(patients), 1)
        self.assertEqual(patients[0].patient_id, "P001")
        self.assertEqual(patients[0].patient_name, "测试患者")
        self.assertEqual(patients[0].age_text, "8岁")
        self.assertEqual(patients[0].birth_date, "2016-01-02")
        self.assertEqual(patients[0].sex, "Female")
        self.assertEqual(patients[0].registered_date, "2024-03-04")

    def test_parse_history_extracts_patient_capture_and_filename_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "HIST0010"
            path.write_bytes(_make_history_record())

            records = parse_history(path)

        self.assertEqual(len(records), 1)
        record = records[0]
        self.assertEqual(record.record_index, 0)
        self.assertEqual(record.patient_id, "P001")
        self.assertEqual(record.patient_name, "测试患者")
        self.assertEqual(record.last_capture_date, "2025-06-07")
        self.assertEqual(record.capture_date, "2025-06-07")
        self.assertEqual(record.capture_time, "08:09:10")
        self.assertEqual(record.capture_source, "DC3Capture")
        self.assertEqual(record.image_number_in_exam, "#2")
        self.assertEqual(record.image_filename, "IM000042.JPG")
        self.assertEqual(record.image_sequence, 42)
        self.assertEqual(record.unknown_uint32_at_776, 12345)
        self.assertEqual(record.unknown_flag_at_790, 2)

    def test_parse_history_rejects_non_multiple_record_size(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "HIST0010"
            path.write_bytes(b"*IM\0" + bytes(815))

            with self.assertRaisesRegex(TopconFormatError, "820"):
                parse_history(path)

    def test_export_dataset_writes_csv_and_report(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            data_dir = root / "data"
            output_dir = root / "output"
            data_dir.mkdir()
            (data_dir / "LABEL").write_bytes(b"00010 TOPCON\r\n")
            (data_dir / "patients.TXT").write_bytes(
                "P001,测试患者,8岁,2016-01-02,Female,2024-03-04\r\n".encode(
                    "gb18030"
                )
            )
            (data_dir / "HIST0010").write_bytes(_make_history_record())

            report = export_dataset(data_dir, output_dir)

            with (output_dir / "patients.csv").open(
                "r", encoding="utf-8-sig", newline=""
            ) as handle:
                patient_rows = list(csv.DictReader(handle))
            with (output_dir / "image_index.csv").open(
                "r", encoding="utf-8-sig", newline=""
            ) as handle:
                image_rows = list(csv.DictReader(handle))
            saved_report = json.loads(
                (output_dir / "parse_report.json").read_text(encoding="utf-8")
            )

        self.assertEqual(len(patient_rows), 1)
        self.assertEqual(len(image_rows), 1)
        self.assertEqual(image_rows[0]["image_filename"], "IM000042.JPG")
        self.assertEqual(report["history"]["record_count"], 1)
        self.assertEqual(saved_report["images"]["extensions"], {"JPG": 1})


if __name__ == "__main__":
    unittest.main()
