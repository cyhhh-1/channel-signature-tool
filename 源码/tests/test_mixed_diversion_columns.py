import csv
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


SOURCE_PATH = (
    Path(__file__).resolve().parents[1]
    / "通道生效签名短链工具"
    / "通道生效签名短链工具源码.py"
)
SPEC = importlib.util.spec_from_file_location("channel_importer", SOURCE_PATH)
assert SPEC is not None and SPEC.loader is not None
channel_importer = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = channel_importer
SPEC.loader.exec_module(channel_importer)


class MixedDiversionColumnsTest(unittest.TestCase):
    def test_separate_link_and_number_columns_drive_mixed_routing(self) -> None:
        rows = [
            ["签名", "手机号码", "座机号码", "链接", "短信模板"],
            ["【无引流】", "", "", "", "模板A"],
            ["【占位值】", "无", "暂无", "—", "模板B"],
            ["【只有手机】", "13800138000", "", "", "模板C"],
            ["【完整引流】", "13800138001", "400-888-8888", "https://example.com", "模板D"],
        ]
        with tempfile.TemporaryDirectory() as temporary_dir:
            source = Path(temporary_dir) / "mixed.csv"
            with source.open("w", encoding="utf-8-sig", newline="") as stream:
                csv.writer(stream).writerows(rows)

            shortlinks, signatures, errors, shortlink_duplicates, signature_duplicates = (
                channel_importer.read_mixed_records(source)
            )

        self.assertEqual(
            signatures,
            [
                ("【无引流】",),
                ("【占位值】",),
                ("【只有手机】",),
                ("【完整引流】",),
            ],
        )
        self.assertEqual(
            shortlinks,
            [
                ("【只有手机】", "", "13800138000"),
                ("【完整引流】", "https://example.com", "13800138001,400-888-8888"),
            ],
        )
        self.assertEqual(errors, [])
        self.assertEqual(shortlink_duplicates, 0)
        self.assertEqual(signature_duplicates, 0)

    def test_vertical_bar_splits_all_diversion_combinations(self) -> None:
        rows = [
            ["签名", "引流"],
            ["【号码竖线】", "13800138000|400-888-8888"],
            ["【号链竖线】", "13800138001|https://one.example/path"],
            ["【双链竖线】", "https://one.example/path|https://two.example/path"],
        ]
        with tempfile.TemporaryDirectory() as temporary_dir:
            source = Path(temporary_dir) / "vertical_bar.csv"
            with source.open("w", encoding="utf-8-sig", newline="") as stream:
                csv.writer(stream).writerows(rows)

            shortlinks, signatures, errors, _, _ = channel_importer.read_mixed_records(source)

        self.assertEqual(
            shortlinks,
            [
                ("【号码竖线】", "", "13800138000,400-888-8888"),
                ("【号链竖线】", "https://one.example/path", "13800138001"),
                ("【双链竖线】", "https://one.example/path", ""),
                ("【双链竖线】", "https://two.example/path", ""),
            ],
        )
        self.assertEqual(
            signatures,
            [("【号码竖线】",), ("【号链竖线】",), ("【双链竖线】",)],
        )
        self.assertEqual(errors, [])

    def test_signature_output_deduplicates_all_rows_without_dropping_shortlinks(self) -> None:
        rows = [
            ["签名", "号码", "链接"],
            ["【纯签名】", "", ""],
            ["【号码签名】", "13800138000", ""],
            ["【链接签名】", "", "https://example.com"],
            ["【号码签名】", "400-888-8888", ""],
        ]
        with tempfile.TemporaryDirectory() as temporary_dir:
            source = Path(temporary_dir) / "mixed.csv"
            with source.open("w", encoding="utf-8-sig", newline="") as stream:
                csv.writer(stream).writerows(rows)

            shortlinks, signatures, errors, _, signature_duplicates = (
                channel_importer.read_mixed_records(source)
            )

        self.assertEqual(
            signatures,
            [("【纯签名】",), ("【号码签名】",), ("【链接签名】",)],
        )
        self.assertEqual(len(shortlinks), 3)
        self.assertEqual(errors, [])
        self.assertEqual(signature_duplicates, 1)


if __name__ == "__main__":
    unittest.main()
