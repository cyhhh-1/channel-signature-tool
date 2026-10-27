import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[2]
SOURCE_PATH = ROOT / "源码" / "通道生效签名短链工具" / "通道生效签名短链工具源码.py"
SPEC = importlib.util.spec_from_file_location("channel_importer_output_relocation", SOURCE_PATH)
channel_importer = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules[SPEC.name] = channel_importer
SPEC.loader.exec_module(channel_importer)


class OutputRelocationTests(unittest.TestCase):
    def test_moves_all_unique_result_files_and_updates_duplicate_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            generated = root / "generated"
            destination = root / "chosen"
            generated.mkdir()
            first = generated / "短链结果.xlsx"
            second = generated / "错误明细.xlsx"
            first.write_bytes(b"first")
            second.write_bytes(b"second")
            summary = channel_importer.ProcessSummary(
                output_path=str(first),
                error_path=str(second),
                generated_rows=1,
                duplicate_rows=0,
                error_rows=1,
                shortlink_output_path=str(first),
                output_paths=(str(first), str(second)),
            )

            relocated = channel_importer.relocate_summary_outputs(
                summary, destination, channel_importer.MODE_MIXED
            )

            result_dir = Path(relocated.output_path).parent
            self.assertEqual(result_dir.parent, destination)
            self.assertRegex(result_dir.name, r"^生成结果_混合表自动分流_\d{4}-\d{2}时\d{2}分$")
            self.assertEqual(relocated.output_path, relocated.shortlink_output_path)
            self.assertEqual(Path(relocated.error_path).parent, result_dir)
            self.assertEqual(Path(relocated.output_dir), result_dir)
            self.assertFalse(first.exists())
            self.assertFalse(second.exists())
            self.assertTrue(Path(relocated.output_path).is_file())
            self.assertTrue(Path(relocated.error_path).is_file())

    def test_existing_name_is_not_overwritten(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            generated = root / "generated"
            destination = root / "chosen"
            generated.mkdir()
            destination.mkdir()
            source = generated / "结果.xlsx"
            source.write_bytes(b"new")
            (destination / "结果.xlsx").write_bytes(b"old")
            summary = channel_importer.ProcessSummary(
                output_path=str(source),
                error_path=None,
                generated_rows=1,
                duplicate_rows=0,
                error_rows=0,
            )

            relocated = channel_importer.relocate_summary_outputs(
                summary, destination, channel_importer.MODE_SHORTLINK
            )

            self.assertEqual((destination / "结果.xlsx").read_bytes(), b"old")
            self.assertEqual(Path(relocated.output_path).name, "结果(1).xlsx")
            self.assertEqual(Path(relocated.output_path).read_bytes(), b"new")

    def test_multi_output_mode_removes_empty_staging_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            staging = root / "生成结果"
            destination = root / "desktop"
            staging.mkdir()
            first = staging / "已分配.xlsx"
            second = staging / "更新总表.xlsx"
            first.write_bytes(b"first")
            second.write_bytes(b"second")
            summary = channel_importer.ProcessSummary(
                output_path=str(first),
                secondary_output_path=str(second),
                error_path=None,
                generated_rows=1,
                duplicate_rows=0,
                error_rows=0,
            )

            relocated = channel_importer.relocate_summary_outputs(
                summary, destination, channel_importer.MODE_PORT_ALLOCATION
            )

            result_dir = Path(relocated.output_dir)
            self.assertRegex(result_dir.name, r"^生成结果_子端口分配_\d{4}-\d{2}时\d{2}分$")
            self.assertEqual({path.name for path in result_dir.iterdir()}, {"已分配.xlsx", "更新总表.xlsx"})
            self.assertFalse(staging.exists())

    def test_all_grouped_modes_use_the_selected_function_name(self) -> None:
        expected = {
            channel_importer.MODE_SELF_MOBILE: "自家移动",
            channel_importer.MODE_SELF_TELECOM: "自家电信",
            channel_importer.MODE_PORT_ALLOCATION: "子端口分配",
            channel_importer.MODE_MIXED: "混合表自动分流",
        }
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for index, (mode, label) in enumerate(expected.items()):
                source = root / f"结果{index}.xlsx"
                source.write_bytes(b"result")
                summary = channel_importer.ProcessSummary(
                    output_path=str(source),
                    error_path=None,
                    generated_rows=1,
                    duplicate_rows=0,
                    error_rows=0,
                )

                relocated = channel_importer.relocate_summary_outputs(summary, root / "chosen", mode)

                self.assertRegex(
                    Path(relocated.output_dir).name,
                    rf"^生成结果_{label}_\d{{4}}-\d{{2}}时\d{{2}}分(?:\(\d+\))?$",
                )


if __name__ == "__main__":
    unittest.main()
