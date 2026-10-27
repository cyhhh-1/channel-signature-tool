import contextlib
import importlib.util
import io
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook


SOURCE_PATH = (
    Path(__file__).resolve().parents[1]
    / "通道生效签名短链工具"
    / "通道生效签名短链工具源码.py"
)
SPEC = importlib.util.spec_from_file_location("channel_importer_cli_regression", SOURCE_PATH)
assert SPEC is not None and SPEC.loader is not None
channel_importer = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = channel_importer
SPEC.loader.exec_module(channel_importer)


class OriginalModesCliRegressionTest(unittest.TestCase):
    def test_signature_mode_accepts_launchers_empty_table_action(self) -> None:
        """原五种模式不能被新增拆分/合并参数拦在业务处理之前。"""
        with tempfile.TemporaryDirectory() as temporary_dir:
            root = Path(temporary_dir)
            source = root / "签名.txt"
            source.write_text("【回归测试签名】\n", encoding="utf-8")

            template = root / channel_importer.SIGNATURE_TEMPLATE
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(channel_importer.SIGNATURE_OUTPUT_HEADERS)
            workbook.save(template)
            workbook.close()

            stdout = io.StringIO()
            stderr = io.StringIO()
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                exit_code = channel_importer.run_cli(
                    [
                        "--mode",
                        channel_importer.MODE_SIGNATURE,
                        "--source",
                        str(source),
                        "--channel",
                        "1160710",
                        "--start-sub-port",
                        "1000",
                        "--operator",
                        "电信",
                        "--business-type",
                        "行业",
                        "--important-support",
                        "否",
                        "--template-dir",
                        str(root),
                        "--table-action",
                        "",
                    ]
                )

            self.assertEqual(exit_code, 0, stderr.getvalue())
            self.assertNotIn("invalid choice", stderr.getvalue())

    def test_table_cli_emits_structured_progress_without_affecting_json_result(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            root = Path(temporary_dir)
            for index in (1, 2):
                workbook = Workbook()
                worksheet = workbook.active
                worksheet.append(["编号", "名称"])
                worksheet.append([index, f"客户{index}"])
                workbook.save(root / f"{index:03d}.xlsx")
                workbook.close()

            stdout = io.StringIO()
            stderr = io.StringIO()
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                exit_code = channel_importer.run_cli(
                    [
                        "--mode",
                        channel_importer.MODE_TABLE_TOOLS,
                        "--source",
                        str(root),
                        "--operator",
                        "电信",
                        "--important-support",
                        "否",
                        "--table-action",
                        "merge",
                    ]
                )

            self.assertEqual(exit_code, 0, stderr.getvalue())
            self.assertIn("@@TABLE_PROGRESS\tmerge\t", stderr.getvalue())
            self.assertIn('"generated_rows": 2', stdout.getvalue())


if __name__ == "__main__":
    unittest.main()
