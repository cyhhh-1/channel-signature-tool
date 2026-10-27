import json
import re
import unittest
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / "源码" / "通道生效签名短链工具"


class VersionMetadataTests(unittest.TestCase):
    def test_version_json_is_valid_single_source(self) -> None:
        data = json.loads((ROOT / "版本.json").read_text(encoding="utf-8"))
        self.assertRegex(data["productVersion"], r"^\d+\.\d+\.\d+$")
        self.assertRegex(data["releaseDate"], r"^\d{4}-\d{2}-\d{2}$")
        date.fromisoformat(data["releaseDate"])

    def test_launcher_reads_versions_from_assembly_metadata(self) -> None:
        launcher = (TOOL / "便携启动器源码.cs").read_text(encoding="utf-8")
        self.assertIn('ReadMetadata("ProductVersion")', launcher)
        self.assertIn('ReadMetadata("RuntimeVersion")', launcher)
        self.assertIn('ReadMetadata("MappingSeedVersion")', launcher)
        self.assertNotRegex(launcher, r"private const string RuntimeVersion\s*=")
        self.assertNotRegex(launcher, r"private const string MappingSeedVersion\s*=")
        self.assertIn('"V" + AppInfo.ProductVersion', launcher)

    def test_build_script_derives_versions_and_passes_them_to_nsis(self) -> None:
        build = (ROOT / "生成安装包.ps1").read_text(encoding="utf-8")
        self.assertIn("版本.json", build)
        self.assertIn("Get-PayloadFingerprint", build)
        self.assertIn("AssemblyMetadata", build)
        self.assertIn("/DAPPVERSION=", build)
        self.assertIn("/DOUTFILE=", build)
        self.assertIn("构建结果.txt", build)
        self.assertNotIn("20261008", build)

    def test_installer_uses_build_supplied_version(self) -> None:
        nsi = (ROOT / "安装程序.nsi").read_text(encoding="utf-8")
        self.assertIn('OutFile "${OUTFILE}"', nsi)
        self.assertIn('VIProductVersion "${APPVERSION}.0"', nsi)
        self.assertIn('"DisplayVersion" "${APPVERSION}"', nsi)
        self.assertNotIn("20261008", nsi)


if __name__ == "__main__":
    unittest.main()
