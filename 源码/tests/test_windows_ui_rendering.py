from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / "源码" / "通道生效签名短链工具"


class WindowsUiRenderingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.launcher = (TOOL / "便携启动器源码.cs").read_text(encoding="utf-8")
        self.mapping = (TOOL / "主码号管理源码.cs").read_text(encoding="utf-8")

    def test_app_declares_per_monitor_dpi_awareness(self) -> None:
        manifest = (TOOL / "应用清单.manifest").read_text(encoding="utf-8")
        build = (ROOT / "生成安装包.ps1").read_text(encoding="utf-8")
        launcher = self.launcher
        self.assertIn("PerMonitorV2", manifest)
        self.assertIn("/win32manifest:$manifest", build)
        self.assertLess(
            launcher.index("EnableHighDpiRendering();"),
            launcher.index("Application.EnableVisualStyles();"),
        )

    def test_every_window_uses_the_application_icon(self) -> None:
        launcher = self.launcher
        mapping = self.mapping
        self.assertIn("ApplyWindowIdentity(this);", launcher)
        self.assertEqual(mapping.count("ApplyWindowIdentity(this);"), 2)

    def test_drop_zone_text_is_clear_and_consistent(self) -> None:
        launcher = self.launcher
        self.assertIn('modernSourceDropTitle.Font = UiFonts.Chinese(10F, FontStyle.Bold);', launcher)
        self.assertIn('modernMasterDropTitle.Font = UiFonts.Chinese(10F, FontStyle.Bold);', launcher)
        self.assertIn('sourceLabel.Font = UiFonts.Chinese(11F, FontStyle.Bold);', launcher)
        self.assertIn('masterLabel.Font = UiFonts.Chinese(11F, FontStyle.Bold);', launcher)

    def test_shadcn_monochrome_design_system_is_shared_across_windows(self) -> None:
        self.assertIn("internal static class UiTheme", self.launcher)
        self.assertIn("Segoe UI", self.launcher)
        self.assertIn("Microsoft YaHei UI", self.launcher)
        self.assertIn("Color.FromArgb(23, 23, 23)", self.launcher)
        self.assertIn('Label brand = MakeLabel("工作台"', self.launcher)
        self.assertIn("new SurfacePanel", self.launcher)
        self.assertIn("new SurfacePanel", self.mapping)
        self.assertIn("UiTheme.StyleTextBox", self.mapping)
        self.assertIn("return Chinese(points, FontStyle.Regular);", self.launcher)
        self.assertIn("grid.DefaultCellStyle.Font = UiFonts.Chinese(9.5F, FontStyle.Bold);", self.mapping)
        self.assertIn("sourceBox.Font = UiFonts.Chinese(10F);", self.launcher)

    def test_main_workspace_keeps_all_eight_business_modes(self) -> None:
        for mode in (
            "self_mobile", "self_telecom", "port_allocation", "shortlink",
            "signature", "mixed", "table_tools", "signature_consistency",
        ):
            self.assertIn(f'AddNavigation(modernSidebar, "{mode}"', self.launcher)

    def test_workspace_is_adaptive_and_content_driven(self) -> None:
        launcher = self.launcher
        self.assertIn("Screen.PrimaryScreen.WorkingArea", launcher)
        self.assertIn("workingArea.Width * 0.86", launcher)
        self.assertIn("workingArea.Height * 0.84", launcher)
        self.assertIn("Math.Min(1240", launcher)
        self.assertIn("modernFooterCard.SetBounds(leftX, actionTop, leftWidth, 70)", launcher)
        self.assertIn("modernContent.AutoScrollMinSize", launcher)

    def test_shared_import_pages_use_the_same_dropzone_copy(self) -> None:
        launcher = self.launcher
        self.assertIn(
            'SelectedMode == "shortlink" || SelectedMode == "signature" || SelectedMode == "mixed"',
            launcher,
        )
        self.assertIn('sharedImport ? "拖入待处理文件"', launcher)

    def test_forms_use_styled_antd_inputs_without_legacy_focus_rectangle(self) -> None:
        launcher = self.launcher
        self.assertIn("private readonly AntdUI.Input channelBox", launcher)
        self.assertIn("private readonly AntdUI.Select operatorBox", launcher)
        self.assertIn("UiTheme.StyleInput", launcher)
        self.assertIn("UiTheme.StyleSelect", launcher)
        self.assertNotIn("ControlPaint.DrawFocusRectangle", launcher)
        self.assertNotIn("if (Focused)", launcher)

    def test_window_and_workspace_use_a_clean_white_background(self) -> None:
        self.assertIn("internal static readonly Color Canvas = Color.White;", self.launcher)
        self.assertIn("captionColorAttribute = 35", self.launcher)
        self.assertIn("int white = 0x00FFFFFF", self.launcher)

    def test_sidebar_is_roomier_and_mapping_manager_is_not_bottom_pinned(self) -> None:
        self.assertIn("internal const int SidebarWidth = 250", self.launcher)
        self.assertIn('button.Font = UiFonts.Chinese(12F, FontStyle.Bold);', self.launcher)
        self.assertIn('CreateNavigationButton("主码号管理", 540', self.launcher)
        self.assertNotIn('Convert.ToString(nav.Tag) == "mapping"', self.launcher)
        self.assertIn("e.Graphics.Clear(UiTheme.Surface);", self.launcher)
        self.assertIn("if (Selected || pressed || hovered)", self.launcher)

    def test_sidebar_uses_text_only_navigation_and_bottom_security_badge(self) -> None:
        self.assertIn('new Rectangle(18, pressed ? 1 : 0, Width - 30, Height)', self.launcher)
        self.assertNotIn('UiIconPainter.Draw(e.Graphics, IconKey, new Rectangle(14', self.launcher)
        self.assertIn('modernSidebar.Controls.Add(modernLocalBadge);', self.launcher)
        self.assertIn('modernLocalBadge.Top = Math.Max(630', self.launcher)
        self.assertNotIn('MakeLabel("本地数据处理"', self.launcher)
        self.assertNotIn('MakeLabel("资料仅保存在本机"', self.launcher)

    def test_brand_and_workspace_dividers_match_the_reference_structure(self) -> None:
        self.assertIn('brand.Font = UiFonts.Chinese(11F, FontStyle.Bold);', self.launcher)
        self.assertIn('brand.AutoSize = true;', self.launcher)
        self.assertIn('brandEnglish.AutoSize = true;', self.launcher)
        self.assertIn('internal sealed class RoundedSidebarPanel', self.launcher)
        self.assertIn('internal sealed class RoundedMainHeaderPanel', self.launcher)
        self.assertIn('MakeLabel("Dashboard"', self.launcher)
        self.assertIn('new VersionBadge { Text = "V" + AppInfo.ProductVersion }', self.launcher)
        self.assertIn('versionBadge.SetBounds(238 - versionBadgeWidth, 24, versionBadgeWidth, 24);', self.launcher)
        self.assertIn('modernBreadcrumb.Left = originX;', self.launcher)
        self.assertIn('internal const int NavigationGapWidth = 12;', self.launcher)
        self.assertIn('modernNavigationGap = new Panel { BackColor = UiTheme.Surface };', self.launcher)
        self.assertIn('UiTheme.SidebarWidth + UiTheme.NavigationGapWidth', self.launcher)
        self.assertNotIn('BackColor = UiTheme.Text, Dock = DockStyle.Left', self.launcher)
        self.assertNotIn('BackColor = UiTheme.Text, Dock = DockStyle.Bottom', self.launcher)

    def test_workspace_removes_nonessential_helper_copy(self) -> None:
        self.assertNotIn('MakeLabel("选择或拖入文件，程序不会修改原始数据。"', self.launcher)
        self.assertNotIn('MakeLabel("根据当前任务填写必要信息。"', self.launcher)
        self.assertIn('modernEyebrow.Visible = false;', self.launcher)
        self.assertIn('modeHint.Visible = false;', self.launcher)
        self.assertIn('modernActionHint.Visible = false;', self.launcher)
        self.assertIn('modernTopbar.SetBounds(0, 0, mainWidth, 72);', self.launcher)

    def test_results_can_be_saved_to_a_user_selected_directory(self) -> None:
        self.assertIn('Label actionTitle = MakeLabel("另存位置"', self.launcher)
        self.assertIn('outputDirectoryBox.PlaceholderText = "默认：原文件所在目录";', self.launcher)
        self.assertIn('browseOutputButton.Text = "选择位置";', self.launcher)
        self.assertIn('private void BrowseOutputDirectory', self.launcher)
        self.assertIn('args.Add("--output-dir");', self.launcher)

    def test_drop_zones_show_a_visible_dragging_state(self) -> None:
        self.assertIn("internal bool DragActive", self.launcher)
        self.assertIn("DragOver += InputDragOver;", self.launcher)
        self.assertIn("DragLeave += InputDragLeave;", self.launcher)
        self.assertIn("SetDropZoneDragState", self.launcher)

    def test_mapping_manager_can_export_the_current_workbook(self) -> None:
        self.assertIn('AddButton("导出表格"', self.mapping)
        self.assertIn("private void ExportWorkbook()", self.mapping)
        self.assertIn("File.Copy(Program.MappingFile, dialog.FileName, true);", self.mapping)

    def test_usage_guide_is_directly_above_the_local_security_badge(self) -> None:
        self.assertIn('modernGuideButton = CreateNavigationButton("使用说明", 580', self.launcher)
        self.assertIn(
            "modernGuideButton.Top = modernLocalBadge.Top - modernGuideButton.Height - 12;",
            self.launcher,
        )
        self.assertIn("internal sealed class UsageGuideForm", self.launcher)
        self.assertIn("using (UsageGuideForm form = new UsageGuideForm())", self.launcher)

    def test_black_application_icon_is_used_inside_and_outside_the_window(self) -> None:
        launcher = self.launcher
        png = ROOT / "发布资源" / "应用图标.png"
        ico = ROOT / "发布资源" / "应用图标.ico"
        self.assertIn('Image.FromFile(Path.Combine(Program.ResourcesDirectory, "应用图标.png"))', launcher)
        self.assertGreater(png.stat().st_size, 100_000)
        self.assertGreater(ico.stat().st_size, 50_000)
        self.assertEqual(ico.read_bytes()[:4], b"\x00\x00\x01\x00")

    def test_tables_and_modals_have_explicit_interaction_states(self) -> None:
        self.assertIn("grid.DefaultCellStyle.SelectionBackColor = UiTheme.PrimarySoft", self.mapping)
        self.assertIn("button.BackColor = UiTheme.DangerSoft", self.mapping)
        self.assertIn("save.FlatAppearance.BorderSize = 0", self.mapping)


if __name__ == "__main__":
    unittest.main()
