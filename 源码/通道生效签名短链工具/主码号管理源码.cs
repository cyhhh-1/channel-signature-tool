using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using System.Windows.Forms;

internal sealed class MappingForm : Form
{
    private readonly TextBox search = new TextBox();
    private readonly DataGridView grid = new DataGridView();
    private readonly Label count = new Label();
    private readonly Panel toolbarCard = new SurfacePanel();
    private readonly Panel tableCard = new SurfacePanel();
    private List<Dictionary<string, object>> records = new List<Dictionary<string, object>>();

    internal MappingForm()
        : this(true)
    {
    }

    internal MappingForm(bool autoLoad)
    {
        Text = "主码号管理";
        Program.ApplyWindowIdentity(this);
        ClientSize = new Size(1040, 680);
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(1000, 620);
        Font = UiFonts.Chinese(10F);
        BackColor = UiTheme.Canvas;

        Panel header = new Panel { BackColor = UiTheme.Surface, Dock = DockStyle.Top, Height = 92 };
        Controls.Add(header);
        Label title = new Label { Text = "主码号管理", Font = UiFonts.Chinese(17F, FontStyle.Bold),
            ForeColor = UiTheme.Text, AutoSize = true, Location = new Point(30, 20) };
        header.Controls.Add(title);
        Label note = new Label { Text = "维护通道、主码号与省份映射，所有修改仅保存在本机。",
            ForeColor = UiTheme.TextSecondary, AutoSize = true, Location = new Point(32, 57) };
        header.Controls.Add(note);
        Label local = new Label { Text = "本地资料库", Font = UiFonts.Chinese(9.5F, FontStyle.Bold),
            ForeColor = UiTheme.Success, BackColor = UiTheme.SuccessSoft, TextAlign = ContentAlignment.MiddleCenter };
        local.SetBounds(875, 28, 132, 32);
        local.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        header.Controls.Add(local);
        header.Controls.Add(new Panel { BackColor = UiTheme.Border, Dock = DockStyle.Bottom, Height = 1 });

        toolbarCard.SetBounds(28, 116, 984, 94);
        toolbarCard.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        Controls.Add(toolbarCard);
        Label searchLabel = new Label { Text = "搜索资料", AutoSize = true, Font = UiFonts.Chinese(10F, FontStyle.Bold),
            ForeColor = UiTheme.TextSecondary, Location = new Point(20, 16) };
        toolbarCard.Controls.Add(searchLabel);
        search.SetBounds(20, 43, 324, 34);
        UiTheme.StyleTextBox(search);
        search.TextChanged += delegate { ShowFiltered(); };
        toolbarCard.Controls.Add(search);
        AddButton("添加记录", 364, 39, 92, true, false, delegate { EditRecord(null); });
        AddButton("修改", 466, 39, 68, false, false, delegate { EditSelected(); });
        AddButton("删除", 544, 39, 68, false, true, delegate { DeleteSelected(); });
        AddButton("导入表格", 622, 39, 88, false, false, ImportWorkbook);
        AddButton("导出表格", 720, 39, 88, false, false, ExportWorkbook);
        AddButton("刷新", 818, 39, 68, false, false, delegate { Reload(); });

        tableCard.SetBounds(28, 226, 984, 410);
        tableCard.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        Controls.Add(tableCard);
        Label tableTitle = new Label { Text = "通道资料", Font = UiFonts.Chinese(12F, FontStyle.Bold),
            ForeColor = UiTheme.Text, AutoSize = true, Location = new Point(20, 17) };
        tableCard.Controls.Add(tableTitle);
        grid.SetBounds(1, 54, 982, 316);
        grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.None;
        grid.GridColor = UiTheme.Border;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.MultiSelect = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.RowHeadersVisible = false;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersHeight = 40;
        grid.RowTemplate.Height = 38;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.ColumnHeadersDefaultCellStyle.BackColor = UiTheme.SurfaceMuted;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = UiTheme.TextSecondary;
        grid.ColumnHeadersDefaultCellStyle.Font = UiFonts.Chinese(9.5F, FontStyle.Bold);
        grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 0, 0, 0);
        grid.DefaultCellStyle.BackColor = UiTheme.Surface;
        grid.DefaultCellStyle.ForeColor = UiTheme.Text;
        grid.DefaultCellStyle.Font = UiFonts.Chinese(9.5F, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = UiTheme.PrimarySoft;
        grid.DefaultCellStyle.SelectionForeColor = UiTheme.Primary;
        grid.DefaultCellStyle.Padding = new Padding(8, 0, 0, 0);
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(252, 253, 255);
        grid.Columns.Add("row", "行号");
        grid.Columns[0].Visible = false;
        grid.Columns.Add("main", "主码号");
        grid.Columns.Add("channel", "通道编号");
        grid.Columns.Add("name", "通道名称");
        grid.Columns.Add("province", "省份");
        grid.Columns.Add("status", "状态");
        grid.Columns[1].FillWeight = 16;
        grid.Columns[2].FillWeight = 18;
        grid.Columns[3].FillWeight = 40;
        grid.Columns[4].FillWeight = 16;
        grid.Columns[5].FillWeight = 10;
        grid.CellDoubleClick += delegate { EditSelected(); };
        tableCard.Controls.Add(grid);
        count.SetBounds(20, 374, 900, 26);
        count.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
        count.ForeColor = UiTheme.TextSecondary;
        count.Font = UiFonts.Chinese(9.5F);
        tableCard.Controls.Add(count);
        if (autoLoad) Shown += delegate { Reload(); };
    }

    private void AddButton(string label, int x, int y, int width, bool primary, bool danger, Action action)
    {
        AntdUI.Button button = new AntdUI.Button();
        button.Text = label;
        button.SetBounds(x, y, width, 38);
        button.Radius = 8;
        button.Font = UiFonts.Chinese(9.5F, primary ? FontStyle.Bold : FontStyle.Regular);
        if (primary)
        {
            button.Type = AntdUI.TTypeMini.Primary;
            button.BackColor = UiTheme.Primary;
            button.ForeColor = Color.White;
        }
        else if (danger)
        {
            button.BackColor = UiTheme.DangerSoft;
            button.ForeColor = UiTheme.Danger;
        }
        else
        {
            button.BackColor = UiTheme.Surface;
            button.ForeColor = UiTheme.TextSecondary;
        }
        button.Click += delegate { action(); };
        toolbarCard.Controls.Add(button);
    }

    private static string Get(Dictionary<string, object> record, string key)
    {
        object value;
        return record.TryGetValue(key, out value) && value != null ? value.ToString() : "";
    }

    private PythonResult Call(string action, int row, string channel, string name, string main, string province)
    {
        return CallFile(Program.MappingFile, action, row, channel, name, main, province);
    }

    private PythonResult CallFile(string file, string action, int row, string channel, string name, string main, string province)
    {
        return Program.ExecutePython(new string[] {
            "--manage-mapping", "--mapping-file", file,
            "--mapping-action", action, "--mapping-row", row.ToString(),
            "--mapping-channel", channel, "--mapping-name", name,
            "--mapping-main", main, "--mapping-province", province
        });
    }

    private void ImportWorkbook()
    {
        using (OpenFileDialog dialog = new OpenFileDialog())
        {
            dialog.Title = "导入主码号表";
            dialog.Filter = "Excel 文件|*.xlsx";
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                PythonResult check = CallFile(dialog.FileName, "list", 0, "", "", "", "");
                if (check.ExitCode != 0) throw new InvalidOperationException(check.Error);
                object parsed = new JavaScriptSerializer().DeserializeObject(check.Output.Trim());
                int importedCount = ((object[])parsed).Length;
                if (importedCount == 0) throw new InvalidOperationException("导入表格没有可用记录。");
                if (MessageBox.Show(this, "将用所选表格的 " + importedCount + " 条记录替换当前主码号资料。\r\n程序会自动保留一份更新前备份。",
                    "确认导入", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                string folder = Path.GetDirectoryName(Program.MappingFile);
                string backup = Path.Combine(folder, "通道编号及对应主码号_导入前备份_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                File.Copy(Program.MappingFile, backup, true);
                File.Copy(dialog.FileName, Program.MappingFile, true);
                Reload();
                MessageBox.Show(this, "已导入 " + importedCount + " 条记录。\r\n备份文件：\r\n" + backup,
                    "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法导入主码号表：\r\n" + ex.Message, "导入失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    private void ExportWorkbook()
    {
        using (SaveFileDialog dialog = new SaveFileDialog())
        {
            dialog.Title = "导出主码号表";
            dialog.Filter = "Excel 文件|*.xlsx";
            dialog.FileName = "通道编号及对应主码号_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            dialog.OverwritePrompt = true;
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                File.Copy(Program.MappingFile, dialog.FileName, true);
                MessageBox.Show(this, "主码号表已导出到：\r\n" + dialog.FileName, "导出完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法导出主码号表：\r\n" + ex.Message, "导出失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    private void Reload()
    {
        try
        {
            PythonResult result = Call("list", 0, "", "", "", "");
            if (result.ExitCode != 0) throw new InvalidOperationException(result.Error);
            object parsed = new JavaScriptSerializer().DeserializeObject(result.Output.Trim());
            records = ((object[])parsed).Cast<Dictionary<string, object>>().ToList();
            ShowFiltered();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "无法读取主码号资料：\r\n" + ex.Message, "读取失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ShowFiltered()
    {
        string query = search.Text.Trim();
        grid.Rows.Clear();
        foreach (Dictionary<string, object> item in records)
        {
            if (query.Length > 0 && !new[] { "main", "channel", "name", "province" }
                .Any(key => Get(item, key).IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)) continue;
            string main = Get(item, "main");
            bool valid = main.Length == 8 && main.All(char.IsDigit);
            int displayed = grid.Rows.Add(Get(item, "row"), main, Get(item, "channel"),
                Get(item, "name"), Get(item, "province"), valid ? "有效" : "主码号非8位");
            if (!valid) grid.Rows[displayed].DefaultCellStyle.ForeColor = Color.Firebrick;
        }
        int invalidCount = records.Count(item => Get(item, "main").Length != 8 || !Get(item, "main").All(char.IsDigit));
        count.Text = "显示 " + grid.Rows.Count + " 条 / 共 " + records.Count + " 条通道记录" +
            (invalidCount > 0 ? "；其中 " + invalidCount + " 条主码号不是8位数字，暂不参与业务匹配" : "");
    }

    private Dictionary<string, object> SelectedRecord()
    {
        if (grid.CurrentRow == null) return null;
        string row = Convert.ToString(grid.CurrentRow.Cells[0].Value);
        return records.FirstOrDefault(item => Get(item, "row") == row);
    }

    private void EditSelected()
    {
        Dictionary<string, object> selected = SelectedRecord();
        if (selected == null) return;
        EditRecord(selected);
    }

    private void EditRecord(Dictionary<string, object> selected)
    {
        using (MappingEditForm editor = new MappingEditForm(selected))
        {
            if (editor.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                int row = selected == null ? 0 : Convert.ToInt32(Get(selected, "row"));
                PythonResult result = Call(selected == null ? "add" : "update", row,
                    editor.Channel, editor.ChannelName, editor.MainCode, editor.Province);
                if (result.ExitCode != 0) throw new InvalidOperationException(result.Error);
                Reload();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    private void DeleteSelected()
    {
        Dictionary<string, object> selected = SelectedRecord();
        if (selected == null) return;
        if (MessageBox.Show(this, "删除主码号 " + Get(selected, "main") + " 的通道“" + Get(selected, "name") + "”？",
            "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        try
        {
            PythonResult result = Call("delete", Convert.ToInt32(Get(selected, "row")), "", "", "", "");
            if (result.ExitCode != 0) throw new InvalidOperationException(result.Error);
            Reload();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

internal sealed class MappingEditForm : Form
{
    private readonly TextBox main = new TextBox();
    private readonly TextBox channel = new TextBox();
    private readonly TextBox name = new TextBox();
    private readonly TextBox province = new TextBox();
    private readonly Panel fieldCard = new SurfacePanel();

    internal string MainCode { get { return main.Text.Trim(); } }
    internal string Channel { get { return channel.Text.Trim(); } }
    internal string ChannelName { get { return name.Text.Trim(); } }
    internal string Province { get { return province.Text.Trim(); } }

    internal MappingEditForm(Dictionary<string, object> record)
    {
        Text = record == null ? "添加通道" : "修改通道";
        Program.ApplyWindowIdentity(this);
        ClientSize = new Size(500, 390);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Font = UiFonts.Chinese(10F);
        BackColor = UiTheme.Canvas;
        Label title = new Label { Text = record == null ? "添加通道记录" : "修改通道记录",
            Font = UiFonts.Chinese(15F, FontStyle.Bold), ForeColor = UiTheme.Text, AutoSize = true,
            Location = new Point(28, 22) };
        Controls.Add(title);
        Label note = new Label { Text = "请确认主码号为 8 位数字，保存后立即应用于业务处理。",
            ForeColor = UiTheme.TextSecondary, AutoSize = true, Location = new Point(30, 59) };
        Controls.Add(note);
        fieldCard.SetBounds(24, 92, 452, 218);
        Controls.Add(fieldCard);
        AddField("主码号", main, 18);
        AddField("通道编号", channel, 66);
        AddField("通道名称", name, 114);
        AddField("省份", province, 162);
        if (record != null)
        {
            main.Text = Convert.ToString(record["main"]);
            channel.Text = Convert.ToString(record["channel"]);
            name.Text = Convert.ToString(record["name"]);
            province.Text = Convert.ToString(record["province"]);
        }
        Button save = new Button { Text = "保存修改", DialogResult = DialogResult.OK,
            FlatStyle = FlatStyle.Flat, BackColor = UiTheme.Primary, ForeColor = Color.White,
            Font = UiFonts.Chinese(10F, FontStyle.Bold), Cursor = Cursors.Hand };
        save.FlatAppearance.BorderSize = 0;
        save.SetBounds(350, 328, 126, 40);
        Button cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel,
            FlatStyle = FlatStyle.Flat, BackColor = UiTheme.Surface, ForeColor = UiTheme.TextSecondary,
            Font = UiFonts.Chinese(10F), Cursor = Cursors.Hand };
        cancel.FlatAppearance.BorderColor = UiTheme.BorderStrong;
        cancel.SetBounds(244, 328, 94, 40);
        Controls.Add(save);
        Controls.Add(cancel);
        AcceptButton = save;
        CancelButton = cancel;
    }

    private void AddField(string label, TextBox box, int y)
    {
        Label nameLabel = new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = UiTheme.TextSecondary, Font = UiFonts.Chinese(9.5F, FontStyle.Bold) };
        nameLabel.SetBounds(20, y, 92, 34);
        box.SetBounds(116, y, 314, 34);
        UiTheme.StyleTextBox(box);
        fieldCard.Controls.Add(nameLabel);
        fieldCard.Controls.Add(box);
    }
}
