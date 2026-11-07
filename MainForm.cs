using System.Globalization;
using System.Text.Json;

namespace ZMerge;

public class MainForm : Form
{
    Store st;
    readonly DataGridView grid = new();
    readonly ToolStripStatusLabel lblStatus = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    readonly ToolStripProgressBar bar = new() { Style = ProgressBarStyle.Marquee, Visible = false, Width = 120 };
    readonly TextBox txtFind = new() { Width = 220 };
    readonly ComboBox cmbCol = new() { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
    readonly Panel banner = new() { Dock = DockStyle.Top, Height = 38, Visible = false, BackColor = Color.FromArgb(255, 244, 214) };
    readonly Label bannerText = new() { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };
    readonly System.Windows.Forms.Timer findTimer = new() { Interval = 450 };
    readonly ToolStripMenuItem miDedupe = new("중복 행 건너뛰기") { CheckOnClick = true };
    ToolStrip tools;
    MenuStrip menu;

    List<long> ids = new();
    readonly Dictionary<int, object[][]> cache = new();
    List<(int col, bool desc)> sorts = new();
    long? srcFrom, srcTo;
    int[] keyCols = Array.Empty<int>();
    bool busy;
    CancellationTokenSource cts;

    static string SettingsPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "z2roN_Merge", "settings.json");
    static string DefaultDb => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "z2roN_Merge", "누적데이터.zmdb");

    public MainForm()
    {
        Text = $"z2roN 엑셀 병합 v{Program.Version}";
        Width = 1280; Height = 780;
        StartPosition = FormStartPosition.CenterScreen;
        AllowDrop = true;
        KeyPreview = true;
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape && cts != null) { cts.Cancel(); lblStatus.Text = "취소하는 중..."; }
            else if (e.KeyCode == Keys.Delete && ActiveControl == grid && !busy) { DeleteSelected(); e.Handled = true; }
            else if (e.KeyCode == Keys.Z && e.Control && ActiveControl == grid && !busy) { UndoDelete(); e.Handled = true; }
        };
        Font = new Font("Malgun Gothic", 9.5f);
        BuildUi();
        Load += (_, _) => OpenInitial();
        FormClosing += (_, e) => { if (busy) { MessageBox.Show("작업이 진행 중입니다. 끝난 뒤 닫아 주세요."); e.Cancel = true; } else st?.Dispose(); };
        DragEnter += (_, e) => { if (e.Data!.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
        DragDrop += async (_, e) => { if (e.Data!.GetData(DataFormats.FileDrop) is string[] f) await ImportFiles(f); };
    }

    // ---------------- UI 구성 ----------------
    void BuildUi()
    {
        menu = new MenuStrip();
        var file = new ToolStripMenuItem("파일(&F)");
        file.DropDownItems.Add(new ToolStripMenuItem("엑셀 올려서 추가...", null, async (_, _) => await PickAndImport()) { ShortcutKeys = Keys.Control | Keys.O });
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add(new ToolStripMenuItem("엑셀(.xlsx)로 내보내기...", null, async (_, _) => await Export(false)));
        file.DropDownItems.Add(new ToolStripMenuItem("CSV로 내보내기...", null, async (_, _) => await Export(true)));
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add(new ToolStripMenuItem("새 누적 파일 만들기...", null, (_, _) => NewDb()));
        file.DropDownItems.Add(new ToolStripMenuItem("누적 파일 열기...", null, (_, _) => OpenDb()));
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add(new ToolStripMenuItem("종료", null, (_, _) => Close()));
        var view = new ToolStripMenuItem("보기(&V)");
        view.DropDownItems.Add(new ToolStripMenuItem("전체 보기", null, (_, _) => ShowAll()) { ShortcutKeys = Keys.F5 });
        view.DropDownItems.Add(new ToolStripMenuItem("업로드 내역 / 되돌리기...", null, (_, _) => ShowHistory()));
        view.DropDownItems.Add(new ToolStripMenuItem("정렬 해제", null, (_, _) => { sorts.Clear(); RefreshView(); }));
        var edit = new ToolStripMenuItem("편집(&E)");
        edit.DropDownItems.Add(new ToolStripMenuItem("선택한 행 삭제", null, (_, _) => DeleteSelected()) { ShortcutKeyDisplayString = "Del" });
        edit.DropDownItems.Add(new ToolStripMenuItem("현재 화면의 모든 행 삭제...", null, (_, _) => DeleteAllShown()));
        edit.DropDownItems.Add(new ToolStripMenuItem("삭제 취소", null, (_, _) => UndoDelete()) { ShortcutKeyDisplayString = "Ctrl+Z" });
        edit.DropDownItems.Add(new ToolStripSeparator());
        edit.DropDownItems.Add(new ToolStripMenuItem("열 기준 중복 값 삭제...", null, (_, _) => DedupeDialog()));
        edit.DropDownItems.Add(new ToolStripMenuItem("현재 정렬 순서를 데이터에 확정(저장)...", null, (_, _) => CommitSortNow()));
        var opt = new ToolStripMenuItem("설정(&S)");
        opt.DropDownItems.Add(miDedupe);
        opt.DropDownItems.Add(new ToolStripMenuItem("중복 기준 열 선택...", null, (_, _) => PickKeyCols()));
        var help = new ToolStripMenuItem("도움말(&H)");
        help.DropDownItems.Add(new ToolStripMenuItem("사용 방법", null, (_, _) => MessageBox.Show(
            "1) 엑셀 파일을 창에 끌어다 놓거나 [엑셀 올리기]를 누르면 누적 데이터 맨 아래에 덧붙고, 올린 내용이 바로 표시됩니다.\n" +
            "2) 첫 파일의 열 구성이 기준이 되며, 이후 파일은 열 이름이 같아야 합니다(순서가 달라도 이름으로 맞춤).\n" +
            "3) 잘못 올렸으면 [되돌리기] 또는 보기 > 업로드 내역에서 파일 단위로 삭제할 수 있습니다.\n" +
            "4) 열 제목을 누르면 정렬 (Shift+클릭: 여러 열 정렬). 찾기 칸에 입력하면 바로 걸러집니다.\n" +
            "5) 줄 삭제: 왼쪽 행 번호를 눌러 줄을 선택(Ctrl/Shift 로 여러 줄)하고 Del 키 또는 [선택 행 삭제]. 직후 Ctrl+Z 로 취소합니다.\n" +
            "6) 편집 > 열 기준 중복 값 삭제: 체크한 열의 값이 같은 행 중 한 줄만 남깁니다.\n" +
            "7) 편집 > 현재 정렬 순서를 데이터에 확정: 정렬한 순서를 데이터 자체의 순서로 저장합니다.\n" +
            "8) 첫 시트만 읽습니다. 결과는 [엑셀로 내보내기]로 저장하세요.", "사용 방법")));
        menu.Items.AddRange(new ToolStripItem[] { file, edit, view, opt, help });

        tools = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden };
        tools.Items.Add(new ToolStripButton("엑셀 올리기", null, async (_, _) => await PickAndImport()) { DisplayStyle = ToolStripItemDisplayStyle.Text, Font = new Font(Font, FontStyle.Bold) });
        tools.Items.Add(new ToolStripSeparator());
        tools.Items.Add(new ToolStripLabel("찾기:"));
        tools.Items.Add(new ToolStripControlHost(txtFind));
        tools.Items.Add(new ToolStripControlHost(cmbCol));
        tools.Items.Add(new ToolStripSeparator());
        tools.Items.Add(new ToolStripButton("선택 행 삭제", null, (_, _) => DeleteSelected()) { DisplayStyle = ToolStripItemDisplayStyle.Text, ToolTipText = "행 번호를 눌러 줄을 선택한 뒤 (Ctrl/Shift 로 여러 줄) 삭제 — Del 키" });
        tools.Items.Add(new ToolStripButton("중복 삭제", null, (_, _) => DedupeDialog()) { DisplayStyle = ToolStripItemDisplayStyle.Text, ToolTipText = "선택한 열의 값이 같은 행 중 하나만 남기고 삭제" });
        tools.Items.Add(new ToolStripButton("삭제 취소", null, (_, _) => UndoDelete()) { DisplayStyle = ToolStripItemDisplayStyle.Text, ToolTipText = "Ctrl+Z" });
        tools.Items.Add(new ToolStripSeparator());
        tools.Items.Add(new ToolStripButton("전체 보기", null, (_, _) => ShowAll()) { DisplayStyle = ToolStripItemDisplayStyle.Text });
        tools.Items.Add(new ToolStripButton("엑셀로 내보내기", null, async (_, _) => await Export(false)) { DisplayStyle = ToolStripItemDisplayStyle.Text });

        var status = new StatusStrip();
        status.Items.Add(lblStatus);
        status.Items.Add(bar);

        grid.Dock = DockStyle.Fill;
        grid.VirtualMode = true;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.RowHeadersWidth = 72;
        grid.ColumnHeadersHeight = 30;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        grid.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        grid.MultiSelect = true;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        grid.BackgroundColor = SystemColors.Window;
        typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.SetValue(grid, true);
        grid.CellValueNeeded += Grid_CellValueNeeded;
        grid.ColumnHeaderMouseClick += Grid_HeaderClick;
        grid.RowPostPaint += (s, e) =>
        {
            var t = (e.RowIndex + 1).ToString("N0");
            var b = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth - 6, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, t, grid.Font, b, Color.Gray, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        };

        var btnUndo = new Button { Text = "방금 올린 것 되돌리기", Dock = DockStyle.Right, Width = 170 };
        var btnAll = new Button { Text = "전체 보기", Dock = DockStyle.Right, Width = 100 };
        btnUndo.Click += (_, _) => UndoLast();
        btnAll.Click += (_, _) => ShowAll();
        banner.Controls.Add(bannerText);
        banner.Controls.Add(btnAll);
        banner.Controls.Add(btnUndo);

        Controls.Add(grid);
        Controls.Add(banner);
        Controls.Add(tools);
        Controls.Add(menu);
        Controls.Add(status);
        MainMenuStrip = menu;

        findTimer.Tick += (_, _) => { findTimer.Stop(); RefreshView(); };
        txtFind.TextChanged += (_, _) => { findTimer.Stop(); findTimer.Start(); };
        cmbCol.SelectedIndexChanged += (_, _) => { if (!string.IsNullOrEmpty(txtFind.Text)) RefreshView(); };
    }

    // ---------------- 누적 파일 ----------------
    void OpenInitial()
    {
        string path = DefaultDb;
        try { if (File.Exists(SettingsPath)) { var d = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(SettingsPath)); if (d != null && d.TryGetValue("db", out var p) && File.Exists(p)) path = p; } } catch { }
        try { LoadDb(path); }
        catch (Exception e) { MessageBox.Show("누적 파일을 열 수 없습니다.\n" + e.Message); }
    }

    void LoadDb(string path)
    {
        st?.Dispose();
        st = Store.Open(path);
        try { Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!); File.WriteAllText(SettingsPath, JsonSerializer.Serialize(new Dictionary<string, string> { ["db"] = path })); } catch { }
        sorts.Clear(); srcFrom = srcTo = null; txtFind.Text = "";
        BuildColumns();
        RefreshView();
    }

    void NewDb()
    {
        using var d = new SaveFileDialog { Filter = "누적 데이터 (*.zmdb)|*.zmdb", FileName = "누적데이터.zmdb", Title = "새 누적 파일 만들기" };
        if (d.ShowDialog() != DialogResult.OK) return;
        if (File.Exists(d.FileName) && MessageBox.Show("이미 있는 파일입니다. 그 파일을 열까요?", "확인", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        LoadDb(d.FileName);
    }

    void OpenDb()
    {
        using var d = new OpenFileDialog { Filter = "누적 데이터 (*.zmdb)|*.zmdb|모든 파일|*.*", Title = "누적 파일 열기" };
        if (d.ShowDialog() == DialogResult.OK) LoadDb(d.FileName);
    }

    void BuildColumns()
    {
        grid.Columns.Clear();
        cmbCol.Items.Clear();
        cmbCol.Items.Add("전체 열");
        foreach (var h in st.Headers)
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = h, Width = 110, SortMode = DataGridViewColumnSortMode.Programmatic, Resizable = DataGridViewTriState.True });
            cmbCol.Items.Add(h);
        }
        cmbCol.SelectedIndex = 0;
        keyCols = Array.Empty<int>();
    }

    // ---------------- 보기(정렬/찾기) ----------------
    void RefreshView()
    {
        if (st == null) return;
        Cursor = Cursors.WaitCursor;
        try
        {
            string where = Store.BuildWhere(txtFind.Text.Trim(), cmbCol.SelectedIndex - 1, srcFrom, srcTo, st.Headers.Length, out var ps);
            ids = st.QueryIds(where, ps, Store.BuildOrder(sorts));
            cache.Clear();
            grid.RowCount = 0;
            grid.RowCount = ids.Count;
            grid.Invalidate();
            UpdateHeaderMarks();
            UpdateStatus();
        }
        catch (Exception e) { MessageBox.Show(e.Message, "오류"); }
        finally { Cursor = Cursors.Default; }
    }

    void UpdateHeaderMarks()
    {
        for (int i = 0; i < grid.Columns.Count; i++)
        {
            int k = sorts.FindIndex(s => s.col == i);
            string mark = k < 0 ? "" : (sorts[k].desc ? " ▼" : " ▲") + (sorts.Count > 1 ? (k + 1).ToString() : "");
            grid.Columns[i].HeaderText = st.Headers[i] + mark;
        }
    }

    void UpdateStatus()
    {
        long total = st.TotalRows();
        string flt = (!string.IsNullOrEmpty(txtFind.Text) ? " (찾기 적용)" : "") + (srcFrom != null ? " (방금 올린 내용만)" : "");
        lblStatus.Text = $"누적 {total:N0}행 · 화면 {ids.Count:N0}행{flt} · 열 {st.Headers.Length}개 · {st.DbPath}";
    }

    void Grid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= ids.Count || st == null) return;
        int blk = e.RowIndex / Store.Block;
        if (!cache.TryGetValue(blk, out var rows))
        {
            if (cache.Count > 40) cache.Clear();
            rows = st.Fetch(ids, blk * Store.Block, Store.Block);
            cache[blk] = rows;
        }
        var v = rows[e.RowIndex % Store.Block][e.ColumnIndex];
        e.Value = v switch
        {
            null => "",
            double d => d.ToString("0.###############", CultureInfo.InvariantCulture),
            _ => v
        };
    }

    void Grid_HeaderClick(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left || busy) return;
        int c = e.ColumnIndex;
        int k = sorts.FindIndex(s => s.col == c);
        bool multi = (ModifierKeys & (Keys.Shift | Keys.Control)) != 0;
        if (!multi) { var cur = k >= 0 ? sorts[k] : default; sorts.Clear(); if (k >= 0) sorts.Add(cur); k = sorts.Count > 0 ? 0 : -1; }
        if (k < 0) sorts.Add((c, false));
        else if (!sorts[k].desc) sorts[k] = (c, true);
        else sorts.RemoveAt(k);
        RefreshView();
    }

    void ShowAll()
    {
        srcFrom = srcTo = null;
        banner.Visible = false;
        txtFind.Text = "";
        if (cmbCol.Items.Count > 0) cmbCol.SelectedIndex = 0;
        RefreshView();
    }

    // ---------------- 행 삭제 / 정렬 확정 ----------------
    void AfterEdit()
    {
        int top = grid.FirstDisplayedScrollingRowIndex;
        RefreshView();
        if (top >= 0 && top < grid.RowCount) { try { grid.FirstDisplayedScrollingRowIndex = top; } catch { } }
    }

    void DeleteSelected()
    {
        if (busy || st == null || st.Headers.Length == 0) return;
        var rows = new SortedSet<int>();
        if (grid.SelectedCells.Count > 300000) { MessageBox.Show("선택한 칸이 너무 많습니다. 많은 행은 [편집 > 현재 화면의 모든 행 삭제]를 사용하세요."); return; }
        foreach (DataGridViewCell c in grid.SelectedCells) rows.Add(c.RowIndex);
        if (rows.Count == 0) { MessageBox.Show("삭제할 줄을 선택하세요.\n\n왼쪽 행 번호를 누르면 줄이 선택됩니다. (Ctrl/Shift 로 여러 줄)", "행 삭제"); return; }
        if (MessageBox.Show($"선택한 {rows.Count:N0}개 행을 삭제할까요?\n(삭제 직후에는 [삭제 취소]로 되돌릴 수 있습니다)", "행 삭제", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        DoDelete(rows.Where(r => r >= 0 && r < ids.Count).Select(r => ids[r]).ToList());
    }

    void DeleteAllShown()
    {
        if (busy || st == null || ids.Count == 0) return;
        string scope = (!string.IsNullOrEmpty(txtFind.Text.Trim()) || srcFrom != null) ? "현재 화면(찾기/방금 올린 것)에 보이는" : "누적 데이터의 전체";
        if (MessageBox.Show($"{scope} {ids.Count:N0}개 행을 모두 삭제할까요?\n(삭제 직후에는 [삭제 취소]로 되돌릴 수 있습니다)", "행 삭제", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        DoDelete(new List<long>(ids));
    }

    void DoDelete(List<long> rids)
    {
        Cursor = Cursors.WaitCursor;
        try { st.DeleteRows(rids); AfterEdit(); lblStatus.Text = $"{rids.Count:N0}개 행을 삭제했습니다. (Ctrl+Z 로 취소) · 누적 {st.TotalRows():N0}행"; }
        catch (Exception e) { MessageBox.Show(e.Message, "삭제 오류"); }
        finally { Cursor = Cursors.Default; }
    }

    void UndoDelete()
    {
        if (busy || st == null) return;
        if (!st.CanUndoDelete()) { lblStatus.Text = "되돌릴 삭제가 없습니다."; return; }
        Cursor = Cursors.WaitCursor;
        try { long n = st.UndoDelete(); AfterEdit(); lblStatus.Text = $"삭제를 취소하고 {n:N0}개 행을 복원했습니다."; }
        catch (Exception e) { MessageBox.Show(e.Message, "오류"); }
        finally { Cursor = Cursors.Default; }
    }

    void DedupeDialog()
    {
        if (busy || st == null || st.Headers.Length == 0) { MessageBox.Show("먼저 엑셀을 올려 데이터를 만드세요."); return; }
        using var f = new Form { Text = "열 기준 중복 값 삭제", Width = 460, Height = 640, StartPosition = FormStartPosition.CenterParent, Font = Font, MinimizeBox = false, MaximizeBox = false };
        var info = new Label { Dock = DockStyle.Top, Height = 62, Padding = new Padding(8), Text = "체크한 열의 값이 모두 같은 행을 중복으로 보고, 그중 한 줄만 남기고 나머지를 삭제합니다.\n(전체 누적 데이터가 대상이며, 삭제 후 Ctrl+Z 로 취소할 수 있습니다)" };
        var clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
        for (int i = 0; i < st.Headers.Length; i++) clb.Items.Add(st.Headers[i]);
        var pnl = new Panel { Dock = DockStyle.Bottom, Height = 190 };
        var rbFirst = new RadioButton { Text = sorts.Count > 0 ? "현재 정렬 순서에서 맨 위(첫) 행을 남김" : "먼저 올라온(위쪽) 행을 남김", Left = 10, Top = 8, Width = 420, Checked = true };
        var rbLast = new RadioButton { Text = sorts.Count > 0 ? "현재 정렬 순서에서 맨 아래(마지막) 행을 남김" : "나중에 올라온(아래쪽) 행을 남김", Left = 10, Top = 32, Width = 420 };
        var cbBlank = new CheckBox { Text = "비어 있는 값끼리는 중복으로 보지 않기", Left = 10, Top = 60, Width = 420, Checked = true };
        var cbNorm = new CheckBox { Text = "앞뒤 공백·대소문자 무시", Left = 10, Top = 84, Width = 420 };
        var lbl = new Label { Left = 10, Top = 112, Width = 420, Height = 36, Text = "열을 고르고 [중복 확인]을 누르세요." };
        var btnCheck = new Button { Text = "중복 확인", Left = 10, Top = 148, Width = 120, Height = 32 };
        var btnDel = new Button { Text = "중복 삭제", Left = 140, Top = 148, Width = 120, Height = 32, Enabled = false };
        var btnClose = new Button { Text = "닫기", Left = 270, Top = 148, Width = 100, Height = 32, DialogResult = DialogResult.Cancel };
        pnl.Controls.AddRange(new Control[] { rbFirst, rbLast, cbBlank, cbNorm, lbl, btnCheck, btnDel, btnClose });
        List<long> found = null;
        void Reset() { found = null; btnDel.Enabled = false; lbl.Text = "열을 고르고 [중복 확인]을 누르세요."; }
        clb.ItemCheck += (_, _) => BeginInvoke(new Action(Reset));
        rbFirst.CheckedChanged += (_, _) => Reset(); cbBlank.CheckedChanged += (_, _) => Reset(); cbNorm.CheckedChanged += (_, _) => Reset();
        btnCheck.Click += (_, _) =>
        {
            var cols = clb.CheckedIndices.Cast<int>().ToArray();
            if (cols.Length == 0) { MessageBox.Show("기준이 될 열을 하나 이상 체크하세요."); return; }
            f.Cursor = Cursors.WaitCursor;
            try
            {
                found = st.FindDuplicateRids(cols, Store.BuildOrder(sorts, rbLast.Checked), cbBlank.Checked, cbNorm.Checked, out long groups);
                lbl.Text = found.Count == 0 ? "중복된 행이 없습니다." : $"중복 값 {groups:N0}종, 삭제될 행 {found.Count:N0}개 (남는 행 {st.TotalRows() - found.Count:N0}개)";
                btnDel.Enabled = found.Count > 0;
            }
            catch (Exception e) { MessageBox.Show(e.Message, "오류"); }
            finally { f.Cursor = Cursors.Default; }
        };
        btnDel.Click += (_, _) =>
        {
            if (found == null || found.Count == 0) return;
            if (MessageBox.Show($"{found.Count:N0}개 중복 행을 삭제할까요?\n(삭제 직후 Ctrl+Z 로 취소할 수 있습니다)", "중복 삭제", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            f.Close();
            srcFrom = srcTo = null; banner.Visible = false;
            DoDelete(found);
        };
        f.Controls.Add(clb); f.Controls.Add(info); f.Controls.Add(pnl);
        f.CancelButton = btnClose;
        f.ShowDialog(this);
    }

    void CommitSortNow()
    {
        if (busy || st == null || st.Headers.Length == 0) return;
        if (sorts.Count == 0) { MessageBox.Show("먼저 열 제목을 눌러 정렬하세요. (Shift+클릭으로 여러 열)"); return; }
        if (MessageBox.Show("현재 정렬 순서대로 누적 데이터의 순서를 확정합니다.\n이후 새로 올리는 행은 맨 아래에 붙고, 이전의 삭제 취소 기록은 지워집니다.\n\n진행할까요?", "정렬 확정", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        Cursor = Cursors.WaitCursor;
        try { st.CommitSort(Store.BuildOrder(sorts)); sorts.Clear(); srcFrom = srcTo = null; banner.Visible = false; RefreshView(); lblStatus.Text = "정렬 순서를 데이터에 확정했습니다."; }
        catch (Exception e) { MessageBox.Show(e.Message, "오류"); }
        finally { Cursor = Cursors.Default; }
    }

    // ---------------- 올리기 ----------------
    async Task PickAndImport()
    {
        if (busy) return;
        using var d = new OpenFileDialog { Filter = "엑셀/CSV (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv|모든 파일|*.*", Multiselect = true, Title = "누적할 엑셀 파일 선택" };
        if (d.ShowDialog() == DialogResult.OK) await ImportFiles(d.FileNames);
    }

    void SetBusy(bool b)
    {
        busy = b;
        bar.Visible = b;
        menu.Enabled = !b;
        tools.Enabled = !b;
        grid.Enabled = !b;
    }

    async Task ImportFiles(string[] files)
    {
        if (busy || st == null) return;
        files = files.Where(f => new[] { ".xls", ".xlsx", ".csv" }.Contains(Path.GetExtension(f).ToLowerInvariant())).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
        if (files.Length == 0) { MessageBox.Show("엑셀(.xls, .xlsx) 또는 .csv 파일만 올릴 수 있습니다."); return; }
        SetBusy(true);
        cts = new CancellationTokenSource();
        var prog = new Progress<string>(m => lblStatus.Text = m);
        long added = 0, skipped = 0, firstId = 0, lastId = 0;
        int okCount = 0;
        bool needColumns = st.Headers.Length == 0;
        bool dd = miDedupe.Checked;
        int[] kcs = keyCols;
        var token = cts.Token;
        try
        {
            foreach (var f in files)
            {
                try
                {
                    var res = await Task.Run(() => Importer.Run(st, f, dd, kcs, prog, token));
                    if (okCount == 0) firstId = res.UploadId;
                    lastId = res.UploadId; added += res.Added; skipped += res.Skipped; okCount++;
                    if (res.CreatedLayout) BuildColumns();
                }
                catch (ImportException ex) { MessageBox.Show($"{Path.GetFileName(f)}\n\n{ex.Message}", "올릴 수 없습니다", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    var hint = f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ? "\n\n(구형 .xls 를 읽지 못하는 경우 엑셀에서 .xlsx 로 다시 저장해서 올려 보세요.)" : "";
                    MessageBox.Show($"{Path.GetFileName(f)}\n\n{ex.Message}{hint}", "읽기 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        finally { SetBusy(false); cts = null; }

        if (okCount > 0)
        {
            sorts.Clear();
            txtFind.Text = "";
            srcFrom = firstId; srcTo = lastId;
            bannerText.Text = $"방금 올린 내용 {added:N0}행 추가됨" + (skipped > 0 ? $" (중복 {skipped:N0}행 건너뜀)" : "") + (okCount > 1 ? $" · 파일 {okCount}개" : "") + "   ← 내용을 확인하세요";
            banner.Visible = true;
            RefreshView();
        }
        else if (needColumns) { BuildColumns(); RefreshView(); }
    }

    void UndoLast()
    {
        if (srcFrom == null || busy) return;
        if (MessageBox.Show("방금 올린 내용을 모두 삭제하고 올리기 전 상태로 되돌릴까요?", "되돌리기", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        Cursor = Cursors.WaitCursor;
        try
        {
            var del = st.Uploads().Where(u => u.Id >= srcFrom && u.Id <= srcTo).Select(u => u.Id).ToList();
            st.DeleteUploads(del);
            banner.Visible = false;
            srcFrom = srcTo = null;
            BuildColumns();
            RefreshView();
        }
        finally { Cursor = Cursors.Default; }
    }

    void ShowHistory()
    {
        if (st == null) return;
        using var f = new Form { Text = "업로드 내역", Width = 760, Height = 460, StartPosition = FormStartPosition.CenterParent, Font = Font };
        var lv = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = true, CheckBoxes = false };
        lv.Columns.Add("번호", 60); lv.Columns.Add("파일 이름", 320); lv.Columns.Add("올린 시각", 150); lv.Columns.Add("추가 행", 80, HorizontalAlignment.Right); lv.Columns.Add("건너뜀", 70, HorizontalAlignment.Right);
        void Fill() { lv.Items.Clear(); foreach (var u in st.Uploads()) lv.Items.Add(new ListViewItem(new[] { u.Id.ToString(), u.Name, u.At, u.Rows.ToString("N0"), u.Skipped.ToString("N0") }) { Tag = u.Id }); }
        Fill();
        var pnl = new Panel { Dock = DockStyle.Bottom, Height = 44 };
        var del = new Button { Text = "선택한 업로드 삭제", Left = 10, Top = 8, Width = 160 };
        var show = new Button { Text = "선택한 것만 보기", Left = 180, Top = 8, Width = 140 };
        del.Click += (_, _) =>
        {
            var sel = lv.SelectedItems.Cast<ListViewItem>().Select(i => (long)i.Tag).ToList();
            if (sel.Count == 0) return;
            if (MessageBox.Show($"선택한 {sel.Count}개 업로드의 데이터를 삭제할까요?", "삭제", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            st.DeleteUploads(sel); Fill(); BuildColumns(); banner.Visible = false; srcFrom = srcTo = null; RefreshView();
        };
        show.Click += (_, _) =>
        {
            var sel = lv.SelectedItems.Cast<ListViewItem>().Select(i => (long)i.Tag).ToList();
            if (sel.Count == 0) return;
            srcFrom = sel.Min(); srcTo = sel.Max();
            bannerText.Text = $"선택한 업로드 보기 (번호 {srcFrom}~{srcTo})"; banner.Visible = true; RefreshView(); f.Close();
        };
        pnl.Controls.Add(del); pnl.Controls.Add(show);
        f.Controls.Add(lv); f.Controls.Add(pnl);
        f.ShowDialog(this);
    }

    void PickKeyCols()
    {
        if (st == null || st.Headers.Length == 0) { MessageBox.Show("먼저 첫 엑셀을 올려 열 구성을 만드세요."); return; }
        using var f = new Form { Text = "중복 기준 열 선택", Width = 380, Height = 520, StartPosition = FormStartPosition.CenterParent, Font = Font };
        var clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
        for (int i = 0; i < st.Headers.Length; i++) clb.Items.Add(st.Headers[i], keyCols.Contains(i));
        var info = new Label { Dock = DockStyle.Top, Height = 44, Text = "체크한 열의 값이 모두 같으면 같은 행으로 봅니다.\n아무것도 체크하지 않으면 모든 열이 같은 행만 중복입니다.", Padding = new Padding(6) };
        var ok = new Button { Text = "확인", Dock = DockStyle.Bottom, Height = 36, DialogResult = DialogResult.OK };
        f.Controls.Add(clb); f.Controls.Add(info); f.Controls.Add(ok);
        if (f.ShowDialog(this) == DialogResult.OK)
        {
            keyCols = clb.CheckedIndices.Cast<int>().ToArray();
            miDedupe.Checked = true;
            lblStatus.Text = keyCols.Length == 0 ? "중복 기준: 모든 열" : "중복 기준: " + string.Join(", ", keyCols.Select(i => st.Headers[i]));
        }
    }

    // ---------------- 내보내기 ----------------
    async Task Export(bool csv)
    {
        if (busy || st == null || st.Headers.Length == 0) { MessageBox.Show("내보낼 데이터가 없습니다."); return; }
        bool filtered = !string.IsNullOrEmpty(txtFind.Text.Trim()) || srcFrom != null;
        bool useView = false;
        if (filtered)
        {
            var a = MessageBox.Show("현재 화면(찾기/방금 올린 것만)에 보이는 행만 내보낼까요?\n\n예: 현재 화면만\n아니오: 전체 누적 데이터", "내보내기", MessageBoxButtons.YesNoCancel);
            if (a == DialogResult.Cancel) return;
            useView = a == DialogResult.Yes;
        }
        using var d = new SaveFileDialog { Filter = csv ? "CSV (*.csv)|*.csv" : "엑셀 (*.xlsx)|*.xlsx", FileName = "합본_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + (csv ? ".csv" : ".xlsx") };
        if (d.ShowDialog() != DialogResult.OK) return;
        string where = ""; List<(string, object)> ps = new();
        if (useView) where = Store.BuildWhere(txtFind.Text.Trim(), cmbCol.SelectedIndex - 1, srcFrom, srcTo, st.Headers.Length, out ps);
        string order = Store.BuildOrder(sorts);
        SetBusy(true);
        var prog = new Progress<string>(m => lblStatus.Text = m);
        try
        {
            if (csv)
            {
                long n = await Task.Run(() => st.ExportCsv(d.FileName, where, ps, order, prog));
                MessageBox.Show($"{n:N0}행을 저장했습니다.\n{d.FileName}");
            }
            else
            {
                var files = await Task.Run(() => st.ExportXlsx(d.FileName, where, ps, order, prog));
                MessageBox.Show(files.Count == 1 ? $"저장했습니다.\n{files[0]}" : $"엑셀 한 시트 한도(약 104만 행) 때문에 {files.Count}개 파일로 나누어 저장했습니다.\n{string.Join("\n", files)}");
            }
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "내보내기 오류", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        finally { SetBusy(false); UpdateStatus(); }
    }
}
