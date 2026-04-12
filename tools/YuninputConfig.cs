using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

internal static class ConfigDefaults
{
    public const string DefaultDictionaryProfile = "zhengma-all";
    public const decimal MinPageSize = 1;
    public const decimal MaxPageSize = 6;
    public const decimal DefaultPageSize = 6;
    public const decimal MinAutoCommitCodeLength = 2;
    public const decimal MaxAutoCommitCodeLength = 8;
    public const decimal DefaultAutoCommitCodeLength = 4;
    public const decimal MinContextAssociationMaxEntries = 1000;
    public const decimal MaxContextAssociationMaxEntries = 50000;
    public const decimal DefaultContextAssociationMaxEntries = 6000;
    public const bool DefaultChinesePunctuation = true;
    public const bool DefaultSmartSymbolPairs = true;
    public const bool DefaultAutoCommitUniqueExact = true;
    public const bool DefaultEmptyCandidateBeep = true;
    public const bool DefaultTabNavigation = true;
    public const bool DefaultEnterExactPriority = true;
    public const bool DefaultContextAssociationEnabled = true;
}

internal sealed class CandidateEntryItem
{
    public string Code;
    public string Text;
    public string RawLine;
    public bool IsAutoPhrase;
    public int Score = 1;

    public override string ToString()
    {
        return Code + "    " + Text;
    }
}

internal sealed class ContextAssocItem
{
    public string PrevText;
    public string NextText;
    public ulong Score;

    public override string ToString()
    {
        return PrevText + " -> " + NextText + "    (" + Score + ")";
    }
}

internal sealed class PhraseReviewLogItem
{
    public string Code;
    public string Text;
    public string Source;
    public string Time;
    public string RawLine;

    public override string ToString()
    {
        return Code + "    " + Text + "    [" + Source + "] " + Time;
    }
}

public class ConfigForm : Form
{
    private static readonly Color PanelBackColor = Color.FromArgb(247, 248, 250);
    private static readonly Color AccentColor = Color.FromArgb(31, 95, 168);
    private static readonly Color BorderColor = Color.FromArgb(214, 219, 226);
    private static readonly Color AlternateRowColor = Color.FromArgb(250, 252, 255);
    private static readonly Color SelectedRowColor = Color.FromArgb(220, 235, 252);

    private CheckBox chkChinese;
    private CheckBox chkFull;
    private CheckBox chkChinesePunctuation;
    private CheckBox chkSmartSymbolPairs;
    private CheckBox chkAutoCommitUniqueExact;
    private NumericUpDown numAutoCommitMinCodeLength;
    private CheckBox chkEmptyCandidateBeep;
    private CheckBox chkTabNavigation;
    private CheckBox chkEnterExactPriority;
    private CheckBox chkContextAssociationEnabled;
    private NumericUpDown numContextAssociationMaxEntries;
    private NumericUpDown numPage;
    private ComboBox cmbHotkey;
    private ComboBox cmbDictionaryProfile;
    private ListBox lstPinned;
    private ListBox lstBlocked;
    private ListBox lstContextAssoc;
    private ListBox lstContextAssocBlacklist;
    private Label lblPinnedPath;
    private Label lblBlockedPath;
    private Label lblContextAssocPath;
    private Label lblContextAssocBlacklistPath;
    private string cfgPath;
    private string localRoot;
    private string userDictPath;
    private string autoPhraseDictPath;
    private string userFreqPath;
    private string blockedPath;
    private string contextAssocPath;
    private string contextAssocBlacklistPath;
    private string contextAssocDryRunReportPath;
    private string manualPhraseReviewPath;
    private string roamingRoot;

    public ConfigForm()
    {
        Text = "\u5300\u7801\u8f93\u5165\u6cd5 \u914d\u7f6e\u4e0e\u8bcd\u6761\u7ba1\u7406";
        ClientSize = new Size(996, 748);
        MinimumSize = new Size(980, 720);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = PanelBackColor;
        AutoScaleMode = AutoScaleMode.Dpi;
        EnableEscapeClose(this);

        localRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "yuninput");
        roamingRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "yuninput");
        Directory.CreateDirectory(localRoot);
        Directory.CreateDirectory(roamingRoot);

        cfgPath = Path.Combine(localRoot, "settings.json");
        userDictPath = Path.Combine(roamingRoot, "yuninput_user.dict");
    autoPhraseDictPath = Path.Combine(roamingRoot, "auto_phrase_runtime.dict");
        userFreqPath = Path.Combine(roamingRoot, "user_freq.txt");
        blockedPath = Path.Combine(roamingRoot, "blocked_entries.txt");
        contextAssocPath = Path.Combine(roamingRoot, "context_assoc.txt");
        contextAssocBlacklistPath = Path.Combine(roamingRoot, "context_assoc_blacklist.txt");
        contextAssocDryRunReportPath = Path.Combine(roamingRoot, "context_assoc_dry_run_report.txt");
        manualPhraseReviewPath = Path.Combine(roamingRoot, "manual_phrase_review.txt");

        var tabs = new TabControl();
        tabs.Left = 16;
        tabs.Top = 16;
        tabs.Width = 964;
        tabs.Height = 646;
        tabs.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        Controls.Add(tabs);

        var generalTab = new TabPage("\u5e38\u89c4\u8bbe\u7f6e");
        var dataTab = new TabPage("\u7528\u6237\u8bcd\u6761");
        var contextTab = new TabPage("\u4e0a\u4e0b\u6587\u5b66\u4e60");
        tabs.TabPages.Add(generalTab);
        tabs.TabPages.Add(dataTab);
        tabs.TabPages.Add(contextTab);

        BuildGeneralTab(generalTab);
        BuildDataTab(dataTab);
        BuildContextTab(contextTab);

        var bottomBar = new Panel
        {
            Left = 16,
            Top = 676,
            Width = 964,
            Height = 40,
            BackColor = PanelBackColor,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };
        Controls.Add(bottomBar);

        var btnHelp = new Button { Left = 428, Top = 3, Width = 126, Height = 34, Text = "\u6253\u5f00\u8bf4\u660e\u4e66", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnSave = new Button { Left = 568, Top = 3, Width = 92, Height = 34, Text = "\u4fdd\u5b58", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnRefresh = new Button { Left = 670, Top = 3, Width = 92, Height = 34, Text = "\u5237\u65b0", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnOpenFolder = new Button { Left = 772, Top = 3, Width = 110, Height = 34, Text = "\u6253\u5f00\u76ee\u5f55", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnClose = new Button { Left = 892, Top = 3, Width = 72, Height = 34, Text = "\u5173\u95ed", Anchor = AnchorStyles.Right | AnchorStyles.Bottom, DialogResult = DialogResult.Cancel };

        btnHelp.Click += (s, e) => OpenManualInNotepad();
        btnSave.Click += (s, e) => SaveConfig(true);
        btnRefresh.Click += (s, e) =>
        {
            LoadConfig();
            RefreshDataLists();
            RefreshContextAssocLists();
        };
        btnOpenFolder.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = roamingRoot, UseShellExecute = true });
        btnClose.Click += (s, e) => Close();

        bottomBar.Controls.Add(btnHelp);
        bottomBar.Controls.Add(btnSave);
        bottomBar.Controls.Add(btnRefresh);
        bottomBar.Controls.Add(btnOpenFolder);
        bottomBar.Controls.Add(btnClose);

        AcceptButton = btnSave;
        CancelButton = btnClose;

        LoadConfig();
        RefreshDataLists();
        RefreshContextAssocLists();
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            Close();
            return true;
        }

        return base.ProcessCmdKey(ref msg, keyData);
    }

    private string ResolveManualPath()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string[] candidates = new[]
        {
            Path.Combine(baseDir, "\u5300\u7801\u8f93\u5165\u6cd5\u8bf4\u660e\u4e66.md"),
            Path.GetFullPath(Path.Combine(baseDir, "..", "\u5300\u7801\u8f93\u5165\u6cd5\u8bf4\u660e\u4e66.md")),
            Path.GetFullPath(Path.Combine(baseDir, "..", "README.md")),
            Path.Combine(baseDir, "README.md")
        };

        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return string.Empty;
    }

    private void OpenManualInNotepad()
    {
        string manualPath = ResolveManualPath();
        if (string.IsNullOrEmpty(manualPath))
        {
            MessageBox.Show("未找到本地说明书文件，请确认安装目录中的《匀码输入法说明书.md》是否存在。", "匀码输入法", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "notepad.exe",
            Arguments = "\"" + manualPath + "\"",
            UseShellExecute = true
        });
    }

    private static void EnableEscapeClose(Form form)
    {
        form.KeyPreview = true;
        form.KeyDown += (s, e) =>
        {
            if (e.KeyCode != Keys.Escape)
            {
                return;
            }

            form.Close();
            e.Handled = true;
        };
    }

    private static string BuildEntryLine(CandidateEntryItem item)
    {
        int score = item == null || item.Score <= 0 ? 1 : item.Score;
        if (item == null)
        {
            return string.Empty;
        }

        return item.Code + " " + item.Text + " " + score + (item.IsAutoPhrase ? " auto" : string.Empty);
    }

    private List<CandidateEntryItem> ReadManagedDictionaryItems()
    {
        var merged = new Dictionary<string, CandidateEntryItem>(StringComparer.OrdinalIgnoreCase);

        foreach (CandidateEntryItem item in ReadEntryFile(autoPhraseDictPath))
        {
            item.IsAutoPhrase = true;
            merged[item.Code + "\t" + item.Text] = item;
        }

        foreach (CandidateEntryItem item in ReadEntryFile(userDictPath))
        {
            item.IsAutoPhrase = false;
            merged[item.Code + "\t" + item.Text] = item;
        }

        var result = new List<CandidateEntryItem>(merged.Values);
        result.Sort((left, right) =>
        {
            if (left.IsAutoPhrase != right.IsAutoPhrase)
            {
                return left.IsAutoPhrase ? 1 : -1;
            }

            int codeComparison = string.Compare(left.Code, right.Code, StringComparison.OrdinalIgnoreCase);
            if (codeComparison != 0)
            {
                return codeComparison;
            }

            return string.Compare(left.Text, right.Text, StringComparison.Ordinal);
        });
        return result;
    }

    private void WriteManagedDictionaryItems(IList<CandidateEntryItem> items)
    {
        var manualLines = new List<string>();
        var autoLines = new List<string>();
        foreach (CandidateEntryItem item in items)
        {
            string line = BuildEntryLine(item);
            if (string.IsNullOrEmpty(line))
            {
                continue;
            }

            if (item.IsAutoPhrase)
            {
                autoLines.Add(line);
            }
            else
            {
                manualLines.Add(line);
            }
        }

        Directory.CreateDirectory(Path.GetDirectoryName(userDictPath));
        Directory.CreateDirectory(Path.GetDirectoryName(autoPhraseDictPath));
        File.WriteAllLines(userDictPath, manualLines.ToArray(), Encoding.UTF8);
        File.WriteAllLines(autoPhraseDictPath, autoLines.ToArray(), Encoding.UTF8);
    }

    private GroupBox CreateGroupBox(TabPage tab, string title, int left, int top, int width, int height)
    {
        var box = new GroupBox
        {
            Left = left,
            Top = top,
            Width = width,
            Height = height,
            Text = title,
            BackColor = Color.White,
            ForeColor = AccentColor
        };
        tab.Controls.Add(box);
        return box;
    }

    private ListBox CreateStyledListBox(int left, int top, int width, int height)
    {
        var listBox = new ListBox
        {
            Left = left,
            Top = top,
            Width = width,
            Height = height,
            DrawMode = DrawMode.OwnerDrawFixed,
            IntegralHeight = false,
            ItemHeight = 30,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White,
            HorizontalScrollbar = false
        };
        listBox.DrawItem += DrawStyledListBoxItem;
        return listBox;
    }

    private void DrawStyledListBoxItem(object sender, DrawItemEventArgs e)
    {
        if (e.Index < 0)
        {
            return;
        }

        var listBox = sender as ListBox;
        if (listBox == null)
        {
            return;
        }

        object item = listBox.Items[e.Index];
        bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        Color rowColor = selected ? SelectedRowColor : (e.Index % 2 == 0 ? Color.White : AlternateRowColor);
        using (var backBrush = new SolidBrush(rowColor))
        {
            e.Graphics.FillRectangle(backBrush, e.Bounds);
        }

        Rectangle inner = new Rectangle(e.Bounds.Left + 8, e.Bounds.Top + 5, e.Bounds.Width - 16, e.Bounds.Height - 10);
        TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis;
        Color textColor = selected ? Color.FromArgb(15, 47, 82) : Color.FromArgb(36, 41, 47);

        CandidateEntryItem entry = item as CandidateEntryItem;
        if (entry != null)
        {
            Rectangle codeRect = new Rectangle(inner.Left, inner.Top, 102, inner.Height);
            Rectangle textRect = new Rectangle(inner.Left + 110, inner.Top, Math.Max(80, inner.Width - 178), inner.Height);
            Rectangle tagRect = new Rectangle(inner.Right - 60, inner.Top, 60, inner.Height);
            TextRenderer.DrawText(e.Graphics, entry.Code ?? string.Empty, Font, codeRect, AccentColor, flags);
            TextRenderer.DrawText(e.Graphics, entry.Text ?? string.Empty, Font, textRect, textColor, flags);

            if (entry.IsAutoPhrase)
            {
                TextRenderer.DrawText(e.Graphics, "自动", Font, tagRect, Color.FromArgb(181, 112, 24), TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            }
        }
        else
        {
            ContextAssocItem assoc = item as ContextAssocItem;
            if (assoc != null)
            {
                Rectangle prevRect = new Rectangle(inner.Left, inner.Top, 72, inner.Height);
                Rectangle arrowRect = new Rectangle(inner.Left + 76, inner.Top, 30, inner.Height);
                Rectangle nextRect = new Rectangle(inner.Left + 108, inner.Top, Math.Max(80, inner.Width - 190), inner.Height);
                Rectangle scoreRect = new Rectangle(inner.Right - 84, inner.Top, 84, inner.Height);
                TextRenderer.DrawText(e.Graphics, assoc.PrevText ?? string.Empty, Font, prevRect, AccentColor, flags);
                TextRenderer.DrawText(e.Graphics, "->", Font, arrowRect, Color.FromArgb(122, 128, 138), flags);
                TextRenderer.DrawText(e.Graphics, assoc.NextText ?? string.Empty, Font, nextRect, textColor, flags);
                TextRenderer.DrawText(e.Graphics, assoc.Score.ToString(), Font, scoreRect, Color.FromArgb(97, 103, 112), TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            }
            else
            {
                PhraseReviewLogItem review = item as PhraseReviewLogItem;
                if (review != null)
                {
                    Rectangle codeRect = new Rectangle(inner.Left, inner.Top, 90, inner.Height);
                    Rectangle textRect = new Rectangle(inner.Left + 98, inner.Top, 90, inner.Height);
                    Rectangle sourceRect = new Rectangle(inner.Left + 194, inner.Top, 92, inner.Height);
                    Rectangle timeRect = new Rectangle(inner.Left + 290, inner.Top, Math.Max(80, inner.Width - 290), inner.Height);
                    TextRenderer.DrawText(e.Graphics, review.Code ?? string.Empty, Font, codeRect, AccentColor, flags);
                    TextRenderer.DrawText(e.Graphics, review.Text ?? string.Empty, Font, textRect, textColor, flags);
                    TextRenderer.DrawText(e.Graphics, review.Source ?? string.Empty, Font, sourceRect, Color.FromArgb(181, 112, 24), flags);
                    TextRenderer.DrawText(e.Graphics, review.Time ?? string.Empty, Font, timeRect, Color.FromArgb(97, 103, 112), flags);
                }
                else
                {
                    TextRenderer.DrawText(e.Graphics, Convert.ToString(item), Font, inner, textColor, flags);
                }
            }
        }

        using (var linePen = new Pen(BorderColor))
        {
            e.Graphics.DrawLine(linePen, e.Bounds.Left + 1, e.Bounds.Bottom - 1, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
        }

        e.DrawFocusRectangle();
    }

    private void BuildGeneralTab(TabPage tab)
    {
        tab.BackColor = PanelBackColor;
        var lblIntro = new Label { Left = 20, Top = 18, Width = 840, Height = 24, Text = "这里集中管理默认状态、候选行为、词库模式和常用维护入口。按 Esc 可直接关闭窗口。" };
        lblIntro.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        tab.Controls.Add(lblIntro);
        var behaviorGroup = CreateGroupBox(tab, "输入与候选行为", 20, 52, 404, 262);
        behaviorGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
        var behaviorPanel = new FlowLayoutPanel
        {
            Left = 16,
            Top = 28,
            Width = 368,
            Height = 214,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = Color.White
        };
        behaviorPanel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

        chkChinese = new CheckBox { Width = 330, Text = "默认中文模式" };
        chkFull = new CheckBox { Width = 330, Text = "默认全角模式" };
        chkChinesePunctuation = new CheckBox { Width = 330, Text = "中文模式下自动输出中文标点" };
        chkSmartSymbolPairs = new CheckBox { Width = 330, Text = "智能引号与书名号映射" };
        chkAutoCommitUniqueExact = new CheckBox { Width = 330, Text = "唯一精确候选继续输入时自动顶屏" };
        chkEmptyCandidateBeep = new CheckBox { Width = 330, Text = "无候选时给出一次提示音" };
        chkTabNavigation = new CheckBox { Width = 330, Text = "启用 Tab/Shift+Tab 候选导航" };
        chkEnterExactPriority = new CheckBox { Width = 330, Text = "回车优先精确候选，否则原码直出" };
        chkContextAssociationEnabled = new CheckBox { Width = 330, Text = "启用上下文学习持久化" };
        behaviorPanel.Controls.Add(chkChinese);
        behaviorPanel.Controls.Add(chkFull);
        behaviorPanel.Controls.Add(chkChinesePunctuation);
        behaviorPanel.Controls.Add(chkSmartSymbolPairs);
        behaviorPanel.Controls.Add(chkAutoCommitUniqueExact);
        behaviorPanel.Controls.Add(chkEmptyCandidateBeep);
        behaviorPanel.Controls.Add(chkTabNavigation);
        behaviorPanel.Controls.Add(chkEnterExactPriority);
        behaviorPanel.Controls.Add(chkContextAssociationEnabled);
        behaviorGroup.Controls.Add(behaviorPanel);

        var profileGroup = CreateGroupBox(tab, "词库与参数", 440, 52, 420, 262);
        profileGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        var lblPage = new Label { Left = 18, Top = 38, Width = 170, Text = "每页候选数" };
        numPage = new NumericUpDown { Left = 214, Top = 34, Width = 136, Minimum = ConfigDefaults.MinPageSize, Maximum = ConfigDefaults.MaxPageSize, Value = ConfigDefaults.DefaultPageSize };

        var lblHotkey = new Label { Left = 18, Top = 76, Width = 170, Text = "中英切换热键" };
        cmbHotkey = new ComboBox { Left = 214, Top = 72, Width = 136, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbHotkey.Items.Add("F9");
        cmbHotkey.Items.Add("F8");
        cmbHotkey.Items.Add("Ctrl+Space");
        cmbHotkey.SelectedIndex = 0;

        var lblProfile = new Label { Left = 18, Top = 114, Width = 170, Text = "码表模式" };
        cmbDictionaryProfile = new ComboBox { Left = 214, Top = 110, Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDictionaryProfile.Items.Add("zhengma-all");
        cmbDictionaryProfile.Items.Add("zhengma-large-pinyin");
        cmbDictionaryProfile.SelectedItem = ConfigDefaults.DefaultDictionaryProfile;

        var lblAutoCommitMin = new Label { Left = 18, Top = 152, Width = 170, Text = "自动顶屏最短码长" };
        numAutoCommitMinCodeLength = new NumericUpDown
        {
            Left = 214,
            Top = 148,
            Width = 136,
            Minimum = ConfigDefaults.MinAutoCommitCodeLength,
            Maximum = ConfigDefaults.MaxAutoCommitCodeLength,
            Value = ConfigDefaults.DefaultAutoCommitCodeLength
        };

        var lblAssocMax = new Label { Left = 18, Top = 190, Width = 170, Text = "上下文关联最大条目" };
        numContextAssociationMaxEntries = new NumericUpDown
        {
            Left = 214,
            Top = 186,
            Width = 136,
            Minimum = ConfigDefaults.MinContextAssociationMaxEntries,
            Maximum = ConfigDefaults.MaxContextAssociationMaxEntries,
            Value = ConfigDefaults.DefaultContextAssociationMaxEntries,
            Increment = 500
        };
        profileGroup.Controls.Add(lblPage);
        profileGroup.Controls.Add(numPage);
        profileGroup.Controls.Add(lblHotkey);
        profileGroup.Controls.Add(cmbHotkey);
        profileGroup.Controls.Add(lblProfile);
        profileGroup.Controls.Add(cmbDictionaryProfile);
        profileGroup.Controls.Add(lblAutoCommitMin);
        profileGroup.Controls.Add(numAutoCommitMinCodeLength);
        profileGroup.Controls.Add(lblAssocMax);
        profileGroup.Controls.Add(numContextAssociationMaxEntries);

        var btnOpenAutoPhraseFromGeneral = new Button
        {
            Left = 18,
            Top = 216,
            Width = 170,
            Height = 30,
            Text = "管理用户词典"
        };
        btnOpenAutoPhraseFromGeneral.Click += (s, e) => OpenAutoPhraseManagerDialog();

        var btnOpenPhraseReviewFromGeneral = new Button
        {
            Left = 198,
            Top = 216,
            Width = 186,
            Height = 30,
            Text = "打开造词审阅日志"
        };
        btnOpenPhraseReviewFromGeneral.Click += (s, e) => OpenPhraseReviewDialog();
        profileGroup.Controls.Add(btnOpenAutoPhraseFromGeneral);
        profileGroup.Controls.Add(btnOpenPhraseReviewFromGeneral);

        var strategyGroup = CreateGroupBox(tab, "当前策略说明", 20, 328, 840, 148);
        strategyGroup.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
        var txtStrategy = new TextBox
        {
            Left = 16,
            Top = 28,
            Width = 808,
            Height = 102,
            Multiline = true,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            Text = "1. 标点映射覆盖常用中英文符号、引号、书名号，连续录入时不需要再额外切换半角全角。\r\n2. 唯一精确候选支持按最短码长自动顶屏，目标是在快速拆字时减少一次确认动作。\r\n3. 回车策略可在 精确候选优先 与 当前高亮优先 之间切换，兼顾严谨录入和连续上屏两种习惯。\r\n4. 上下文学习会记录前后字搭配，帮助下一拍排序更贴近日常词序；用户可随时导出、裁剪、拉黑或重建。\r\n5. 用户词典同时承载手工置顶、自动造词和屏蔽词，主程序会延迟合并写回，尽量避开连续输入时段。"
        };
        txtStrategy.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        strategyGroup.Controls.Add(txtStrategy);
    }

    private void BuildDataTab(TabPage tab)
    {
        tab.BackColor = PanelBackColor;
        var hint = new Label { Left = 20, Top = 18, Width = 840, Height = 24, Text = "这里管理用户词典中的置顶词、自动造词和屏蔽词。列表按列展示，方便直接比对编码和词条。" };
        hint.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        tab.Controls.Add(hint);

        var grpPinned = CreateGroupBox(tab, "已置顶词条", 20, 52, 410, 422);
        var grpBlocked = CreateGroupBox(tab, "已屏蔽词条", 450, 52, 410, 422);
        grpPinned.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
        grpBlocked.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

        lstPinned = CreateStyledListBox(14, 28, 382, 304);
        lstBlocked = CreateStyledListBox(14, 28, 382, 304);
        lstPinned.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        lstBlocked.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        grpPinned.Controls.Add(lstPinned);
        grpBlocked.Controls.Add(lstBlocked);

        var btnRemovePinned = new Button { Left = 14, Top = 346, Width = 120, Text = "移除置顶" };
        var btnRefreshPinned = new Button { Left = 144, Top = 346, Width = 88, Text = "刷新" };
        var btnUnblock = new Button { Left = 14, Top = 346, Width = 120, Text = "取消屏蔽" };
        var btnRefreshBlocked = new Button { Left = 144, Top = 346, Width = 88, Text = "刷新" };
        grpPinned.Controls.Add(btnRemovePinned);
        grpPinned.Controls.Add(btnRefreshPinned);
        grpBlocked.Controls.Add(btnUnblock);
        grpBlocked.Controls.Add(btnRefreshBlocked);

        lblPinnedPath = new Label { Left = 14, Top = 382, Width = 382, Height = 32, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        lblBlockedPath = new Label { Left = 14, Top = 382, Width = 382, Height = 32, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        grpPinned.Controls.Add(lblPinnedPath);
        grpBlocked.Controls.Add(lblBlockedPath);

        btnRemovePinned.Click += (s, e) => RemovePinnedSelected();
        btnRefreshPinned.Click += (s, e) => RefreshDataLists();
        btnUnblock.Click += (s, e) => RemoveBlockedSelected();
        btnRefreshBlocked.Click += (s, e) => RefreshDataLists();

        var btnManageAutoPhrase = new Button { Left = 20, Top = 490, Width = 168, Height = 32, Text = "管理用户词典", Anchor = AnchorStyles.Left | AnchorStyles.Bottom };
        var btnOpenAutoPhraseFolder = new Button { Left = 198, Top = 490, Width = 168, Height = 32, Text = "打开用户词典目录" };
        btnOpenAutoPhraseFolder.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
        var btnOpenPhraseReview = new Button { Left = 376, Top = 490, Width = 156, Height = 32, Text = "打开造词审阅日志", Anchor = AnchorStyles.Left | AnchorStyles.Bottom };
        btnManageAutoPhrase.Click += (s, e) => OpenAutoPhraseManagerDialog();
        btnOpenAutoPhraseFolder.Click += (s, e) =>
        {
            string autoPhrasePath = ResolveAutoPhrasePath();
            Process.Start(new ProcessStartInfo { FileName = Path.GetDirectoryName(autoPhrasePath), UseShellExecute = true });
        };
        btnOpenPhraseReview.Click += (s, e) => OpenPhraseReviewDialog();
        tab.Controls.Add(btnManageAutoPhrase);
        tab.Controls.Add(btnOpenAutoPhraseFolder);
        tab.Controls.Add(btnOpenPhraseReview);

        tab.Controls.Add(new Label
        {
            Left = 20,
            Top = 532,
            Width = 840,
            Height = 24,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            Text = "这里只管理 yuninput_user.dict 相关的用户词条：手工置顶、自动造词和屏蔽词。系统词库本身不在此处直接编辑。"
        });

        tab.Controls.Add(new Label
        {
            Left = 548,
            Top = 496,
            Width = 312,
            Height = 36,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
            ForeColor = Color.FromArgb(97, 103, 112),
            Text = "用户词典文件: " + userDictPath + "\r\n自动造词文件: " + autoPhraseDictPath
        });
    }

    private string ResolveProductDataRoot()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string[] candidates = new[]
        {
            Path.Combine(baseDir, "data"),
            Path.GetFullPath(Path.Combine(baseDir, "..", "data")),
            Path.Combine(localRoot, "data")
        };

        foreach (string candidate in candidates)
        {
            if (Directory.Exists(candidate))
            {
                return candidate;
            }
        }

        return candidates[0];
    }

    private string ResolveAutoPhrasePath()
    {
        return autoPhraseDictPath;
    }

    private void OpenAutoPhraseManagerDialog()
    {
        string autoPath = ResolveAutoPhrasePath();
        Directory.CreateDirectory(Path.GetDirectoryName(userDictPath));
        Directory.CreateDirectory(Path.GetDirectoryName(autoPath));
        if (!File.Exists(userDictPath))
        {
            File.WriteAllText(userDictPath, string.Empty, Encoding.UTF8);
        }
        if (!File.Exists(autoPath))
        {
            File.WriteAllText(autoPath, string.Empty, Encoding.UTF8);
        }

        var workingEntries = ReadManagedDictionaryItems();
        bool dirty = false;

        var dialog = new Form
        {
            Text = "管理用户词典",
            ClientSize = new Size(880, 640),
            MinimumSize = new Size(860, 620),
            StartPosition = FormStartPosition.CenterParent,
            BackColor = PanelBackColor,
            Font = Font
        };
        EnableEscapeClose(dialog);

        var lblFilter = new Label { Left = 18, Top = 18, Width = 40, Height = 24, Text = "查询:" };
        var txtFilter = new TextBox { Left = 62, Top = 14, Width = 286, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        var lblSource = new Label { Left = 360, Top = 18, Width = 52, Height = 24, Text = "类型:" };
        var cmbSource = new ComboBox { Left = 414, Top = 14, Width = 134, DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Top | AnchorStyles.Right };
        cmbSource.Items.Add("全部词条");
        cmbSource.Items.Add("仅手工词条");
        cmbSource.Items.Add("仅自动造词");
        cmbSource.SelectedIndex = 0;
        var btnReload = new Button { Left = 562, Top = 12, Width = 78, Height = 30, Text = "重载", Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var btnSave = new Button { Left = 646, Top = 12, Width = 90, Height = 30, Text = "保存", Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var btnClose = new Button { Left = 742, Top = 12, Width = 120, Height = 30, Text = "关闭", Anchor = AnchorStyles.Top | AnchorStyles.Right, DialogResult = DialogResult.Cancel };

        var lstEntries = CreateStyledListBox(18, 54, 844, 378);
        lstEntries.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        lstEntries.SelectionMode = SelectionMode.MultiExtended;

        var editorGroup = new GroupBox
        {
            Left = 18,
            Top = 444,
            Width = 844,
            Height = 144,
            Text = "编辑词条",
            BackColor = Color.White,
            ForeColor = AccentColor,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };

        var lblCode = new Label { Left = 18, Top = 34, Width = 44, Height = 24, Text = "编码:" };
        var txtCode = new TextBox { Left = 66, Top = 30, Width = 186 };
        var lblText = new Label { Left = 270, Top = 34, Width = 44, Height = 24, Text = "词条:" };
        var txtText = new TextBox { Left = 318, Top = 30, Width = 186 };
        var lblType = new Label { Left = 522, Top = 34, Width = 44, Height = 24, Text = "类型:" };
        var cmbEntryType = new ComboBox { Left = 570, Top = 30, Width = 136, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbEntryType.Items.Add("手工用户词条");
        cmbEntryType.Items.Add("自动造词");
        cmbEntryType.SelectedIndex = 0;
        var lblScore = new Label { Left = 724, Top = 34, Width = 44, Height = 24, Text = "分值:" };
        var numScore = new NumericUpDown { Left = 772, Top = 30, Width = 54, Minimum = 1, Maximum = 9999, Value = 1 };

        var btnAdd = new Button { Left = 18, Top = 82, Width = 92, Height = 30, Text = "新增" };
        var btnUpdate = new Button { Left = 118, Top = 82, Width = 92, Height = 30, Text = "修改" };
        var btnDelete = new Button { Left = 218, Top = 82, Width = 92, Height = 30, Text = "删除选中" };
        var btnOpenManual = new Button { Left = 470, Top = 82, Width = 140, Height = 30, Text = "打开手工词典" };
        var btnOpenAuto = new Button { Left = 618, Top = 82, Width = 140, Height = 30, Text = "打开自动造词文件" };
        var btnOpenFolder = new Button { Left = 766, Top = 82, Width = 60, Height = 30, Text = "目录" };
        var lblStatus = new Label { Left = 18, Top = 116, Width = 808, Height = 22, ForeColor = Color.FromArgb(97, 103, 112) };

        editorGroup.Controls.Add(lblCode);
        editorGroup.Controls.Add(txtCode);
        editorGroup.Controls.Add(lblText);
        editorGroup.Controls.Add(txtText);
        editorGroup.Controls.Add(lblType);
        editorGroup.Controls.Add(cmbEntryType);
        editorGroup.Controls.Add(lblScore);
        editorGroup.Controls.Add(numScore);
        editorGroup.Controls.Add(btnAdd);
        editorGroup.Controls.Add(btnUpdate);
        editorGroup.Controls.Add(btnDelete);
        editorGroup.Controls.Add(btnOpenManual);
        editorGroup.Controls.Add(btnOpenAuto);
        editorGroup.Controls.Add(btnOpenFolder);
        editorGroup.Controls.Add(lblStatus);

        var lblPaths = new Label
        {
            Left = 18,
            Top = 596,
            Width = 844,
            Height = 36,
            ForeColor = Color.FromArgb(97, 103, 112),
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            Text = "手工词典: " + userDictPath + "\r\n自动造词: " + autoPath
        };

        dialog.Controls.Add(lblFilter);
        dialog.Controls.Add(txtFilter);
        dialog.Controls.Add(lblSource);
        dialog.Controls.Add(cmbSource);
        dialog.Controls.Add(btnReload);
        dialog.Controls.Add(btnSave);
        dialog.Controls.Add(btnClose);
        dialog.Controls.Add(lstEntries);
        dialog.Controls.Add(editorGroup);
        dialog.Controls.Add(lblPaths);

        Func<string, string, bool> validateInput = (code, text) =>
        {
            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show("请先填写编码和词条", "匀码输入法", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (code.IndexOfAny(new[] { ' ', '\t' }) >= 0 || text.IndexOfAny(new[] { ' ', '\t' }) >= 0)
            {
                MessageBox.Show("编码和词条不能包含空白字符", "匀码输入法", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        };

        Action refresh = () =>
        {
            string keyword = (txtFilter.Text ?? string.Empty).Trim();
            lstEntries.BeginUpdate();
            lstEntries.Items.Clear();
            int manualCount = 0;
            int autoCount = 0;

            foreach (CandidateEntryItem item in workingEntries)
            {
                if (item.IsAutoPhrase)
                {
                    autoCount++;
                }
                else
                {
                    manualCount++;
                }

                if (cmbSource.SelectedIndex == 1 && item.IsAutoPhrase)
                {
                    continue;
                }
                if (cmbSource.SelectedIndex == 2 && !item.IsAutoPhrase)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(keyword) &&
                    item.Code.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0 &&
                    item.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                lstEntries.Items.Add(item);
            }

            lstEntries.EndUpdate();
            lblStatus.Text = "当前共 " + workingEntries.Count + " 条，其中手工词条 " + manualCount + " 条，自动造词 " + autoCount + " 条" + (dirty ? "，有未保存修改" : string.Empty);
        };

        Action reload = () =>
        {
            workingEntries = ReadManagedDictionaryItems();
            dirty = false;
            refresh();
        };

        Action save = () =>
        {
            WriteManagedDictionaryItems(workingEntries);
            dirty = false;
            RefreshDataLists();
            refresh();
        };

        lstEntries.SelectedIndexChanged += (s, e) =>
        {
            if (lstEntries.SelectedItems.Count != 1)
            {
                txtCode.Text = string.Empty;
                txtText.Text = string.Empty;
                cmbEntryType.SelectedIndex = 0;
                numScore.Value = 1;
                return;
            }

            CandidateEntryItem item = lstEntries.SelectedItem as CandidateEntryItem;
            if (item == null)
            {
                return;
            }

            txtCode.Text = item.Code;
            txtText.Text = item.Text;
            cmbEntryType.SelectedIndex = item.IsAutoPhrase ? 1 : 0;
            numScore.Value = item.Score > 0 ? item.Score : 1;
        };

        btnAdd.Click += (s, e) =>
        {
            string code = (txtCode.Text ?? string.Empty).Trim().ToLowerInvariant();
            string text = (txtText.Text ?? string.Empty).Trim();
            if (!validateInput(code, text))
            {
                return;
            }

            bool isAuto = cmbEntryType.SelectedIndex == 1;
            foreach (CandidateEntryItem item in workingEntries)
            {
                if (string.Equals(item.Code, code, StringComparison.OrdinalIgnoreCase) && string.Equals(item.Text, text, StringComparison.Ordinal))
                {
                    item.IsAutoPhrase = isAuto;
                    item.Score = (int)numScore.Value;
                    dirty = true;
                    refresh();
                    return;
                }
            }

            workingEntries.Add(new CandidateEntryItem
            {
                Code = code,
                Text = text,
                IsAutoPhrase = isAuto,
                Score = (int)numScore.Value,
                RawLine = string.Empty
            });
            workingEntries.Sort((left, right) =>
            {
                if (left.IsAutoPhrase != right.IsAutoPhrase)
                {
                    return left.IsAutoPhrase ? 1 : -1;
                }

                int codeComparison = string.Compare(left.Code, right.Code, StringComparison.OrdinalIgnoreCase);
                if (codeComparison != 0)
                {
                    return codeComparison;
                }

                return string.Compare(left.Text, right.Text, StringComparison.Ordinal);
            });
            dirty = true;
            refresh();
        };

        btnUpdate.Click += (s, e) =>
        {
            if (lstEntries.SelectedItems.Count != 1)
            {
                MessageBox.Show("请先单选一条要修改的词条", "匀码输入法", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            CandidateEntryItem selected = lstEntries.SelectedItem as CandidateEntryItem;
            if (selected == null)
            {
                return;
            }

            string code = (txtCode.Text ?? string.Empty).Trim().ToLowerInvariant();
            string text = (txtText.Text ?? string.Empty).Trim();
            if (!validateInput(code, text))
            {
                return;
            }

            selected.Code = code;
            selected.Text = text;
            selected.IsAutoPhrase = cmbEntryType.SelectedIndex == 1;
            selected.Score = (int)numScore.Value;
            dirty = true;
            refresh();
        };

        btnDelete.Click += (s, e) =>
        {
            if (lstEntries.SelectedItems.Count == 0)
            {
                return;
            }

            var selectedItems = new List<CandidateEntryItem>();
            foreach (object selectedItem in lstEntries.SelectedItems)
            {
                CandidateEntryItem item = selectedItem as CandidateEntryItem;
                if (item != null)
                {
                    selectedItems.Add(item);
                }
            }

            foreach (CandidateEntryItem item in selectedItems)
            {
                workingEntries.Remove(item);
            }

            dirty = true;
            refresh();
        };

        txtFilter.TextChanged += (s, e) => refresh();
        cmbSource.SelectedIndexChanged += (s, e) => refresh();
        btnReload.Click += (s, e) => reload();
        btnSave.Click += (s, e) => save();
        btnOpenManual.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + userDictPath + "\"", UseShellExecute = true });
        btnOpenAuto.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + autoPath + "\"", UseShellExecute = true });
        btnOpenFolder.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = roamingRoot, UseShellExecute = true });
        btnClose.Click += (s, e) => dialog.Close();
        dialog.AcceptButton = btnSave;
        dialog.CancelButton = btnClose;
        dialog.FormClosing += (s, e) =>
        {
            if (!dirty)
            {
                return;
            }

            DialogResult result = MessageBox.Show("当前有未保存的用户词典修改，是否先保存？", "匀码输入法", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }
            if (result == DialogResult.Yes)
            {
                save();
            }
        };

        refresh();
        dialog.ShowDialog(this);
    }

    private static List<PhraseReviewLogItem> ReadPhraseReviewLogFile(string path)
    {
        var result = new List<PhraseReviewLogItem>();
        if (!File.Exists(path))
        {
            return result;
        }

        foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
            {
                continue;
            }

            string[] parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 4)
            {
                continue;
            }

            result.Add(new PhraseReviewLogItem
            {
                Code = parts[0],
                Text = parts[1],
                Source = parts[2],
                Time = parts[3],
                RawLine = line
            });
        }

        return result;
    }

    private static void RemoveLineFromFile(string path, string rawLine)
    {
        RemoveLinesFromFile(path, new[] { rawLine });
    }

    private static void RemoveLinesFromFile(string path, IEnumerable<string> rawLines)
    {
        var removalCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string rawLine in rawLines)
        {
            if (string.IsNullOrEmpty(rawLine))
            {
                continue;
            }

            int count;
            removalCounts.TryGetValue(rawLine, out count);
            removalCounts[rawLine] = count + 1;
        }

        var remaining = new List<string>();
        if (File.Exists(path))
        {
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                int count;
                if (removalCounts.TryGetValue(line, out count) && count > 0)
                {
                    removalCounts[line] = count - 1;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(line))
                {
                    remaining.Add(line);
                }
            }
        }

        Directory.CreateDirectory(Path.GetDirectoryName(path));
        File.WriteAllLines(path, remaining.ToArray(), Encoding.UTF8);
    }

    private static DateTime ParsePhraseReviewTime(string text)
    {
        DateTime parsed;
        if (DateTime.TryParseExact(
            text,
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
            out parsed))
        {
            return parsed;
        }

        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
        {
            return parsed.ToUniversalTime();
        }

        return DateTime.MinValue;
    }

    private void OpenPhraseReviewDialog()
    {
        string autoPhrasePath = ResolveAutoPhrasePath();
        Directory.CreateDirectory(Path.GetDirectoryName(autoPhrasePath));
        Directory.CreateDirectory(Path.GetDirectoryName(manualPhraseReviewPath));
        if (!File.Exists(autoPhrasePath))
        {
            File.WriteAllText(autoPhrasePath, string.Empty, Encoding.UTF8);
        }
        if (!File.Exists(manualPhraseReviewPath))
        {
            File.WriteAllText(manualPhraseReviewPath, string.Empty, Encoding.UTF8);
        }

        var dialog = new Form
        {
            Text = "造词审阅日志",
            ClientSize = new Size(980, 620),
            MinimumSize = new Size(940, 600),
            StartPosition = FormStartPosition.CenterParent,
            BackColor = PanelBackColor,
            Font = Font
        };
        EnableEscapeClose(dialog);

        var grpAuto = new GroupBox { Left = 16, Top = 16, Width = 456, Height = 520, Text = "自动造词快照（只读）", BackColor = Color.White, ForeColor = AccentColor, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left };
        var grpManual = new GroupBox { Left = 488, Top = 16, Width = 476, Height = 520, Text = "手工造词日志（manual_phrase_review.txt）", BackColor = Color.White, ForeColor = AccentColor, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
        dialog.Controls.Add(grpAuto);
        dialog.Controls.Add(grpManual);

        var lstAuto = CreateStyledListBox(14, 28, 428, 398);
        lstAuto.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        var lblSourceFilter = new Label { Left = 14, Top = 30, Width = 44, Height = 22, Text = "来源:" };
        var cmbSourceFilter = new ComboBox { Left = 62, Top = 26, Width = 136, DropDownStyle = ComboBoxStyle.DropDownList };
        var chkSortNewestFirst = new CheckBox { Left = 210, Top = 28, Width = 216, Text = "按时间倒序（最新在前）" };
        var lstManual = CreateStyledListBox(14, 58, 448, 368);
        lstManual.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        cmbSourceFilter.Items.Add("\u5168\u90e8\u6765\u6e90");
        cmbSourceFilter.Items.Add("\u4ec5\u81ea\u52a8\u9020\u8bcd");
        cmbSourceFilter.Items.Add("\u4ec5\u624b\u5de5\u9020\u8bcd");
        cmbSourceFilter.SelectedIndex = 0;
        chkSortNewestFirst.Checked = true;
        grpAuto.Controls.Add(lstAuto);
        grpManual.Controls.Add(lblSourceFilter);
        grpManual.Controls.Add(cmbSourceFilter);
        grpManual.Controls.Add(chkSortNewestFirst);
        grpManual.Controls.Add(lstManual);

        var lblAutoSummary = new Label { Left = 14, Top = 436, Width = 428, Height = 24, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        var lblAutoPath = new Label { Left = 14, Top = 462, Width = 428, Height = 40, Text = "文件: " + autoPhrasePath, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        var lblManualPath = new Label { Left = 14, Top = 436, Width = 448, Height = 40, Text = "文件: " + manualPhraseReviewPath, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        grpAuto.Controls.Add(lblAutoSummary);
        grpAuto.Controls.Add(lblAutoPath);
        grpManual.Controls.Add(lblManualPath);

        var btnManageDictionary = new Button { Left = 14, Top = 470, Width = 134, Height = 30, Text = "管理用户词典", Anchor = AnchorStyles.Left | AnchorStyles.Bottom };
        var btnOpenAutoFile = new Button { Left = 156, Top = 470, Width = 134, Height = 30, Text = "打开自动造词文件", Anchor = AnchorStyles.Left | AnchorStyles.Bottom };
        var btnManualDelete = new Button { Left = 14, Top = 476, Width = 118, Height = 30, Text = "删除选中日志", Anchor = AnchorStyles.Left | AnchorStyles.Bottom };
        var btnOpenLog = new Button { Left = 140, Top = 476, Width = 120, Height = 30, Text = "打开日志文件", Anchor = AnchorStyles.Left | AnchorStyles.Bottom };
        var btnRefresh = new Button { Left = 776, Top = 548, Width = 86, Height = 32, Text = "刷新", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnClose = new Button { Left = 878, Top = 548, Width = 86, Height = 32, Text = "关闭", Anchor = AnchorStyles.Right | AnchorStyles.Bottom, DialogResult = DialogResult.Cancel };
        grpAuto.Controls.Add(btnManageDictionary);
        grpAuto.Controls.Add(btnOpenAutoFile);
        grpManual.Controls.Add(btnManualDelete);
        grpManual.Controls.Add(btnOpenLog);
        dialog.Controls.Add(btnRefresh);
        dialog.Controls.Add(btnClose);

        Action refresh = () =>
        {
            lstAuto.BeginUpdate();
            lstAuto.Items.Clear();
            int autoCount = 0;
            foreach (CandidateEntryItem item in ReadEntryFile(autoPhrasePath))
            {
                if (!item.IsAutoPhrase)
                {
                    continue;
                }

                lstAuto.Items.Add(item);
                autoCount++;
            }
            lstAuto.EndUpdate();
            lblAutoSummary.Text = "当前自动造词快照共 " + autoCount + " 条。编辑和保存请回到“管理用户词典”。";

            lstManual.BeginUpdate();
            lstManual.Items.Clear();

            string selectedSource = string.Empty;
            if (cmbSourceFilter.SelectedIndex == 1)
            {
                selectedSource = "auto_phrase";
            }
            else if (cmbSourceFilter.SelectedIndex == 2)
            {
                selectedSource = "manual_hotkey";
            }

            var manualItems = new List<PhraseReviewLogItem>();
            foreach (PhraseReviewLogItem item in ReadPhraseReviewLogFile(manualPhraseReviewPath))
            {
                if (!string.IsNullOrEmpty(selectedSource) && !string.Equals(item.Source, selectedSource, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                manualItems.Add(item);
            }

            manualItems.Sort((a, b) => ParsePhraseReviewTime(a.Time).CompareTo(ParsePhraseReviewTime(b.Time)));
            if (chkSortNewestFirst.Checked)
            {
                manualItems.Reverse();
            }

            foreach (PhraseReviewLogItem item in manualItems)
            {
                lstManual.Items.Add(item);
            }
            lstManual.EndUpdate();
        };

        btnManualDelete.Click += (s, e) =>
        {
            var selectedLines = new List<string>();
            foreach (object selectedItem in lstManual.SelectedItems)
            {
                PhraseReviewLogItem item = selectedItem as PhraseReviewLogItem;
                if (item != null && !string.IsNullOrEmpty(item.RawLine))
                {
                    selectedLines.Add(item.RawLine);
                }
            }

            if (selectedLines.Count == 0)
            {
                return;
            }

            RemoveLinesFromFile(manualPhraseReviewPath, selectedLines);
            refresh();
        };

        btnRefresh.Click += (s, e) => refresh();
        cmbSourceFilter.SelectedIndexChanged += (s, e) => refresh();
        chkSortNewestFirst.CheckedChanged += (s, e) => refresh();
        btnManageDictionary.Click += (s, e) => OpenAutoPhraseManagerDialog();
        btnOpenAutoFile.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + autoPhrasePath + "\"", UseShellExecute = true });
        btnOpenLog.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + manualPhraseReviewPath + "\"", UseShellExecute = true });
        btnClose.Click += (s, e) => dialog.Close();
        dialog.CancelButton = btnClose;

        refresh();
        dialog.ShowDialog(this);
    }

    private void BuildContextTab(TabPage tab)
    {
        tab.BackColor = PanelBackColor;
        var hint = new Label { Left = 20, Top = 18, Width = 840, Height = 24, Text = "上下文学习会影响候选排序。你可以在这里查看高频搭配、导出、裁剪、拉黑或重建。" };
        hint.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        tab.Controls.Add(hint);

        var grpAssoc = CreateGroupBox(tab, "上下文关联学习条目", 20, 52, 560, 434);
        var grpBlacklist = CreateGroupBox(tab, "上下文黑名单", 600, 52, 260, 434);
        grpAssoc.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
        grpBlacklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

        lstContextAssoc = CreateStyledListBox(14, 28, 532, 286);
        lstContextAssoc.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        grpAssoc.Controls.Add(lstContextAssoc);

        lstContextAssocBlacklist = CreateStyledListBox(14, 28, 232, 286);
        lstContextAssocBlacklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        grpBlacklist.Controls.Add(lstContextAssocBlacklist);

        var btnRefreshAssoc = new Button { Left = 14, Top = 326, Width = 72, Text = "刷新" };
        var btnClearAssoc = new Button { Left = 92, Top = 326, Width = 72, Text = "清空" };
        var btnExportAssoc = new Button { Left = 170, Top = 326, Width = 72, Text = "导出" };
        var btnTrimAssoc = new Button { Left = 248, Top = 326, Width = 90, Text = "按分裁剪" };
        var lblTrim = new Label { Left = 344, Top = 331, Width = 54, Text = "最低分" };
        var numTrimMinScore = new NumericUpDown { Left = 400, Top = 326, Width = 74, Minimum = 1, Maximum = 999999, Value = 2 };
        var btnBlacklistSelected = new Button { Left = 14, Top = 362, Width = 120, Text = "加入黑名单" };
        var btnRebuildAssoc = new Button { Left = 140, Top = 362, Width = 126, Text = "一键重建关联" };
        grpAssoc.Controls.Add(btnRefreshAssoc);
        grpAssoc.Controls.Add(btnClearAssoc);
        grpAssoc.Controls.Add(btnExportAssoc);
        grpAssoc.Controls.Add(btnTrimAssoc);
        grpAssoc.Controls.Add(lblTrim);
        grpAssoc.Controls.Add(numTrimMinScore);
        grpAssoc.Controls.Add(btnBlacklistSelected);
        grpAssoc.Controls.Add(btnRebuildAssoc);

        var btnRefreshBlacklist = new Button { Left = 14, Top = 326, Width = 72, Text = "刷新" };
        var btnRemoveBlacklist = new Button { Left = 92, Top = 326, Width = 120, Text = "移出黑名单" };
        grpBlacklist.Controls.Add(btnRefreshBlacklist);
        grpBlacklist.Controls.Add(btnRemoveBlacklist);

        lblContextAssocPath = new Label { Left = 14, Top = 396, Width = 532, Height = 32, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        lblContextAssocBlacklistPath = new Label { Left = 14, Top = 396, Width = 232, Height = 32, ForeColor = Color.FromArgb(97, 103, 112), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
        grpAssoc.Controls.Add(lblContextAssocPath);
        grpBlacklist.Controls.Add(lblContextAssocBlacklistPath);

        btnRefreshAssoc.Click += (s, e) => RefreshContextAssocLists();
        btnRefreshBlacklist.Click += (s, e) => RefreshContextAssocLists();

        btnClearAssoc.Click += (s, e) =>
        {
            if (MessageBox.Show("\u786e\u5b9a\u6e05\u7a7a\u6240\u6709\u4e0a\u4e0b\u6587\u5173\u8054\u5b66\u4e60\u6570\u636e\u5417\uff1f", "\u786e\u8ba4", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            File.WriteAllText(contextAssocPath, string.Empty, Encoding.UTF8);
            RefreshContextAssocLists();
        };

        btnExportAssoc.Click += (s, e) => ExportContextAssoc();

        btnTrimAssoc.Click += (s, e) =>
        {
            TrimContextAssocByMinScore((ulong)numTrimMinScore.Value);
            RefreshContextAssocLists();
        };

        btnBlacklistSelected.Click += (s, e) =>
        {
            var item = lstContextAssoc.SelectedItem as ContextAssocItem;
            if (item == null)
            {
                return;
            }

            AddContextAssocToBlacklist(item);
            RefreshContextAssocLists();
        };

        btnRebuildAssoc.Click += (s, e) =>
        {
            RebuildContextAssocFromUserData();
            RefreshContextAssocLists();
        };

        btnRemoveBlacklist.Click += (s, e) =>
        {
            var item = lstContextAssocBlacklist.SelectedItem as CandidateEntryItem;
            if (item == null)
            {
                return;
            }

            WriteFilteredFile(contextAssocBlacklistPath, item);
            RefreshContextAssocLists();
        };

        tab.Controls.Add(new Label
        {
            Left = 20,
            Top = 500,
            Width = 840,
            Height = 24,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            Text = "说明：可从 user_dict 与 user_freq 一键重建关联；黑名单中的前后字组合不会再参与提权。"
        });
    }

    private static bool ParseBool(string text, string key, bool fallback)
    {
        string probe = "\"" + key + "\":";
        int idx = text.IndexOf(probe, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
        {
            return fallback;
        }

        string tail = text.Substring(idx + probe.Length).TrimStart();
        if (tail.StartsWith("true", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        if (tail.StartsWith("false", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }
        return fallback;
    }

    private static int ParseInt(string text, string key, int fallback)
    {
        string probe = "\"" + key + "\":";
        int idx = text.IndexOf(probe, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
        {
            return fallback;
        }

        string tail = text.Substring(idx + probe.Length);
        var digits = new StringBuilder();
        foreach (char ch in tail)
        {
            if (char.IsDigit(ch))
            {
                digits.Append(ch);
            }
            else if (digits.Length > 0)
            {
                break;
            }
        }

        int value;
        return int.TryParse(digits.ToString(), out value) ? value : fallback;
    }

    private static string ParseString(string text, string key, string fallback)
    {
        string probe = "\"" + key + "\":";
        int idx = text.IndexOf(probe, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
        {
            return fallback;
        }

        string tail = text.Substring(idx + probe.Length).TrimStart();
        if (!tail.StartsWith("\"", StringComparison.Ordinal))
        {
            return fallback;
        }

        int end = tail.IndexOf('"', 1);
        if (end <= 0)
        {
            return fallback;
        }

        return tail.Substring(1, end - 1);
    }

    private void LoadConfig()
    {
        if (!File.Exists(cfgPath))
        {
            SaveConfig(false);
            return;
        }

        string text = File.ReadAllText(cfgPath, Encoding.UTF8);
        chkChinese.Checked = ParseBool(text, "chinese_mode", true);
        chkFull.Checked = ParseBool(text, "full_shape", false);
        chkChinesePunctuation.Checked = ParseBool(text, "chinese_punctuation", ConfigDefaults.DefaultChinesePunctuation);
        chkSmartSymbolPairs.Checked = ParseBool(text, "smart_symbol_pairs", ConfigDefaults.DefaultSmartSymbolPairs);
        chkAutoCommitUniqueExact.Checked = ParseBool(text, "auto_commit_unique_exact", ConfigDefaults.DefaultAutoCommitUniqueExact);
        chkEmptyCandidateBeep.Checked = ParseBool(text, "empty_candidate_beep", ConfigDefaults.DefaultEmptyCandidateBeep);
        chkTabNavigation.Checked = ParseBool(text, "tab_navigation", ConfigDefaults.DefaultTabNavigation);
        chkEnterExactPriority.Checked = ParseBool(text, "enter_exact_priority", ConfigDefaults.DefaultEnterExactPriority);
        chkContextAssociationEnabled.Checked = ParseBool(text, "context_association_enabled", ConfigDefaults.DefaultContextAssociationEnabled);

        int pageValue = ParseInt(text, "candidate_page_size", (int)ConfigDefaults.DefaultPageSize);
        if (pageValue < ConfigDefaults.MinPageSize) pageValue = (int)ConfigDefaults.MinPageSize;
        if (pageValue > ConfigDefaults.MaxPageSize) pageValue = (int)ConfigDefaults.MaxPageSize;
        numPage.Value = pageValue;

        int autoCommitLength = ParseInt(text, "auto_commit_min_code_length", (int)ConfigDefaults.DefaultAutoCommitCodeLength);
        if (autoCommitLength < ConfigDefaults.MinAutoCommitCodeLength) autoCommitLength = (int)ConfigDefaults.MinAutoCommitCodeLength;
        if (autoCommitLength > ConfigDefaults.MaxAutoCommitCodeLength) autoCommitLength = (int)ConfigDefaults.MaxAutoCommitCodeLength;
        numAutoCommitMinCodeLength.Value = autoCommitLength;

        int contextAssocMaxEntries = ParseInt(text, "context_association_max_entries", (int)ConfigDefaults.DefaultContextAssociationMaxEntries);
        if (contextAssocMaxEntries < ConfigDefaults.MinContextAssociationMaxEntries) contextAssocMaxEntries = (int)ConfigDefaults.MinContextAssociationMaxEntries;
        if (contextAssocMaxEntries > ConfigDefaults.MaxContextAssociationMaxEntries) contextAssocMaxEntries = (int)ConfigDefaults.MaxContextAssociationMaxEntries;
        numContextAssociationMaxEntries.Value = contextAssocMaxEntries;

        string hotkey = ParseString(text, "toggle_hotkey", "F9");
        if (hotkey == "Ctrl+Space") cmbHotkey.SelectedItem = "Ctrl+Space";
        else if (hotkey == "F8") cmbHotkey.SelectedItem = "F8";
        else cmbHotkey.SelectedItem = "F9";

        string dictionaryProfile = ParseString(text, "dictionary_profile", ConfigDefaults.DefaultDictionaryProfile);
        if (dictionaryProfile == "zhengma-large-pinyin" || dictionaryProfile == "pinyin") cmbDictionaryProfile.SelectedItem = "zhengma-large-pinyin";
        else if (dictionaryProfile == "zhengma-all" || dictionaryProfile == "zhengma-large") cmbDictionaryProfile.SelectedItem = ConfigDefaults.DefaultDictionaryProfile;
        else cmbDictionaryProfile.SelectedItem = ConfigDefaults.DefaultDictionaryProfile;
    }

    private void SaveConfig(bool showMessage)
    {
        var sb = new StringBuilder();
        sb.AppendLine("{");
        sb.AppendLine("  \"chinese_mode\": " + (chkChinese.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"full_shape\": " + (chkFull.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"chinese_punctuation\": " + (chkChinesePunctuation.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"smart_symbol_pairs\": " + (chkSmartSymbolPairs.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"auto_commit_unique_exact\": " + (chkAutoCommitUniqueExact.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"auto_commit_min_code_length\": " + numAutoCommitMinCodeLength.Value + ",");
        sb.AppendLine("  \"empty_candidate_beep\": " + (chkEmptyCandidateBeep.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"tab_navigation\": " + (chkTabNavigation.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"enter_exact_priority\": " + (chkEnterExactPriority.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"context_association_enabled\": " + (chkContextAssociationEnabled.Checked ? "true" : "false") + ",");
        sb.AppendLine("  \"context_association_max_entries\": " + numContextAssociationMaxEntries.Value + ",");
        sb.AppendLine("  \"candidate_page_size\": " + numPage.Value + ",");
        sb.AppendLine("  \"dictionary_profile\": \"" + cmbDictionaryProfile.SelectedItem + "\",");
        sb.AppendLine("  \"toggle_hotkey\": \"" + cmbHotkey.SelectedItem + "\"");
        sb.AppendLine("}");
        File.WriteAllText(cfgPath, sb.ToString(), Encoding.UTF8);

        if (showMessage)
        {
            MessageBox.Show("\u914d\u7f6e\u5df2\u4fdd\u5b58", "\u5300\u7801\u8f93\u5165\u6cd5", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static List<CandidateEntryItem> ReadEntryFile(string path)
    {
        var result = new List<CandidateEntryItem>();
        if (!File.Exists(path))
        {
            return result;
        }

        foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
            {
                continue;
            }

            string[] parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
            {
                continue;
            }

            result.Add(new CandidateEntryItem
            {
                Code = parts[0],
                Text = parts[1],
                RawLine = line,
                IsAutoPhrase = HasAutoPhraseTag(parts),
                Score = ParseEntryScore(parts)
            });
        }

        return result;
    }

    private static int ParseEntryScore(string[] parts)
    {
        if (parts == null || parts.Length < 3)
        {
            return 1;
        }

        int parsed;
        if (int.TryParse(parts[2], out parsed) && parsed > 0)
        {
            return parsed;
        }

        return 1;
    }

    private static bool HasAutoPhraseTag(string[] parts)
    {
        for (int i = 2; i < parts.Length; i++)
        {
            if (string.Equals(parts[i], "auto", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static List<ContextAssocItem> ReadContextAssocFile(string path)
    {
        var result = new List<ContextAssocItem>();
        if (!File.Exists(path))
        {
            return result;
        }

        foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
            {
                continue;
            }

            string[] parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
            {
                continue;
            }

            ulong score;
            if (!ulong.TryParse(parts[2], out score))
            {
                continue;
            }

            result.Add(new ContextAssocItem
            {
                PrevText = parts[0],
                NextText = parts[1],
                Score = score
            });
        }

        result.Sort((a, b) => b.Score.CompareTo(a.Score));
        return result;
    }

    private static List<Tuple<string, string, ulong>> ReadWeightedEntriesFile(string path)
    {
        var result = new List<Tuple<string, string, ulong>>();
        if (!File.Exists(path))
        {
            return result;
        }

        foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
            {
                continue;
            }

            string[] parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
            {
                continue;
            }

            ulong weight = 1;
            if (parts.Length >= 3)
            {
                ulong parsed;
                if (ulong.TryParse(parts[2], out parsed) && parsed > 0)
                {
                    weight = parsed;
                }
            }

            result.Add(Tuple.Create(parts[0], parts[1], weight));
        }

        return result;
    }

    private void RefreshDataLists()
    {
        lstPinned.BeginUpdate();
        lstBlocked.BeginUpdate();
        lstPinned.Items.Clear();
        lstBlocked.Items.Clear();

        foreach (CandidateEntryItem item in ReadEntryFile(userDictPath))
        {
            if (item.IsAutoPhrase)
            {
                continue;
            }

            lstPinned.Items.Add(item);
        }

        foreach (CandidateEntryItem item in ReadEntryFile(blockedPath))
        {
            lstBlocked.Items.Add(item);
        }

        lstPinned.EndUpdate();
        lstBlocked.EndUpdate();

        lblPinnedPath.Text = "\u6587\u4ef6: " + userDictPath;
        lblBlockedPath.Text = "\u6587\u4ef6: " + blockedPath;
    }

    private void RefreshContextAssocLists()
    {
        if (lstContextAssoc == null || lstContextAssocBlacklist == null)
        {
            return;
        }

        lstContextAssoc.BeginUpdate();
        lstContextAssoc.Items.Clear();
        foreach (ContextAssocItem item in ReadContextAssocFile(contextAssocPath))
        {
            lstContextAssoc.Items.Add(item);
        }
        lstContextAssoc.EndUpdate();

        lstContextAssocBlacklist.BeginUpdate();
        lstContextAssocBlacklist.Items.Clear();
        foreach (CandidateEntryItem item in ReadEntryFile(contextAssocBlacklistPath))
        {
            lstContextAssocBlacklist.Items.Add(item);
        }
        lstContextAssocBlacklist.EndUpdate();

        lblContextAssocPath.Text = "\u6587\u4ef6: " + contextAssocPath;
        lblContextAssocBlacklistPath.Text = "\u6587\u4ef6: " + contextAssocBlacklistPath;
    }

    private void ExportContextAssoc()
    {
        var items = ReadContextAssocFile(contextAssocPath);
        if (items.Count == 0)
        {
            MessageBox.Show("\u6ca1\u6709\u53ef\u5bfc\u51fa\u7684\u4e0a\u4e0b\u6587\u5b66\u4e60\u6570\u636e", "\u5300\u7801\u8f93\u5165\u6cd5", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using (var dialog = new SaveFileDialog())
        {
            dialog.Title = "\u5bfc\u51fa\u4e0a\u4e0b\u6587\u5b66\u4e60\u6570\u636e";
            dialog.FileName = "context_assoc_export.txt";
            dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            var lines = new List<string>();
            foreach (ContextAssocItem item in items)
            {
                lines.Add(item.PrevText + " " + item.NextText + " " + item.Score);
            }
            File.WriteAllLines(dialog.FileName, lines.ToArray(), Encoding.UTF8);
        }
    }

    private void TrimContextAssocByMinScore(ulong minScore)
    {
        var items = ReadContextAssocFile(contextAssocPath);
        var lines = new List<string>();
        foreach (ContextAssocItem item in items)
        {
            if (item.Score >= minScore)
            {
                lines.Add(item.PrevText + " " + item.NextText + " " + item.Score);
            }
        }

        Directory.CreateDirectory(Path.GetDirectoryName(contextAssocPath));
        File.WriteAllLines(contextAssocPath, lines.ToArray(), Encoding.UTF8);
    }

    private void AddContextAssocToBlacklist(ContextAssocItem item)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(contextAssocBlacklistPath));
        var existing = new HashSet<string>(StringComparer.Ordinal);
        foreach (CandidateEntryItem blocked in ReadEntryFile(contextAssocBlacklistPath))
        {
            existing.Add(blocked.Code + "\t" + blocked.Text);
        }

        string key = item.PrevText + "\t" + item.NextText;
        if (!existing.Contains(key))
        {
            File.AppendAllText(contextAssocBlacklistPath, item.PrevText + " " + item.NextText + Environment.NewLine, Encoding.UTF8);
        }

        var kept = new List<string>();
        foreach (ContextAssocItem assoc in ReadContextAssocFile(contextAssocPath))
        {
            if (assoc.PrevText == item.PrevText && assoc.NextText == item.NextText)
            {
                continue;
            }
            kept.Add(assoc.PrevText + " " + assoc.NextText + " " + assoc.Score);
        }
        File.WriteAllLines(contextAssocPath, kept.ToArray(), Encoding.UTF8);
    }

    private void RebuildContextAssocFromUserData()
    {
        if (MessageBox.Show("\u786e\u5b9a\u57fa\u4e8e user_dict \u548c user_freq \u91cd\u5efa\u4e0a\u4e0b\u6587\u5173\u8054\u6570\u636e\u5417\uff1f", "\u786e\u8ba4", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
        {
            return;
        }

        int skippedInvalidPairs;
        int skippedBlacklistedPairs;
        int rebuiltCount = RebuildContextAssocFromUserDataCore(false, out skippedInvalidPairs, out skippedBlacklistedPairs);
        MessageBox.Show(
            "\u5df2\u91cd\u5efa\u4e0a\u4e0b\u6587\u5173\u8054\u6761\u76ee: " + rebuiltCount + Environment.NewLine +
            "\u5df2\u8fc7\u6ee4\u5f02\u5e38\u7ec4\u5408: " + skippedInvalidPairs + Environment.NewLine +
            "\u5df2\u8df3\u8fc7\u9ed1\u540d\u5355\u7ec4\u5408: " + skippedBlacklistedPairs,
            "\u5300\u7801\u8f93\u5165\u6cd5",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private int RebuildContextAssocFromUserDataCore(bool dryRun, out int skippedInvalidPairs, out int skippedBlacklistedPairs)
    {
        skippedInvalidPairs = 0;
        skippedBlacklistedPairs = 0;
        int localSkippedInvalidPairs = 0;
        int localSkippedBlacklistedPairs = 0;
        var blacklist = new HashSet<string>(StringComparer.Ordinal);
        foreach (CandidateEntryItem item in ReadEntryFile(contextAssocBlacklistPath))
        {
            blacklist.Add(item.Code + "\t" + item.Text);
        }

        var scoreMap = new Dictionary<string, ulong>(StringComparer.Ordinal);
        Action<string, string, ulong> addScore = (prevText, nextText, weight) =>
        {
            if (string.IsNullOrWhiteSpace(prevText) || string.IsNullOrWhiteSpace(nextText) || weight == 0)
            {
                return;
            }

            if (!IsValidAssocToken(prevText) || !IsValidAssocToken(nextText))
            {
                localSkippedInvalidPairs++;
                return;
            }

            string key = prevText + "\t" + nextText;
            if (blacklist.Contains(key))
            {
                localSkippedBlacklistedPairs++;
                return;
            }

            ulong existing;
            if (!scoreMap.TryGetValue(key, out existing))
            {
                scoreMap[key] = weight;
            }
            else
            {
                scoreMap[key] = existing + weight;
            }
        };

        Action<string, ulong> digestText = (text, weight) =>
        {
            if (string.IsNullOrWhiteSpace(text) || text.Length < 2 || weight == 0)
            {
                return;
            }

            for (int i = 0; i + 1 < text.Length; i++)
            {
                string prev = text[i].ToString();
                string next = text[i + 1].ToString();
                addScore(prev, next, weight);
            }
        };

        foreach (var tuple in ReadWeightedEntriesFile(userDictPath))
        {
            digestText(tuple.Item2, tuple.Item3 < 1 ? 1UL : tuple.Item3);
        }

        foreach (var tuple in ReadWeightedEntriesFile(userFreqPath))
        {
            digestText(tuple.Item2, tuple.Item3 < 1 ? 1UL : tuple.Item3);
        }

        var ranked = new List<ContextAssocItem>();
        foreach (var pair in scoreMap)
        {
            string[] parts = pair.Key.Split('\t');
            if (parts.Length != 2)
            {
                continue;
            }

            ranked.Add(new ContextAssocItem
            {
                PrevText = parts[0],
                NextText = parts[1],
                Score = pair.Value
            });
        }

        ranked.Sort((a, b) => b.Score.CompareTo(a.Score));
        int maxEntries = (int)numContextAssociationMaxEntries.Value;
        if (ranked.Count > maxEntries)
        {
            ranked = ranked.GetRange(0, maxEntries);
        }

        if (dryRun)
        {
            var previewLines = new List<string>();
            previewLines.Add("dry_run=true");
            previewLines.Add("generated_at=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            previewLines.Add("entries=" + ranked.Count);
            previewLines.Add("max_entries=" + maxEntries);
            previewLines.Add("source_files=yuninput_user.dict,user_freq.txt");
            previewLines.Add("skipped_invalid_pairs=" + localSkippedInvalidPairs);
            previewLines.Add("skipped_blacklisted_pairs=" + localSkippedBlacklistedPairs);
            previewLines.Add("preview_top=30");
            int limit = ranked.Count < 30 ? ranked.Count : 30;
            for (int i = 0; i < limit; i++)
            {
                ContextAssocItem item = ranked[i];
                previewLines.Add(item.PrevText + " " + item.NextText + " " + item.Score);
            }

            Directory.CreateDirectory(Path.GetDirectoryName(contextAssocDryRunReportPath));
            File.WriteAllLines(contextAssocDryRunReportPath, previewLines.ToArray(), Encoding.UTF8);
            skippedInvalidPairs = localSkippedInvalidPairs;
            skippedBlacklistedPairs = localSkippedBlacklistedPairs;
            return ranked.Count;
        }

        var lines = new List<string>();
        foreach (ContextAssocItem item in ranked)
        {
            lines.Add(item.PrevText + " " + item.NextText + " " + item.Score);
        }

        Directory.CreateDirectory(Path.GetDirectoryName(contextAssocPath));
        File.WriteAllLines(contextAssocPath, lines.ToArray(), Encoding.UTF8);
        skippedInvalidPairs = localSkippedInvalidPairs;
        skippedBlacklistedPairs = localSkippedBlacklistedPairs;
        return ranked.Count;
    }

    private static bool IsValidAssocToken(string token)
    {
        if (string.IsNullOrEmpty(token) || token.Length != 1)
        {
            return false;
        }

        char ch = token[0];
        if (char.IsControl(ch) || char.IsWhiteSpace(ch) || ch == '\uFFFD')
        {
            return false;
        }

        UnicodeCategory category = char.GetUnicodeCategory(ch);
        if (category == UnicodeCategory.Surrogate || category == UnicodeCategory.PrivateUse || category == UnicodeCategory.OtherNotAssigned)
        {
            return false;
        }

        return true;
    }

    public static int RunContextAssocRebuildHeadless(bool dryRun)
    {
        using (var form = new ConfigForm())
        {
            int skippedInvalidPairs;
            int skippedBlacklistedPairs;
            return form.RebuildContextAssocFromUserDataCore(dryRun, out skippedInvalidPairs, out skippedBlacklistedPairs);
        }
    }

    private void RemovePinnedSelected()
    {
        CandidateEntryItem item = lstPinned.SelectedItem as CandidateEntryItem;
        if (item == null)
        {
            return;
        }

        WriteFilteredFile(userDictPath, item);
        RefreshDataLists();
    }

    private void RemoveBlockedSelected()
    {
        CandidateEntryItem item = lstBlocked.SelectedItem as CandidateEntryItem;
        if (item == null)
        {
            return;
        }

        WriteFilteredFile(blockedPath, item);
        RefreshDataLists();
    }

    private static void WriteFilteredFile(string path, CandidateEntryItem removedItem)
    {
        var remaining = new List<string>();
        if (File.Exists(path))
        {
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                string[] parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2 && parts[0] == removedItem.Code && parts[1] == removedItem.Text)
                {
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(line))
                {
                    remaining.Add(line);
                }
            }
        }

        Directory.CreateDirectory(Path.GetDirectoryName(path));
        File.WriteAllLines(path, remaining.ToArray(), Encoding.UTF8);
    }

    private static void RemoveEntryItemsFromFile(string path, IEnumerable<CandidateEntryItem> removedItems)
    {
        var rawLines = new List<string>();
        foreach (CandidateEntryItem item in removedItems)
        {
            if (item != null && !string.IsNullOrEmpty(item.RawLine))
            {
                rawLines.Add(item.RawLine);
            }
        }

        if (rawLines.Count == 0)
        {
            return;
        }

        RemoveLinesFromFile(path, rawLines);
    }
}

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        bool runRebuild = false;
        bool dryRun = false;

        if (args != null)
        {
            foreach (string arg in args)
            {
                if (string.Equals(arg, "--rebuild-context-assoc", StringComparison.OrdinalIgnoreCase))
                {
                    runRebuild = true;
                }

                if (string.Equals(arg, "--dry-run", StringComparison.OrdinalIgnoreCase))
                {
                    dryRun = true;
                }
            }
        }

        if (runRebuild)
        {
            ConfigForm.RunContextAssocRebuildHeadless(dryRun);
            return;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new ConfigForm());
    }
}
