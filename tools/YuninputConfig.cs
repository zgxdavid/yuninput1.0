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
        Width = 860;
        Height = 600;
        MinimumSize = new Size(860, 600);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

        localRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "yuninput");
        roamingRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "yuninput");
        Directory.CreateDirectory(localRoot);
        Directory.CreateDirectory(roamingRoot);

        cfgPath = Path.Combine(localRoot, "settings.json");
        userDictPath = Path.Combine(roamingRoot, "yuninput_user.dict");
        autoPhraseDictPath = userDictPath;
        userFreqPath = Path.Combine(roamingRoot, "user_freq.txt");
        blockedPath = Path.Combine(roamingRoot, "blocked_entries.txt");
        contextAssocPath = Path.Combine(roamingRoot, "context_assoc.txt");
        contextAssocBlacklistPath = Path.Combine(roamingRoot, "context_assoc_blacklist.txt");
        contextAssocDryRunReportPath = Path.Combine(roamingRoot, "context_assoc_dry_run_report.txt");
        manualPhraseReviewPath = Path.Combine(roamingRoot, "manual_phrase_review.txt");

        var tabs = new TabControl();
        tabs.Left = 16;
        tabs.Top = 16;
        tabs.Width = 812;
        tabs.Height = 480;
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

        var btnSave = new Button { Left = 462, Top = 514, Width = 86, Text = "\u4fdd\u5b58", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnRefresh = new Button { Left = 558, Top = 514, Width = 86, Text = "\u5237\u65b0", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnOpenFolder = new Button { Left = 654, Top = 514, Width = 86, Text = "\u6253\u5f00\u76ee\u5f55", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        var btnClose = new Button { Left = 750, Top = 514, Width = 78, Text = "\u5173\u95ed", Anchor = AnchorStyles.Right | AnchorStyles.Bottom };

        btnSave.Click += (s, e) => SaveConfig(true);
        btnRefresh.Click += (s, e) =>
        {
            LoadConfig();
            RefreshDataLists();
            RefreshContextAssocLists();
        };
        btnOpenFolder.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = roamingRoot, UseShellExecute = true });
        btnClose.Click += (s, e) => Close();

        Controls.Add(btnSave);
        Controls.Add(btnRefresh);
        Controls.Add(btnOpenFolder);
        Controls.Add(btnClose);

        LoadConfig();
        RefreshDataLists();
        RefreshContextAssocLists();
    }

    private void BuildGeneralTab(TabPage tab)
    {
        var lblIntro = new Label { Left = 20, Top = 18, Width = 720, Text = "\u8fd9\u91cc\u63a7\u5236\u9ed8\u8ba4\u6a21\u5f0f\u3001\u5019\u9009\u884c\u4e3a\u548c\u4e2d\u6587\u6807\u70b9\u7b56\u7565\u3002" };
        tab.Controls.Add(lblIntro);

        chkChinese = new CheckBox { Left = 24, Top = 56, Width = 220, Text = "\u9ed8\u8ba4\u4e2d\u6587\u6a21\u5f0f" };
        chkFull = new CheckBox { Left = 24, Top = 88, Width = 220, Text = "\u9ed8\u8ba4\u5168\u89d2\u6a21\u5f0f" };
        chkChinesePunctuation = new CheckBox { Left = 24, Top = 120, Width = 320, Text = "\u4e2d\u6587\u6a21\u5f0f\u4e0b\u81ea\u52a8\u8f93\u51fa\u4e2d\u6587\u6807\u70b9" };
        chkSmartSymbolPairs = new CheckBox { Left = 24, Top = 152, Width = 320, Text = "\u667a\u80fd\u5f15\u53f7\u4e0e\u4e66\u540d\u53f7\u6620\u5c04" };
        chkAutoCommitUniqueExact = new CheckBox { Left = 24, Top = 184, Width = 360, Text = "\u552f\u4e00\u7cbe\u786e\u5019\u9009\u5728\u7ee7\u7eed\u8f93\u5165\u65f6\u81ea\u52a8\u9876\u5c4f" };
        chkEmptyCandidateBeep = new CheckBox { Left = 24, Top = 216, Width = 360, Text = "\u65e0\u5019\u9009\u65f6\u7ed9\u51fa\u4e00\u6b21\u63d0\u793a\u97f3" };
        chkTabNavigation = new CheckBox { Left = 24, Top = 248, Width = 360, Text = "\u542f\u7528 Tab/Shift+Tab \u5019\u9009\u5bfc\u822a" };
        chkEnterExactPriority = new CheckBox { Left = 24, Top = 280, Width = 360, Text = "\u56de\u8f66\u952e\u4f18\u5148\u7cbe\u786e\u5019\u9009\u5426\u5219\u539f\u7801\u76f4\u51fa" };
        chkContextAssociationEnabled = new CheckBox { Left = 24, Top = 312, Width = 360, Text = "\u542f\u7528\u4e0a\u4e0b\u6587\u5b66\u4e60\u6301\u4e45\u5316" };

        tab.Controls.Add(chkChinese);
        tab.Controls.Add(chkFull);
        tab.Controls.Add(chkChinesePunctuation);
        tab.Controls.Add(chkSmartSymbolPairs);
        tab.Controls.Add(chkAutoCommitUniqueExact);
        tab.Controls.Add(chkEmptyCandidateBeep);
        tab.Controls.Add(chkTabNavigation);
        tab.Controls.Add(chkEnterExactPriority);
        tab.Controls.Add(chkContextAssociationEnabled);

        tab.Controls.Add(new Label { Left = 420, Top = 56, Width = 170, Text = "\u6bcf\u9875\u5019\u9009\u6570" });
        numPage = new NumericUpDown { Left = 598, Top = 52, Width = 90, Minimum = ConfigDefaults.MinPageSize, Maximum = ConfigDefaults.MaxPageSize, Value = ConfigDefaults.DefaultPageSize };
        tab.Controls.Add(numPage);

        tab.Controls.Add(new Label { Left = 420, Top = 92, Width = 170, Text = "\u4e2d\u82f1\u5207\u6362\u70ed\u952e" });
        cmbHotkey = new ComboBox { Left = 598, Top = 88, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbHotkey.Items.Add("F9");
        cmbHotkey.Items.Add("F8");
        cmbHotkey.Items.Add("Ctrl+Space");
        cmbHotkey.SelectedIndex = 0;
        tab.Controls.Add(cmbHotkey);

        tab.Controls.Add(new Label { Left = 420, Top = 124, Width = 170, Text = "\u7801\u8868\u6a21\u5f0f" });
        cmbDictionaryProfile = new ComboBox { Left = 598, Top = 120, Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDictionaryProfile.Items.Add("zhengma-all");
        cmbDictionaryProfile.Items.Add("zhengma-large-pinyin");
        cmbDictionaryProfile.SelectedItem = ConfigDefaults.DefaultDictionaryProfile;
        tab.Controls.Add(cmbDictionaryProfile);

        tab.Controls.Add(new Label { Left = 420, Top = 156, Width = 170, Text = "\u81ea\u52a8\u9876\u5c4f\u6700\u77ed\u7801\u957f" });
        numAutoCommitMinCodeLength = new NumericUpDown
        {
            Left = 598,
            Top = 152,
            Width = 90,
            Minimum = ConfigDefaults.MinAutoCommitCodeLength,
            Maximum = ConfigDefaults.MaxAutoCommitCodeLength,
            Value = ConfigDefaults.DefaultAutoCommitCodeLength
        };
        tab.Controls.Add(numAutoCommitMinCodeLength);

        tab.Controls.Add(new Label { Left = 420, Top = 192, Width = 170, Text = "\u4e0a\u4e0b\u6587\u5173\u8054\u6700\u5927\u6761\u76ee" });
        numContextAssociationMaxEntries = new NumericUpDown
        {
            Left = 598,
            Top = 188,
            Width = 110,
            Minimum = ConfigDefaults.MinContextAssociationMaxEntries,
            Maximum = ConfigDefaults.MaxContextAssociationMaxEntries,
            Value = ConfigDefaults.DefaultContextAssociationMaxEntries,
            Increment = 500
        };
        tab.Controls.Add(numContextAssociationMaxEntries);

        var btnOpenAutoPhraseFromGeneral = new Button
        {
            Left = 420,
            Top = 232,
            Width = 300,
            Height = 30,
            Text = "\u7ba1\u7406\u81ea\u9020\u8bcd\uff08yuninput_user.dict\uff09"
        };
        btnOpenAutoPhraseFromGeneral.Click += (s, e) => OpenAutoPhraseManagerDialog();
        tab.Controls.Add(btnOpenAutoPhraseFromGeneral);

        var btnOpenPhraseReviewFromGeneral = new Button
        {
            Left = 420,
            Top = 268,
            Width = 300,
            Height = 30,
            Text = "\u6253\u5f00\u9020\u8bcd\u96c6\u4e2d\u5ba1\u9605\u9762\u677f"
        };
        btnOpenPhraseReviewFromGeneral.Click += (s, e) => OpenPhraseReviewDialog();
        tab.Controls.Add(btnOpenPhraseReviewFromGeneral);

        tab.Controls.Add(new Label
        {
            Left = 24,
            Top = 348,
            Width = 720,
            Height = 128,
            Text = "\u7b56\u7565\u8bf4\u660e:\r\n- \u6807\u70b9\u6620\u5c04\u63a7\u5236 , . / ; : ! ? \u4e0e\u5f15\u53f7\u3001\u4e66\u540d\u53f7\u7684\u8f93\u51fa\r\n- \u552f\u4e00\u7cbe\u786e\u5019\u9009\u81ea\u52a8\u9876\u5c4f\u53ef\u964d\u4f4e\u8fde\u7eed\u8f93\u5165\u6309\u952e\u6210\u672c\r\n- \u56de\u8f66\u7b56\u7565\u53ef\u9009\uff1a\u4f18\u5148\u7cbe\u786e\u5019\u9009\u6216\u4f7f\u7528\u5f53\u524d\u9ad8\u4eae\u5019\u9009\r\n- Ctrl+1-9 \u53ef\u7f6e\u9876\u5019\u9009\uff0cCtrl+Delete \u53ef\u5c4f\u853d\u5f53\u524d\u5019\u9009"
        });
    }

    private void BuildDataTab(TabPage tab)
    {
        var grpPinned = new GroupBox { Left = 18, Top = 18, Width = 370, Height = 400, Text = "\u5df2\u7f6e\u9876\u8bcd\u6761" };
        var grpBlocked = new GroupBox { Left = 402, Top = 18, Width = 370, Height = 400, Text = "\u5df2\u5c4f\u853d\u8bcd\u6761" };
        tab.Controls.Add(grpPinned);
        tab.Controls.Add(grpBlocked);

        lstPinned = new ListBox { Left = 14, Top = 26, Width = 340, Height = 296 };
        lstBlocked = new ListBox { Left = 14, Top = 26, Width = 340, Height = 296 };
        grpPinned.Controls.Add(lstPinned);
        grpBlocked.Controls.Add(lstBlocked);

        var btnRemovePinned = new Button { Left = 14, Top = 334, Width = 120, Text = "\u79fb\u9664\u7f6e\u9876" };
        var btnRefreshPinned = new Button { Left = 144, Top = 334, Width = 88, Text = "\u5237\u65b0" };
        var btnUnblock = new Button { Left = 14, Top = 334, Width = 120, Text = "\u53d6\u6d88\u5c4f\u853d" };
        var btnRefreshBlocked = new Button { Left = 144, Top = 334, Width = 88, Text = "\u5237\u65b0" };
        grpPinned.Controls.Add(btnRemovePinned);
        grpPinned.Controls.Add(btnRefreshPinned);
        grpBlocked.Controls.Add(btnUnblock);
        grpBlocked.Controls.Add(btnRefreshBlocked);

        lblPinnedPath = new Label { Left = 14, Top = 366, Width = 340, Height = 26 };
        lblBlockedPath = new Label { Left = 14, Top = 366, Width = 340, Height = 26 };
        grpPinned.Controls.Add(lblPinnedPath);
        grpBlocked.Controls.Add(lblBlockedPath);

        btnRemovePinned.Click += (s, e) => RemovePinnedSelected();
        btnRefreshPinned.Click += (s, e) => RefreshDataLists();
        btnUnblock.Click += (s, e) => RemoveBlockedSelected();
        btnRefreshBlocked.Click += (s, e) => RefreshDataLists();

        var btnManageAutoPhrase = new Button { Left = 20, Top = 458, Width = 180, Text = "\u7ba1\u7406\u81ea\u542f\u9020\u8bcd\uff08\u81ea\u52a8\u9020\u8bcd\uff09" };
        var btnOpenAutoPhraseFolder = new Button { Left = 208, Top = 458, Width = 164, Text = "\u6253\u5f00\u7528\u6237\u8bcd\u5178\u76ee\u5f55" };
        var btnOpenPhraseReview = new Button { Left = 380, Top = 458, Width = 130, Text = "\u9020\u8bcd\u96c6\u4e2d\u5ba1\u9605" };
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
            Top = 432,
            Width = 740,
            Height = 24,
            Text = "\u8fd9\u91cc\u53ea\u7ba1\u7406 yuninput_user.dict \u91cc\u7684\u7528\u6237\u8bcd\u6761\uff1a\u624b\u5de5\u7f6e\u9876\u8bcd\u3001\u81ea\u52a8\u9020\u8bcd\u548c\u5c4f\u853d\u8bcd\u3002\u914d\u7f6e\u5de5\u5177\u4e0d\u5bf9\u7cfb\u7edf\u8bcd\u5e93 zhengma-all.dict \u6216 yuninput_user-extend.dict \u63d0\u4f9b\u589e\u5220\u6539\u67e5\u3002"
        });

        tab.Controls.Add(new Label
        {
            Left = 450,
            Top = 462,
            Width = 320,
            Height = 20,
            Text = "\u7528\u6237\u8bcd\u5178\u6587\u4ef6: " + userDictPath
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
        return userDictPath;
    }

    private void OpenAutoPhraseManagerDialog()
    {
        string path = ResolveAutoPhrasePath();
        string directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (!File.Exists(path))
        {
            var initLines = new[]
            {
                "# yuninput_user.dict",
                "# format: code phrase [optional_score] [optional_source_tag:auto]",
                "# delete unwanted lines and save"
            };
            File.WriteAllLines(path, initLines, Encoding.UTF8);
        }

        var dialog = new Form
        {
            Text = "自造词管理（yuninput_user.dict）",
            Width = 760,
            Height = 560,
            StartPosition = FormStartPosition.CenterParent,
            MinimizeBox = false,
            MaximizeBox = false
        };

        var lblFilter = new Label { Left = 16, Top = 18, Width = 58, Height = 22, Text = "查询:" };
        var txtFilter = new TextBox { Left = 72, Top = 14, Width = 488 };
        var btnRefresh = new Button { Left = 568, Top = 12, Width = 80, Text = "刷新" };
        var btnOpenFile = new Button { Left = 654, Top = 12, Width = 88, Text = "打开文件" };
        var lstAuto = new ListBox { Left = 16, Top = 44, Width = 726, Height = 338, SelectionMode = SelectionMode.MultiExtended };

        var lblCode = new Label { Left = 16, Top = 392, Width = 58, Height = 22, Text = "编码:" };
        var txtCode = new TextBox { Left = 72, Top = 388, Width = 224 };
        var lblText = new Label { Left = 312, Top = 392, Width = 58, Height = 22, Text = "词条:" };
        var txtText = new TextBox { Left = 364, Top = 388, Width = 222 };

        var btnAdd = new Button { Left = 16, Top = 424, Width = 104, Text = "新增" };
        var btnUpdate = new Button { Left = 126, Top = 424, Width = 104, Text = "修改" };
        var btnDelete = new Button { Left = 236, Top = 424, Width = 104, Text = "删除" };
        var btnClose = new Button { Left = 638, Top = 486, Width = 104, Text = "关闭" };

        var lblPath = new Label { Left = 16, Top = 458, Width = 726, Height = 22, Text = "文件: " + path };

        dialog.Controls.Add(lblFilter);
        dialog.Controls.Add(txtFilter);
        dialog.Controls.Add(btnRefresh);
        dialog.Controls.Add(btnOpenFile);
        dialog.Controls.Add(lstAuto);
        dialog.Controls.Add(lblCode);
        dialog.Controls.Add(txtCode);
        dialog.Controls.Add(lblText);
        dialog.Controls.Add(txtText);
        dialog.Controls.Add(btnAdd);
        dialog.Controls.Add(btnUpdate);
        dialog.Controls.Add(btnDelete);
        dialog.Controls.Add(lblPath);
        dialog.Controls.Add(btnClose);

        Action refresh = () =>
        {
            string keyword = (txtFilter.Text ?? string.Empty).Trim();
            lstAuto.BeginUpdate();
            lstAuto.Items.Clear();
            foreach (CandidateEntryItem item in ReadEntryFile(path))
            {
                if (!item.IsAutoPhrase)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(keyword) &&
                    item.Code.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0 &&
                    item.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                lstAuto.Items.Add(item);
            }
            lstAuto.EndUpdate();
        };

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

        lstAuto.SelectedIndexChanged += (s, e) =>
        {
            if (lstAuto.SelectedItems.Count != 1)
            {
                txtCode.Text = string.Empty;
                txtText.Text = string.Empty;
                return;
            }

            CandidateEntryItem item = lstAuto.SelectedItem as CandidateEntryItem;
            if (item == null)
            {
                return;
            }

            txtCode.Text = item.Code;
            txtText.Text = item.Text;
        };

        btnAdd.Click += (s, e) =>
        {
            string code = (txtCode.Text ?? string.Empty).Trim().ToLowerInvariant();
            string text = (txtText.Text ?? string.Empty).Trim();
            if (!validateInput(code, text))
            {
                return;
            }

            bool exists = false;
            foreach (CandidateEntryItem item in ReadEntryFile(path))
            {
                if (string.Equals(item.Code, code, StringComparison.OrdinalIgnoreCase) && string.Equals(item.Text, text, StringComparison.Ordinal))
                {
                    exists = true;
                    break;
                }
            }

            if (!exists)
            {
                File.AppendAllText(path, code + " " + text + " 1 auto" + Environment.NewLine, Encoding.UTF8);
            }

            refresh();
        };

        btnUpdate.Click += (s, e) =>
        {
            if (lstAuto.SelectedItems.Count != 1)
            {
                MessageBox.Show("请先单选一条要修改的词条", "匀码输入法", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            CandidateEntryItem selected = lstAuto.SelectedItem as CandidateEntryItem;
            if (selected == null)
            {
                MessageBox.Show("请先选择要修改的词条", "匀码输入法", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string newCode = (txtCode.Text ?? string.Empty).Trim().ToLowerInvariant();
            string newText = (txtText.Text ?? string.Empty).Trim();
            if (!validateInput(newCode, newText))
            {
                return;
            }

            string newLine = newCode + " " + newText + " 1 auto";
            var lines = new List<string>();
            bool replaced = false;
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (!replaced)
                {
                    string[] parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2 &&
                        string.Equals(parts[0], selected.Code, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(parts[1], selected.Text, StringComparison.Ordinal))
                    {
                        lines.Add(newLine);
                        replaced = true;
                        continue;
                    }
                }

                lines.Add(line);
            }

            if (!replaced)
            {
                lines.Add(newLine);
            }

            File.WriteAllLines(path, lines.ToArray(), Encoding.UTF8);
            refresh();
        };

        btnDelete.Click += (s, e) =>
        {
            var selectedItems = new List<CandidateEntryItem>();
            foreach (object selectedItem in lstAuto.SelectedItems)
            {
                CandidateEntryItem item = selectedItem as CandidateEntryItem;
                if (item != null)
                {
                    selectedItems.Add(item);
                }
            }

            if (selectedItems.Count == 0)
            {
                return;
            }

            RemoveEntryItemsFromFile(path, selectedItems);
            refresh();
        };

        txtFilter.TextChanged += (s, e) => refresh();
        btnRefresh.Click += (s, e) => refresh();
        btnOpenFile.Click += (s, e) => Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + path + "\"", UseShellExecute = true });
        btnClose.Click += (s, e) => dialog.Close();

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
            Text = "\u9020\u8bcd\u5ba1\u9605\u9762\u677f",
            Width = 920,
            Height = 560,
            StartPosition = FormStartPosition.CenterParent,
            MinimizeBox = false,
            MaximizeBox = false
        };

        var grpAuto = new GroupBox { Left = 14, Top = 12, Width = 430, Height = 458, Text = "\u81ea\u52a8\u9020\u8bcd\uff08\u5b58\u4e8e yuninput_user.dict\uff09" };
        var grpManual = new GroupBox { Left = 456, Top = 12, Width = 430, Height = 458, Text = "\u624b\u5de5\u9020\u8bcd\u65e5\u5fd7\uff08manual_phrase_review.txt\uff09" };
        dialog.Controls.Add(grpAuto);
        dialog.Controls.Add(grpManual);

        var lstAuto = new ListBox { Left = 12, Top = 24, Width = 404, Height = 300, SelectionMode = SelectionMode.MultiExtended };
        var lstManual = new ListBox { Left = 12, Top = 56, Width = 404, Height = 328, SelectionMode = SelectionMode.MultiExtended };
        var lblAutoCode = new Label { Left = 12, Top = 332, Width = 36, Height = 22, Text = "\u7801:" };
        var txtAutoCode = new TextBox { Left = 52, Top = 328, Width = 120 };
        var lblAutoText = new Label { Left = 180, Top = 332, Width = 36, Height = 22, Text = "\u8bcd:" };
        var txtAutoText = new TextBox { Left = 220, Top = 328, Width = 126 };
        var btnAutoAdd = new Button { Left = 352, Top = 327, Width = 64, Text = "\u6dfb\u52a0" };
        var lblSourceFilter = new Label { Left = 12, Top = 28, Width = 52, Height = 20, Text = "\u6765\u6e90:" };
        var cmbSourceFilter = new ComboBox { Left = 68, Top = 24, Width = 132, DropDownStyle = ComboBoxStyle.DropDownList };
        var chkSortNewestFirst = new CheckBox { Left = 214, Top = 26, Width = 198, Text = "\u6309\u65f6\u95f4\u5012\u5e8f\uff08\u6700\u65b0\u5728\u524d\uff09" };
        cmbSourceFilter.Items.Add("\u5168\u90e8\u6765\u6e90");
        cmbSourceFilter.Items.Add("\u4ec5\u81ea\u52a8\u9020\u8bcd");
        cmbSourceFilter.Items.Add("\u4ec5\u624b\u5de5\u9020\u8bcd");
        cmbSourceFilter.SelectedIndex = 0;
        chkSortNewestFirst.Checked = true;
        grpAuto.Controls.Add(lstAuto);
        grpAuto.Controls.Add(lblAutoCode);
        grpAuto.Controls.Add(txtAutoCode);
        grpAuto.Controls.Add(lblAutoText);
        grpAuto.Controls.Add(txtAutoText);
        grpAuto.Controls.Add(btnAutoAdd);
        grpManual.Controls.Add(lblSourceFilter);
        grpManual.Controls.Add(cmbSourceFilter);
        grpManual.Controls.Add(chkSortNewestFirst);
        grpManual.Controls.Add(lstManual);

        var lblAutoPath = new Label { Left = 12, Top = 388, Width = 404, Height = 28, Text = "\u6587\u4ef6: " + autoPhrasePath };
        var lblManualPath = new Label { Left = 12, Top = 388, Width = 404, Height = 28, Text = "\u6587\u4ef6: " + manualPhraseReviewPath };
        grpAuto.Controls.Add(lblAutoPath);
        grpManual.Controls.Add(lblManualPath);

        var btnAutoDelete = new Button { Left = 12, Top = 422, Width = 110, Text = "\u5220\u9664\u9009\u4e2d\u81ea\u52a8\u8bcd" };
        var btnManualDelete = new Button { Left = 12, Top = 422, Width = 110, Text = "\u5220\u9664\u9009\u4e2d\u65e5\u5fd7" };
        var btnRefresh = new Button { Left = 680, Top = 482, Width = 86, Text = "\u5237\u65b0" };
        var btnOpenFiles = new Button { Left = 736, Top = 482, Width = 150, Text = "\u6253\u5f00\u7528\u6237\u8bcd\u5178\u4e0e\u65e5\u5fd7" };
        var btnClose = new Button { Left = 588, Top = 482, Width = 86, Text = "\u5173\u95ed" };
        grpAuto.Controls.Add(btnAutoDelete);
        grpManual.Controls.Add(btnManualDelete);
        dialog.Controls.Add(btnRefresh);
        dialog.Controls.Add(btnOpenFiles);
        dialog.Controls.Add(btnClose);

        Action refresh = () =>
        {
            lstAuto.BeginUpdate();
            lstAuto.Items.Clear();
            foreach (CandidateEntryItem item in ReadEntryFile(autoPhrasePath))
            {
                if (!item.IsAutoPhrase)
                {
                    continue;
                }

                lstAuto.Items.Add(item);
            }
            lstAuto.EndUpdate();

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

        btnAutoDelete.Click += (s, e) =>
        {
            var selectedItems = new List<CandidateEntryItem>();
            foreach (object selectedItem in lstAuto.SelectedItems)
            {
                CandidateEntryItem item = selectedItem as CandidateEntryItem;
                if (item != null)
                {
                    selectedItems.Add(item);
                }
            }

            if (selectedItems.Count == 0)
            {
                return;
            }

            RemoveEntryItemsFromFile(autoPhrasePath, selectedItems);
            refresh();
        };

        lstAuto.SelectedIndexChanged += (s, e) =>
        {
            if (lstAuto.SelectedItems.Count != 1)
            {
                txtAutoCode.Text = string.Empty;
                txtAutoText.Text = string.Empty;
                return;
            }

            CandidateEntryItem item = lstAuto.SelectedItem as CandidateEntryItem;
            if (item == null)
            {
                return;
            }

            txtAutoCode.Text = item.Code;
            txtAutoText.Text = item.Text;
        };

        btnAutoAdd.Click += (s, e) =>
        {
            string code = (txtAutoCode.Text ?? string.Empty).Trim().ToLowerInvariant();
            string text = (txtAutoText.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show("\u8bf7\u5148\u586b\u5199\u7f16\u7801\u548c\u8bcd\u6761", "\u5300\u7801\u8f93\u5165\u6cd5", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (code.IndexOfAny(new[] { ' ', '\t' }) >= 0 || text.IndexOfAny(new[] { ' ', '\t' }) >= 0)
            {
                MessageBox.Show("\u7f16\u7801\u548c\u8bcd\u6761\u4e0d\u80fd\u5305\u542b\u7a7a\u767d\u5b57\u7b26", "\u5300\u7801\u8f93\u5165\u6cd5", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool exists = false;
            foreach (CandidateEntryItem item in ReadEntryFile(autoPhrasePath))
            {
                if (!item.IsAutoPhrase)
                {
                    continue;
                }

                if (string.Equals(item.Code, code, StringComparison.OrdinalIgnoreCase) && string.Equals(item.Text, text, StringComparison.Ordinal))
                {
                    exists = true;
                    break;
                }
            }
            if (!exists)
            {
                File.AppendAllText(autoPhrasePath, code + " " + text + " 1 auto" + Environment.NewLine, Encoding.UTF8);
            }

            refresh();
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
        btnOpenFiles.Click += (s, e) =>
        {
            Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + autoPhrasePath + "\"", UseShellExecute = true });
            Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + manualPhraseReviewPath + "\"", UseShellExecute = true });
        };
        btnClose.Click += (s, e) => dialog.Close();

        refresh();
        dialog.ShowDialog(this);
    }

    private void BuildContextTab(TabPage tab)
    {
        var grpAssoc = new GroupBox { Left = 18, Top = 18, Width = 516, Height = 412, Text = "\u4e0a\u4e0b\u6587\u5173\u8054\u5b66\u4e60\u6761\u76ee" };
        var grpBlacklist = new GroupBox { Left = 544, Top = 18, Width = 228, Height = 412, Text = "\u4e0a\u4e0b\u6587\u9ed1\u540d\u5355" };
        tab.Controls.Add(grpAssoc);
        tab.Controls.Add(grpBlacklist);

        lstContextAssoc = new ListBox { Left = 14, Top = 26, Width = 486, Height = 276 };
        grpAssoc.Controls.Add(lstContextAssoc);

        lstContextAssocBlacklist = new ListBox { Left = 14, Top = 26, Width = 198, Height = 276 };
        grpBlacklist.Controls.Add(lstContextAssocBlacklist);

        var btnRefreshAssoc = new Button { Left = 14, Top = 312, Width = 72, Text = "\u5237\u65b0" };
        var btnClearAssoc = new Button { Left = 92, Top = 312, Width = 72, Text = "\u6e05\u7a7a" };
        var btnExportAssoc = new Button { Left = 170, Top = 312, Width = 72, Text = "\u5bfc\u51fa" };
        var btnTrimAssoc = new Button { Left = 248, Top = 312, Width = 90, Text = "\u6309\u5206\u88c1\u526a" };
        var lblTrim = new Label { Left = 344, Top = 317, Width = 54, Text = "\u6700\u4f4e\u5206" };
        var numTrimMinScore = new NumericUpDown { Left = 400, Top = 312, Width = 64, Minimum = 1, Maximum = 999999, Value = 2 };
        var btnBlacklistSelected = new Button { Left = 14, Top = 344, Width = 120, Text = "\u52a0\u5165\u9ed1\u540d\u5355" };
        var btnRebuildAssoc = new Button { Left = 140, Top = 344, Width = 126, Text = "\u4e00\u952e\u91cd\u5efa\u5173\u8054" };
        grpAssoc.Controls.Add(btnRefreshAssoc);
        grpAssoc.Controls.Add(btnClearAssoc);
        grpAssoc.Controls.Add(btnExportAssoc);
        grpAssoc.Controls.Add(btnTrimAssoc);
        grpAssoc.Controls.Add(lblTrim);
        grpAssoc.Controls.Add(numTrimMinScore);
        grpAssoc.Controls.Add(btnBlacklistSelected);
        grpAssoc.Controls.Add(btnRebuildAssoc);

        var btnRefreshBlacklist = new Button { Left = 14, Top = 312, Width = 72, Text = "\u5237\u65b0" };
        var btnRemoveBlacklist = new Button { Left = 92, Top = 312, Width = 120, Text = "\u79fb\u51fa\u9ed1\u540d\u5355" };
        grpBlacklist.Controls.Add(btnRefreshBlacklist);
        grpBlacklist.Controls.Add(btnRemoveBlacklist);

        lblContextAssocPath = new Label { Left = 14, Top = 374, Width = 486, Height = 30 };
        lblContextAssocBlacklistPath = new Label { Left = 14, Top = 374, Width = 198, Height = 30 };
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
            Top = 438,
            Width = 740,
            Height = 24,
            Text = "\u8bf4\u660e\uff1a\u53ef\u4ece user_dict/user_freq \u4e00\u952e\u91cd\u5efa\u5173\u8054\uff0c\u9ed1\u540d\u5355\u7ec4\u5408\u4e0d\u53c2\u4e0e\u63d0\u6743\u3002"
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
                IsAutoPhrase = HasAutoPhraseTag(parts)
            });
        }

        return result;
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
