$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$projectRoot = Split-Path -Parent $PSScriptRoot
$buildRoot = Join-Path $projectRoot "dist\installer_build"
$payloadDir = Join-Path $buildRoot "payload"
$payloadZip = Join-Path $buildRoot "payload.zip"
$sourceCs = Join-Path $buildRoot "GuiInstaller.cs"
$outExe = Join-Path (Split-Path -Parent $projectRoot) "yuninput_setup.exe"
$configBuildScript = Join-Path $projectRoot "scripts\build_config_app.ps1"
$configExe = Join-Path $projectRoot "tools\yuninput_config.exe"
$iconPath = Join-Path $projectRoot "assets\icon_yun.ico"

$dllPath = Join-Path $projectRoot "build\Release\yuninput.dll"
if (-not (Test-Path $dllPath)) {
    throw "Missing build output: $dllPath"
}

if (-not (Test-Path $configExe) -and (Test-Path $configBuildScript)) {
    & $configBuildScript
}

if (Test-Path $buildRoot) {
    Remove-Item -Recurse -Force $buildRoot
}

New-Item -ItemType Directory -Path (Join-Path $payloadDir "build\Release") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $payloadDir "scripts") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $payloadDir "data") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $payloadDir "tools") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $payloadDir "assets") | Out-Null

Copy-Item (Join-Path $projectRoot "build\Release\yuninput.dll") (Join-Path $payloadDir "build\Release\yuninput.dll")
Copy-Item (Join-Path $projectRoot "scripts\register_ime.ps1") (Join-Path $payloadDir "scripts\register_ime.ps1")
Copy-Item (Join-Path $projectRoot "scripts\unregister_ime.ps1") (Join-Path $payloadDir "scripts\unregister_ime.ps1")
Copy-Item (Join-Path $projectRoot "scripts\install_enable.ps1") (Join-Path $payloadDir "scripts\install_enable.ps1")
Copy-Item (Join-Path $projectRoot "scripts\uninstall_clean.ps1") (Join-Path $payloadDir "scripts\uninstall_clean.ps1")
Copy-Item (Join-Path $projectRoot "scripts\build_config_app.ps1") (Join-Path $payloadDir "scripts\build_config_app.ps1")
Copy-Item (Join-Path $projectRoot "data\yuninput_basic.dict") (Join-Path $payloadDir "data\yuninput_basic.dict")
if (Test-Path (Join-Path $projectRoot "data\yuninput_user.dict")) {
    Copy-Item (Join-Path $projectRoot "data\yuninput_user.dict") (Join-Path $payloadDir "data\yuninput_user.dict")
}
if (Test-Path $configExe) {
    Copy-Item $configExe (Join-Path $payloadDir "tools\yuninput_config.exe")
}
if (Test-Path $iconPath) {
    Copy-Item $iconPath (Join-Path $payloadDir "assets\icon_yun.ico")
}

$setupCmdPath = Join-Path $payloadDir "setup.cmd"
$setupCmdLines = @(
    "@echo off",
    "setlocal",
    'powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\install_enable.ps1"',
    "exit /b %errorlevel%"
)
$setupCmdLines | Set-Content -Encoding ascii $setupCmdPath

if (Test-Path $payloadZip) {
    Remove-Item -Force $payloadZip
}
Compress-Archive -Path (Join-Path $payloadDir "*") -DestinationPath $payloadZip -Force

$b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($payloadZip))

$template = @"
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

public class InstallerForm : Form
{
    private CheckBox chkChineseMode;
    private CheckBox chkFullShape;
    private NumericUpDown numCandidateCount;
    private ComboBox cmbToggleHotkey;
    private TextBox txtLog;
    private Button btnInstall;
    private Button btnCancel;
    private bool installCompleted;

    public InstallerForm()
    {
        Text = "\u5300\u7801\u8f93\u5165\u6cd5 \u5b89\u88c5\u5411\u5bfc";
        Width = 620;
        Height = 520;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

        var title = new Label();
        title.Text = "\u5300\u7801\u8f93\u5165\u6cd5 \u5b89\u88c5\u4e0e\u914d\u7f6e";
        title.AutoSize = true;
        title.Left = 20;
        title.Top = 16;
        title.Font = new System.Drawing.Font("Microsoft YaHei UI", 12f, System.Drawing.FontStyle.Bold);
        Controls.Add(title);

        var desc = new Label();
        desc.Text = "\u5b89\u88c5\u524d\u53ef\u8bbe\u7f6e\u9ed8\u8ba4\u53c2\u6570\uff0c\u5b89\u88c5\u540e\u53ef\u7ee7\u7eed\u8c03\u6574\u3002";
        desc.AutoSize = true;
        desc.Left = 20;
        desc.Top = 48;
        Controls.Add(desc);

        var grp = new GroupBox();
        grp.Text = "\u9ed8\u8ba4\u914d\u7f6e";
        grp.Left = 20;
        grp.Top = 78;
        grp.Width = 560;
        grp.Height = 170;
        Controls.Add(grp);

        chkChineseMode = new CheckBox();
        chkChineseMode.Text = "\u9ed8\u8ba4\u4e2d\u6587\u6a21\u5f0f";
        chkChineseMode.Checked = true;
        chkChineseMode.Left = 18;
        chkChineseMode.Top = 30;
        grp.Controls.Add(chkChineseMode);

        chkFullShape = new CheckBox();
        chkFullShape.Text = "\u9ed8\u8ba4\u5168\u89d2\u6a21\u5f0f";
        chkFullShape.Checked = false;
        chkFullShape.Left = 220;
        chkFullShape.Top = 30;
        grp.Controls.Add(chkFullShape);

        var lblCount = new Label();
        lblCount.Text = "\u6bcf\u9875\u5019\u9009\u6570:";
        lblCount.Left = 18;
        lblCount.Top = 68;
        lblCount.AutoSize = true;
        grp.Controls.Add(lblCount);

        numCandidateCount = new NumericUpDown();
        numCandidateCount.Left = 150;
        numCandidateCount.Top = 64;
        numCandidateCount.Width = 80;
        numCandidateCount.Minimum = 1;
        numCandidateCount.Maximum = 6;
        numCandidateCount.Value = 6;
        grp.Controls.Add(numCandidateCount);

        var lblHotkey = new Label();
        lblHotkey.Text = "\u4e2d\u82f1\u5207\u6362\u70ed\u952e:";
        lblHotkey.Left = 18;
        lblHotkey.Top = 102;
        lblHotkey.AutoSize = true;
        grp.Controls.Add(lblHotkey);

        cmbToggleHotkey = new ComboBox();
        cmbToggleHotkey.Left = 150;
        cmbToggleHotkey.Top = 98;
        cmbToggleHotkey.Width = 140;
        cmbToggleHotkey.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbToggleHotkey.Items.Add("F9");
        cmbToggleHotkey.Items.Add("F8");
        cmbToggleHotkey.Items.Add("Ctrl+Space");
        cmbToggleHotkey.SelectedIndex = 0;
        grp.Controls.Add(cmbToggleHotkey);

        txtLog = new TextBox();
        txtLog.Left = 20;
        txtLog.Top = 262;
        txtLog.Width = 560;
        txtLog.Height = 170;
        txtLog.Multiline = true;
        txtLog.ScrollBars = ScrollBars.Vertical;
        txtLog.ReadOnly = true;
        Controls.Add(txtLog);

        btnInstall = new Button();
        btnInstall.Text = "\u5f00\u59cb\u5b89\u88c5";
        btnInstall.Left = 380;
        btnInstall.Top = 445;
        btnInstall.Width = 95;
        btnInstall.Click += (s, e) => InstallNow();
        Controls.Add(btnInstall);

        btnCancel = new Button();
        btnCancel.Text = "\u53d6\u6d88";
        btnCancel.Left = 485;
        btnCancel.Top = 445;
        btnCancel.Width = 95;
        btnCancel.Click += (s, e) => Close();
        Controls.Add(btnCancel);
    }

    private void Log(string text)
    {
        txtLog.AppendText(text + Environment.NewLine);
    }

    private void InstallNow()
    {
        btnInstall.Enabled = false;
        btnCancel.Enabled = false;
        try
        {
            Log("\u6b63\u5728\u51c6\u5907\u5b89\u88c5\u6587\u4ef6...");
            string tempDir = Path.Combine(Path.GetTempPath(), "yuninput_installer_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            byte[] zipBytes = Convert.FromBase64String(Payload.Base64);
            string zipPath = Path.Combine(tempDir, "payload.zip");
            File.WriteAllBytes(zipPath, zipBytes);

            string extractDir = Path.Combine(tempDir, "payload");
            ZipFile.ExtractToDirectory(zipPath, extractDir);

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string userDir = Path.Combine(localAppData, "yuninput");
            Directory.CreateDirectory(userDir);
            string configPath = Path.Combine(userDir, "settings.json");

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"chinese_mode\": " + (chkChineseMode.Checked ? "true" : "false") + ",");
            sb.AppendLine("  \"full_shape\": " + (chkFullShape.Checked ? "true" : "false") + ",");
            sb.AppendLine("  \"chinese_punctuation\": true,");
            sb.AppendLine("  \"smart_symbol_pairs\": true,");
            sb.AppendLine("  \"auto_commit_unique_exact\": true,");
            sb.AppendLine("  \"auto_commit_min_code_length\": 4,");
            sb.AppendLine("  \"empty_candidate_beep\": true,");
            sb.AppendLine("  \"tab_navigation\": true,");
            sb.AppendLine("  \"enter_exact_priority\": true,");
            sb.AppendLine("  \"context_association_enabled\": true,");
            sb.AppendLine("  \"context_association_max_entries\": 6000,");
            sb.AppendLine("  \"candidate_page_size\": " + numCandidateCount.Value.ToString() + ",");
            sb.AppendLine("  \"dictionary_profile\": \"zhengma-large\",");
            sb.AppendLine("  \"toggle_hotkey\": \"" + cmbToggleHotkey.SelectedItem.ToString() + "\"");
            sb.AppendLine("}");
            File.WriteAllText(configPath, sb.ToString(), Encoding.UTF8);
            Log("\u914d\u7f6e\u5df2\u4fdd\u5b58: " + configPath);

            string installPs1 = Path.Combine(extractDir, "scripts", "install_enable.ps1");
            if (!File.Exists(installPs1))
            {
                throw new Exception("\u7f3a\u5c11\u5b89\u88c5\u811a\u672c install_enable.ps1");
            }

            string logPath = Path.Combine(extractDir, "install.log");
            string args = "-NoProfile -ExecutionPolicy Bypass -Command \"& '" + installPs1 + "' *>&1 | Tee-Object -FilePath '" + logPath + "'; exit $LASTEXITCODE\"";
            Log("\u6b63\u5728\u8bf7\u6c42\u7ba1\u7406\u5458\u6743\u9650...");
            var psi = new ProcessStartInfo("powershell.exe", args);
            psi.WorkingDirectory = extractDir;
            psi.UseShellExecute = true;
            psi.Verb = "runas";
            var p = Process.Start(psi);
            p.WaitForExit();

            if (File.Exists(logPath))
            {
                Log("----- \u5b89\u88c5\u65e5\u5fd7 -----");
                Log(File.ReadAllText(logPath, Encoding.UTF8));
                Log("--------------------");
            }

            if (p.ExitCode != 0)
            {
                throw new Exception("\u5b89\u88c5\u5931\u8d25\uff0c\u9519\u8bef\u7801: " + p.ExitCode.ToString());
            }

            installCompleted = true;
            btnCancel.Text = "\u5173\u95ed";
            Log("\u5b89\u88c5\u5b8c\u6210\u3002");
            MessageBox.Show("\u5300\u7801\u8f93\u5165\u6cd5\u5b89\u88c5\u5b8c\u6210\u3002", "\u6210\u529f", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            Log("\u5b89\u88c5\u5931\u8d25: " + ex.Message);
            MessageBox.Show("\u5b89\u88c5\u5931\u8d25: " + ex.Message, "\u9519\u8bef", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnInstall.Enabled = true;
            btnCancel.Enabled = true;
            if (installCompleted)
            {
                btnInstall.Enabled = false;
            }
        }
    }
}

static class Payload
{
    public static readonly string Base64 = @"__PAYLOAD__";
}

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new InstallerForm());
    }
}
"@

$template = $template.Replace("__PAYLOAD__", $b64)
Set-Content -Path $sourceCs -Value $template -Encoding utf8

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    throw "csc.exe not found"
}

if (Test-Path $iconPath) {
    & $csc /nologo /target:winexe /platform:x64 /win32icon:$iconPath /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.IO.Compression.FileSystem.dll /out:$outExe $sourceCs
} else {
    & $csc /nologo /target:winexe /platform:x64 /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.IO.Compression.FileSystem.dll /out:$outExe $sourceCs
}
if ($LASTEXITCODE -ne 0) {
    throw "Failed to compile GUI installer"
}

Get-Item $outExe | Select-Object FullName,Length,LastWriteTime | Format-List
