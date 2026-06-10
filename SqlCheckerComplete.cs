using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Windows.Forms;
using Microsoft.Win32;

/// <summary>
/// SQL Server競合チェック ＋ SALONPOS LinQ2 稼働条件チェック 統合ツール
/// ExitCode:
///   0 = OK   (SQL未検出 かつ スペック適合)
///   1 = NG   (SQL検出あり、またはスペック非適合)
///   2 = UNKNOWN (低信頼度・判定不能)
/// </summary>
class SqlChecker
{
    // ── SQL Server スコア閾値 ──────────────────────────────
    const int THRESHOLD_INSTALLED = 5;
    const int THRESHOLD_WARN      = 2;

    // ── LinQ2 ハードウェア最低基準 ────────────────────────
    const double MIN_RAM_GB   = 8.0;
    const double MIN_DISK_GB  = 16.0;
    const int    MIN_WIDTH    = 1280;   // 1280×800 または 1366×768 のうち横幅最小値

    [STAThread]
    static void Main()
    {
        var result = Detect();
        ShowResult(result);
        WriteLog(result);

        // ExitCode決定:
        //   IsInstalled → SQL競合あり → NG=1
        //   !IsSpecOk   → スペック非適合 → NG=1
        //   LowConfidenceNoDetect → UNKNOWN=2
        //   それ以外 → OK=0
        int exitCode;
        if (result.IsInstalled || !result.IsSpecOk)
            exitCode = 1;
        else if (result.LowConfidenceNoDetect)
            exitCode = 2;
        else
            exitCode = 0;

        Environment.Exit(exitCode);
    }

    // ═══════════════════════════════════════════════════════
    //   DetectionResult: SQL判定 ＋ スペック判定を一元管理
    // ═══════════════════════════════════════════════════════
    class DetectionResult
    {
        // ── SQL Server 検出スコア関連 ──
        public int  Score          = 0;
        public bool HasFullSqlServer = false;
        public bool HasLocalDb       = false;
        public List<string> Details = new List<string>();
        public List<string> Skipped = new List<string>();

        // ── LinQ2 スペック判定関連 ──
        /// <summary>false になった時点で最終判定はNG</summary>
        public bool IsSpecOk = true;
        /// <summary>WARNレベル（動作可能だが推奨外）</summary>
        public bool IsSpecWarn = false;
        public List<string> SpecDetails = new List<string>();

        // ── 集約プロパティ ──
        public bool IsInstalled =>
            HasFullSqlServer || Score >= THRESHOLD_INSTALLED;

        public bool IsWarn =>
            !IsInstalled && Score >= THRESHOLD_WARN;

        public bool LowConfidenceNoDetect =>
            Score == 0 && !HasLocalDb &&
            Skipped.Any(s => s.Contains("サービス")) &&
            Skipped.Any(s => s.Contains("レジストリ"));

        public string SqlConfidence =>
            Score >= 8 ? "高" : Score >= 5 ? "中" : Score >= 2 ? "低" : "なし";

        /// <summary>
        /// 総合判定テキスト
        ///   NG:      SQL競合あり、またはスペック非適合
        ///   WARN:    SQLはグレー、またはスペックWARN（Win10等）
        ///   UNKNOWN: 低信頼度で判定不能
        ///   OK:      問題なし
        /// </summary>
        public string JudgeText
        {
            get
            {
                if (IsInstalled || !IsSpecOk) return "NG";
                if (LowConfidenceNoDetect)    return "UNKNOWN";
                if (IsWarn || IsSpecWarn)     return "WARN";
                return "OK";
            }
        }
    }

    // ═══════════════════════════════════════════════════════
    //   メイン検出処理
    // ═══════════════════════════════════════════════════════
    static DetectionResult Detect()
    {
        var r = new DetectionResult();

        // SQL Server 既存チェック群
        CheckServices(r);
        CheckRegistry(r);
        CheckLocalDb(r);
        CheckDirectories(r);
        CheckPathEnv(r);

        // LinQ2 スペックチェック（本改修の追加部分）
        CheckLinQ2HardwareSpec(r);

        return r;
    }

    // ═══════════════════════════════════════════════════════
    //   LinQ2 スペック判定ロジック（改修追加）
    // ═══════════════════════════════════════════════════════
    static void CheckLinQ2HardwareSpec(DetectionResult r)
    {
        r.SpecDetails.Add("════════════════════════════════════");
        r.SpecDetails.Add("【LinQ2 稼働条件判定結果】");
        r.SpecDetails.Add("════════════════════════════════════");

        CheckDeviceModel(r);
        CheckCpu(r);
        CheckRam(r);
        CheckDisk(r);
        CheckResolution(r);
        CheckOsEdition(r);

        // 物理ポートは自動判定不可のため、常に注意書きを付与
        r.SpecDetails.Add("");
        r.SpecDetails.Add("⚠️ [要目視確認] USB Type-Aポートおよび有線LANポートが");
        r.SpecDetails.Add("   物理的に存在するか必ずご確認ください。");
        r.SpecDetails.Add("   （レシートプリンタ・CTI機器の接続に必須です）");
    }

    /// <summary>1. デバイスメーカー・モデル確認（Surface / Mac の排除）</summary>
    static void CheckDeviceModel(DetectionResult r)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Manufacturer, Model FROM Win32_ComputerSystem");
            foreach (ManagementObject mo in searcher.Get())
            {
                string mfr   = mo["Manufacturer"]?.ToString() ?? "";
                string model = mo["Model"]?.ToString()        ?? "";

                bool isSurface = mfr.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) >= 0
                              && model.IndexOf("Surface",  StringComparison.OrdinalIgnoreCase) >= 0;
                bool isMac     = mfr.IndexOf("Apple", StringComparison.OrdinalIgnoreCase) >= 0;

                if (isSurface)
                {
                    r.SpecDetails.Add(
                        $"❌ 端末モデル: {mfr} {model}" + Environment.NewLine +
                        "   → Microsoft Surfaceシリーズは非対応OSが搭載されている場合があり、" + Environment.NewLine +
                        "     解像度も非推奨のため、LinQ2 導入不可パソコンです。");
                    r.IsSpecOk = false;
                }
                else if (isMac)
                {
                    r.SpecDetails.Add(
                        $"❌ 端末モデル: {mfr} {model}" + Environment.NewLine +
                        "   → Macパソコンでは動作しません。" + Environment.NewLine +
                        "     Windows環境を構築しても動作しません。");
                    r.IsSpecOk = false;
                }
                else
                {
                    r.SpecDetails.Add($"✅ 端末モデル: {mfr} {model}");
                }
            }
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"デバイスモデル確認失敗（WMI）: {ex.Message}");
            r.SpecDetails.Add("⚠️ 端末モデル: 確認失敗（WMI不可）");
        }
    }

    /// <summary>
    /// 2. CPU確認
    ///    ・Intel製Coreシリーズ / AMD製Ryzenのみ対応
    ///    ・ARMアーキテクチャは非対応（Surface Pro X 等）
    ///    ・Atom / Celeron / Pentium 等は非対応
    /// </summary>
    static void CheckCpu(DetectionResult r)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Name, Architecture FROM Win32_Processor");
            foreach (ManagementObject mo in searcher.Get())
            {
                string cpuName = mo["Name"]?.ToString() ?? "(不明)";

                // Architecture: 0=x86, 9=x64, 12=ARM64
                uint arch = (uint)(mo["Architecture"] ?? 9u);
                bool isArm = (arch == 12);

                bool isCore  = cpuName.IndexOf("Core",  StringComparison.OrdinalIgnoreCase) >= 0;
                bool isRyzen = cpuName.IndexOf("Ryzen", StringComparison.OrdinalIgnoreCase) >= 0;

                if (isArm)
                {
                    r.SpecDetails.Add(
                        $"❌ CPU: {cpuName}" + Environment.NewLine +
                        "   → ARMアーキテクチャのCPUは非対応です（ARM版Windows11も非対応）。");
                    r.IsSpecOk = false;
                }
                else if (!isCore && !isRyzen)
                {
                    r.SpecDetails.Add(
                        $"❌ CPU: {cpuName}" + Environment.NewLine +
                        "   → Intel製CoreシリーズまたはAMD製Ryzenのみ対応です。" + Environment.NewLine +
                        "     Celeron / Pentium / Atom 等は非対応です。");
                    r.IsSpecOk = false;
                }
                else
                {
                    r.SpecDetails.Add($"✅ CPU: {cpuName}（条件適合）");
                }
            }
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"CPU確認失敗（WMI）: {ex.Message}");
            r.SpecDetails.Add("⚠️ CPU: 確認失敗（WMI不可）");
        }
    }

    /// <summary>3. メモリ容量確認（8GB以上）</summary>
    static void CheckRam(DetectionResult r)
    {
        try
        {
            var compInfo   = new Microsoft.VisualBasic.Devices.ComputerInfo();
            double totalGb = Math.Round(
                (double)compInfo.TotalPhysicalMemory / (1024.0 * 1024 * 1024), 1);
            bool ok = totalGb >= MIN_RAM_GB;

            r.SpecDetails.Add(
                $"{(ok ? "✅" : "❌")} メモリ: {totalGb} GB" +
                $"（基準: {MIN_RAM_GB} GB以上）");
            if (!ok) r.IsSpecOk = false;
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"メモリ確認失敗: {ex.Message}");
            r.SpecDetails.Add("⚠️ メモリ: 確認失敗");
        }
    }

    /// <summary>4. Cドライブ空き容量確認（16GB以上・推奨SSD）</summary>
    static void CheckDisk(DetectionResult r)
    {
        try
        {
            var   drive    = new DriveInfo("C");
            double freeGb  = Math.Round(
                (double)drive.AvailableFreeSpace / (1024.0 * 1024 * 1024), 1);
            bool ok = freeGb >= MIN_DISK_GB;

            r.SpecDetails.Add(
                $"{(ok ? "✅" : "❌")} Cドライブ空き容量: {freeGb} GB" +
                $"（基準: {MIN_DISK_GB} GB以上、推奨SSD）");
            if (!ok) r.IsSpecOk = false;
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"ディスク確認失敗: {ex.Message}");
            r.SpecDetails.Add("⚠️ ディスク: 確認失敗");
        }
    }

    /// <summary>
    /// 5. 解像度確認
    ///    NG:   横幅 < 1280
    ///    WARN: 1280≤横幅 かつ 推奨値（1366×768 / 1280×800）未満
    ///          → キャッシャー画面の来店履歴一覧が表示されない可能性あり
    /// </summary>
    static void CheckResolution(DetectionResult r)
    {
        int w = Screen.PrimaryScreen.Bounds.Width;
        int h = Screen.PrimaryScreen.Bounds.Height;

        // NG: 横幅が絶対最小値未満
        if (w < MIN_WIDTH)
        {
            r.SpecDetails.Add(
                $"❌ 画面解像度: {w}×{h}" + Environment.NewLine +
                $"   → 横幅 {MIN_WIDTH}px 未満のため非対応です。");
            r.IsSpecOk = false;
            return;
        }

        // 推奨基準チェック（16:9 → 1366×768以上、16:10 → 1280×800以上）
        bool meets1366x768 = (w >= 1366 && h >= 768);
        bool meets1280x800 = (w >= 1280 && h >= 800);
        bool meetsRecommended = meets1366x768 || meets1280x800;

        if (meetsRecommended)
        {
            r.SpecDetails.Add($"✅ 画面解像度: {w}×{h}（推奨基準を満たしています）");
        }
        else
        {
            // 横幅は1280以上あるが推奨基準を下回る → WARN
            r.SpecDetails.Add(
                $"⚠️ 画面解像度: {w}×{h}" + Environment.NewLine +
                "   → 推奨基準（1366×768 または 1280×800）を下回っています。" + Environment.NewLine +
                "     キャッシャー画面で来店履歴一覧が表示されない場合があります。");
            r.IsSpecWarn = true;
        }
    }

    /// <summary>
    /// 6. OSエディション確認
    ///    NG:   N / S / Education / Workstations / ARM版 / Server / Mac
    ///    WARN: Windows 10 Home / Pro（サポート終了済みだが動作可）
    ///    OK:   Windows 11 Home / Pro（日本語版）
    /// </summary>
    static void CheckOsEdition(DetectionResult r)
    {
        try
        {
            // ── OS名取得: VisualBasic.ComputerInfo.OSFullName は内部バージョン番号
            //    ("Microsoft Windows 10.0.26100" 等) を返す場合があるため、
            //    WMI の Caption（"Microsoft Windows 11 Pro" 等の正式表示名）を優先使用。
            //    BuildNumber で Win10/Win11 を確定判定し、Caption の誤表記を補正する。
            string osName    = "";
            int    buildNum  = 0;

            try
            {
                using var osSearcher = new ManagementObjectSearcher(
                    "SELECT Caption, BuildNumber FROM Win32_OperatingSystem");
                foreach (ManagementObject mo in osSearcher.Get())
                {
                    osName   = mo["Caption"]?.ToString()     ?? "";
                    int.TryParse(mo["BuildNumber"]?.ToString(), out buildNum);
                    break;
                }
            }
            catch
            {
                // WMI失敗時はフォールバックとしてEnvironment.OSVersionを使用
                osName = $"Windows (Build {Environment.OSVersion.Version.Build})";
                buildNum = Environment.OSVersion.Version.Build;
            }

            // BuildNumber による Win10 / Win11 確定判定
            // Win11 は Build 22000以上（21H2〜）、Win10 は 10240〜19045
            bool isActuallyWin11 = buildNum >= 22000;
            bool isActuallyWin10 = buildNum >= 10240 && buildNum < 22000;

            // Caption に "Windows 10" と書かれていてもBuild番号が22000以上なら Win11 に補正
            if (isActuallyWin11 && osName.Contains("Windows 10"))
                osName = osName.Replace("Windows 10", "Windows 11");

            // ─ 非対応エディションキーワード ─
            string[] ngKeywords =
            {
                "Home N",
                "Home S",
                "Pro N",
                "Pro Education",
                "Education N",
                "Education",
                "Workstations",
                "Enterprise",
                "Server",
                "ARM"
            };

            // ─ ARM版の独立チェック（WMI Architecture で補完） ─
            bool isArmOs = false;
            try
            {
                using var cpuSearcher = new ManagementObjectSearcher(
                    "SELECT Architecture FROM Win32_Processor");
                foreach (ManagementObject mo in cpuSearcher.Get())
                {
                    if ((uint)(mo["Architecture"] ?? 9u) == 12)
                        isArmOs = true;
                }
            }
            catch { /* WMI失敗時はOS名文字列のみで判定 */ }

            bool isNgEdition = ngKeywords.Any(k =>
                osName.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                || isArmOs;

            // BuildNumber 補正済みの文字列で判定
            bool isWin10 = isActuallyWin10;
            bool isWin11 = isActuallyWin11;

            if (isNgEdition)
            {
                r.SpecDetails.Add(
                    $"❌ OS: {osName}" + Environment.NewLine +
                    "   → このOSエディションはLinQ2非対応です。" + Environment.NewLine +
                    "     対応OS: Windows 11 Home / Windows 11 Pro（日本語版）のみ。");
                r.IsSpecOk = false;
            }
            else if (isWin10)
            {
                // Windows 10 Home / Pro のみ動作可（サポート終了済み警告）
                r.SpecDetails.Add(
                    $"⚠️ OS: {osName}" + Environment.NewLine +
                    "   → 動作可能ですが、Microsoftのサポートが2025年10月14日に終了しています。" + Environment.NewLine +
                    "     新規導入・PC購入の場合はWindows 11を強く推奨します。");
                r.IsSpecWarn = true;
            }
            else if (isWin11)
            {
                r.SpecDetails.Add($"✅ OS: {osName}（対応OS）");
            }
            else
            {
                // Windows 7/8/XP 等、想定外のOS
                r.SpecDetails.Add(
                    $"❌ OS: {osName}" + Environment.NewLine +
                    "   → LinQ2が対応していないOSバージョンです。");
                r.IsSpecOk = false;
            }
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"OS確認失敗: {ex.Message}");
            r.SpecDetails.Add("⚠️ OS: 確認失敗");
        }
    }

    // ═══════════════════════════════════════════════════════
    //   SQL Server 既存チェック群（変更なし・骨格のみ示す）
    // ═══════════════════════════════════════════════════════

    static void CheckServices(DetectionResult r)
    {
        try
        {
            var services = ServiceController.GetServices();
            var sqlSvcs  = services.Where(s =>
                s.ServiceName.IndexOf("MSSQL", StringComparison.OrdinalIgnoreCase) >= 0 ||
                s.DisplayName.IndexOf("SQL Server", StringComparison.OrdinalIgnoreCase) >= 0
            ).ToList();

            if (sqlSvcs.Count == 0)
            {
                r.Details.Add("✅ SQLサービス: 未検出");
                return;
            }

            foreach (var svc in sqlSvcs)
            {
                bool running = svc.Status == ServiceControllerStatus.Running;
                r.Details.Add(
                    $"{(running ? "❌" : "⚠️")} SQLサービス検出: {svc.ServiceName} [{svc.Status}]");

                if (svc.ServiceName.IndexOf("SQLEXPRESS", StringComparison.OrdinalIgnoreCase) < 0
                 && svc.ServiceName.IndexOf("LocalDB",    StringComparison.OrdinalIgnoreCase) < 0)
                {
                    r.HasFullSqlServer = true;
                }

                r.Score += running ? 3 : 1;
            }
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"サービス確認失敗: {ex.Message}");
        }
    }

    static void CheckRegistry(DetectionResult r)
    {
        try
        {
            string[] keys =
            {
                @"SOFTWARE\Microsoft\Microsoft SQL Server",
                @"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server"
            };

            foreach (string key in keys)
            {
                using var rk = Registry.LocalMachine.OpenSubKey(key);
                if (rk == null) continue;

                r.Details.Add($"❌ レジストリ検出: HKLM\\{key}");
                r.Score += 2;
            }

            if (!r.Details.Any(d => d.Contains("レジストリ")))
                r.Details.Add("✅ レジストリ: 未検出");
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"レジストリ確認失敗: {ex.Message}");
        }
    }

    static void CheckLocalDb(DetectionResult r)
    {
        try
        {
            string[] localDbPaths =
            {
                @"SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Installed Versions",
                @"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server Local DB\Installed Versions"
            };

            foreach (string path in localDbPaths)
            {
                using var rk = Registry.LocalMachine.OpenSubKey(path);
                if (rk == null) continue;

                r.HasLocalDb = true;
                r.Details.Add($"⚠️ LocalDB検出: {path}");
                r.Score += 1;
            }

            if (!r.HasLocalDb)
                r.Details.Add("✅ LocalDB: 未検出");
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"LocalDB確認失敗: {ex.Message}");
        }
    }

    static void CheckDirectories(DetectionResult r)
    {
        try
        {
            string[] dirs =
            {
                @"C:\Program Files\Microsoft SQL Server",
                @"C:\Program Files (x86)\Microsoft SQL Server"
            };

            bool found = false;
            foreach (string dir in dirs)
            {
                if (!Directory.Exists(dir)) continue;
                found = true;
                r.Details.Add($"⚠️ SQLディレクトリ検出: {dir}");
                r.Score += 1;
            }

            if (!found)
                r.Details.Add("✅ SQLディレクトリ: 未検出");
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"ディレクトリ確認失敗: {ex.Message}");
        }
    }

    static void CheckPathEnv(DetectionResult r)
    {
        try
        {
            string path = Environment.GetEnvironmentVariable("PATH") ?? "";
            bool hasSqlPath = path.Split(';').Any(p =>
                p.IndexOf("SQL", StringComparison.OrdinalIgnoreCase) >= 0 &&
                p.IndexOf("Microsoft SQL Server", StringComparison.OrdinalIgnoreCase) >= 0);

            if (hasSqlPath)
            {
                r.Details.Add("⚠️ PATH環境変数にSQL Serverパスを検出");
                r.Score += 1;
            }
            else
            {
                r.Details.Add("✅ PATH環境変数: SQL Server未検出");
            }
        }
        catch (Exception ex)
        {
            r.Skipped.Add($"PATH確認失敗: {ex.Message}");
        }
    }

    // ═══════════════════════════════════════════════════════
    //   結果表示
    // ═══════════════════════════════════════════════════════
    static void ShowResult(DetectionResult r)
    {
        string border = new string('═', 52);
        var lines = new List<string>();

        lines.Add(border);
        lines.Add($"  SALONPOS LinQ2 導入前チェックツール");
        lines.Add($"  実行日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
        lines.Add(border);
        lines.Add($"  総合判定: 【 {r.JudgeText} 】");
        lines.Add(border);

        // ── SQL Server 判定セクション ──
        lines.Add("");
        lines.Add("■ SQL Server 競合チェック");
        lines.Add($"  スコア: {r.Score}  信頼度: {r.SqlConfidence}");
        if (r.HasFullSqlServer)
            lines.Add("  → フルSQL Serverを検出。LinQ2のSQL Serverと競合する可能性が高い。");
        else if (r.IsInstalled)
            lines.Add("  → SQL Server関連コンポーネントを検出。要精査。");
        else if (r.IsWarn)
            lines.Add("  → 軽微な痕跡を検出。引き続き経過観察を推奨。");
        else if (r.LowConfidenceNoDetect)
            lines.Add("  → 判定に必要な情報が取得できませんでした（権限不足の可能性）。");
        else
            lines.Add("  → SQL Server競合なし。");

        foreach (var d in r.Details)
            lines.Add("    " + d);

        if (r.Skipped.Count > 0)
        {
            lines.Add("  [スキップ項目]");
            foreach (var s in r.Skipped)
                lines.Add("    " + s);
        }

        // ── LinQ2 スペック判定セクション ──
        lines.Add("");
        lines.Add("■ LinQ2 稼働条件チェック");
        if (!r.IsSpecOk)
            lines.Add("  → ❌ 必須スペックを満たしていません。LinQ2は導入できません。");
        else if (r.IsSpecWarn)
            lines.Add("  → ⚠️ 動作可能ですが、一部推奨基準を下回っています。");
        else
            lines.Add("  → ✅ 稼働条件をすべて満たしています。");

        foreach (var d in r.SpecDetails)
            lines.Add("  " + d);

        lines.Add("");
        lines.Add(border);

        // ダイアログ表示
        string message = string.Join(Environment.NewLine, lines);
        MessageBox.Show(message, $"LinQ2 チェック結果: {r.JudgeText}",
            MessageBoxButtons.OK,
            r.JudgeText == "NG"      ? MessageBoxIcon.Error   :
            r.JudgeText == "WARN"    ? MessageBoxIcon.Warning :
            r.JudgeText == "UNKNOWN" ? MessageBoxIcon.Question :
                                       MessageBoxIcon.Information);
    }

    // ═══════════════════════════════════════════════════════
    //   ログ出力
    // ═══════════════════════════════════════════════════════
    static void WriteLog(DetectionResult r)
    {
        try
        {
            string logPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                $"linq2_check_{DateTime.Now:yyyyMMdd_HHmmss}.log");

            using var sw = new StreamWriter(logPath, false,
                System.Text.Encoding.UTF8);

            sw.WriteLine($"SALONPOS LinQ2 導入前チェックログ");
            sw.WriteLine($"実行日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
            sw.WriteLine($"総合判定: {r.JudgeText}");
            sw.WriteLine();

            sw.WriteLine("─── SQL Server チェック ───");
            sw.WriteLine($"スコア={r.Score}, 信頼度={r.SqlConfidence}");
            sw.WriteLine($"HasFullSqlServer={r.HasFullSqlServer}, HasLocalDb={r.HasLocalDb}");
            foreach (var d in r.Details)  sw.WriteLine(d);
            foreach (var s in r.Skipped) sw.WriteLine("[SKIP] " + s);

            sw.WriteLine();
            sw.WriteLine("─── LinQ2 スペックチェック ───");
            sw.WriteLine($"IsSpecOk={r.IsSpecOk}, IsSpecWarn={r.IsSpecWarn}");
            foreach (var d in r.SpecDetails) sw.WriteLine(d);

            sw.WriteLine();
            sw.WriteLine("※ USB Type-Aポート・有線LANポートの有無は目視確認が必要です。");
        }
        catch
        {
            // ログ書き込み失敗は無視（チェック結果自体に影響させない）
        }
    }
}
