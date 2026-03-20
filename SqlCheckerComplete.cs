using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Windows.Forms;
using Microsoft.Win32;

class SqlChecker
{
    const int THRESHOLD_INSTALLED = 5;
    const int THRESHOLD_WARN      = 2;

    [STAThread]
    static void Main()
    {
        var result = Detect();
        ShowResult(result);
        WriteLog(result);

        // 終了コード: 0=インストール可, 1=インストール不可, 2=判定不能
        // インストーラ統合・バッチ自動化での利用を想定
        int exitCode = result.IsInstalled ? 1 : (result.LowConfidenceNoDetect ? 2 : 0);
        Environment.Exit(exitCode);
    }

    class DetectionResult
    {
        public int Score = 0;
        public bool HasFullSqlServer = false;
        public bool HasLocalDb       = false;

        public List<string> Details = new List<string>();
        public List<string> Skipped = new List<string>();

        public bool IsInstalled => HasFullSqlServer || Score >= THRESHOLD_INSTALLED;
        public bool IsWarn      => !IsInstalled && Score >= THRESHOLD_WARN;

        // サービスとレジストリの両系統がスキップされた場合のみ「判定不能」
        // LocalDB1件スキップ程度では UNKNOWN にしない
        public bool LowConfidenceNoDetect =>
            Score == 0 && !HasLocalDb &&
            Skipped.Any(s => s.Contains("サービス")) &&
            Skipped.Any(s => s.Contains("レジストリ"));

        public string Confidence =>
            Score >= 8 ? "高" :
            Score >= 5 ? "中" :
            Score >= 2 ? "低（残骸の可能性）" : "なし";

        public string JudgeText =>
            IsInstalled          ? "NG"      :
            LowConfidenceNoDetect ? "UNKNOWN" :
            IsWarn               ? "WARN"    : "OK";
    }

    static DetectionResult Detect()
    {
        var r = new DetectionResult();
        CheckServices(r);
        CheckRegistry(r);
        CheckLocalDb(r);
        CheckDirectories(r);
        CheckPathEnv(r);
        return r;
    }

    // ── 戦略1: Windowsサービス（最強証拠 +5）────────────────────────────
    static void CheckServices(DetectionResult r)
    {
        bool mainCounted    = false;
        bool relatedCounted = false;

        try
        {
            foreach (var svc in ServiceController.GetServices())
            {
                string name = svc.ServiceName;

                if (name.Equals("MSSQLSERVER", StringComparison.OrdinalIgnoreCase) ||
                    name.StartsWith("MSSQL$", StringComparison.OrdinalIgnoreCase))
                {
                    r.HasFullSqlServer = true;
                    r.Details.Add($"[強:Service] {name} (Status={svc.Status})");
                    if (!mainCounted) { r.Score += 5; mainCounted = true; }
                }
                // Fix1: SQLSERVERAGENT も明示的に含める（Gemini版は StartsWith("SQLAGENT") のみでエージェントサービス名の変形を見落とす可能性）
                else if (name.Equals("SQLBrowser",     StringComparison.OrdinalIgnoreCase) ||
                         name.Equals("SQLSERVERAGENT",  StringComparison.OrdinalIgnoreCase) ||
                         name.StartsWith("SQLAGENT$",   StringComparison.OrdinalIgnoreCase))
                {
                    r.Details.Add($"[中:Service] 関連サービス: {name} (Status={svc.Status})");
                    if (!relatedCounted) { r.Score += 2; relatedCounted = true; }
                }
            }
        }
        catch (Exception ex) { r.Skipped.Add($"サービス取得スキップ: {ex.Message}"); }
    }

    // ── 戦略2: レジストリ（+3 中 / サブキー残存 +1 弱）──────────────────
    static void CheckRegistry(DetectionResult r)
    {
        TryRegistryCheck(r, Registry.LocalMachine, @"SOFTWARE\Microsoft\Microsoft SQL Server",             "HKLM");
        TryRegistryCheck(r, Registry.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server", "WOW6432");
        TryRegistryCheck(r, Registry.CurrentUser,  @"SOFTWARE\Microsoft\Microsoft SQL Server",             "HKCU");
    }

    static void TryRegistryCheck(DetectionResult r, RegistryKey hive, string path, string label)
    {
        try
        {
            using var key = hive.OpenSubKey(path);
            if (key == null) return;

            var instances = key.GetValue("InstalledInstances") as string[];
            if (instances?.Length > 0)
            {
                r.HasFullSqlServer = true;
                r.Score += 3;
                r.Details.Add($"[中:Registry-{label}] InstalledInstances: {string.Join(", ", instances)}");
                return;
            }

            // サブキー残存はまとめて +1（Gemini版の keyCounted を採用: 複数残骸での重複スコア防止）
            bool keyCounted = false;
            foreach (var sk in key.GetSubKeyNames())
            {
                if (sk.StartsWith("MSSQL", StringComparison.OrdinalIgnoreCase))
                {
                    if (!keyCounted) { r.Score += 1; keyCounted = true; }
                    r.Details.Add($"[弱:Registry-{label}] サブキー残存: {sk}");
                }
            }
        }
        catch (Exception ex) { r.Skipped.Add($"レジストリ({label})スキップ: {ex.Message}"); }
    }

    // ── 戦略3: LocalDB（+2、IsWarnに影響するがIsInstalledには影響しない）──
    static void CheckLocalDb(DetectionResult r)
    {
        string[] paths = {
            @"SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Installed Versions",
            @"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server Local DB\Installed Versions"
        };

        foreach (var p in paths)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(p);
                var versions = key?.GetSubKeyNames();
                if (versions?.Length > 0)
                {
                    r.HasLocalDb = true;
                    r.Score += 2;
                    r.Details.Add($"[LocalDB] 検出: {string.Join(", ", versions)}");
                    return; // 二重カウント防止
                }
            }
            catch (Exception ex) { r.Skipped.Add($"LocalDBチェックスキップ: {ex.Message}"); }
        }
    }

    // ── 戦略4: インストールディレクトリ（sqlservr.exe +3 中 / フォルダのみ +1 弱）
    static void CheckDirectories(DetectionResult r)
    {
        string[] bases = {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),    "Microsoft SQL Server"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft SQL Server")
        };

        // Gemini版の countedDir を採用: 64bit/32bit 両パスでの重複スコア防止
        bool countedDir = false;

        foreach (var root in bases)
        {
            if (!Directory.Exists(root)) continue;

            try
            {
                foreach (var dir in Directory.GetDirectories(root, "MSSQL*"))
                {
                    string binnDir = Path.Combine(dir, "MSSQL", "Binn");
                    if (Directory.Exists(binnDir) && Directory.GetFiles(binnDir, "sqlservr.exe").Any())
                    {
                        r.HasFullSqlServer = true;
                        r.Details.Add($"[中:Directory] sqlservr.exe 確認: {dir}");
                        if (!countedDir) { r.Score += 3; countedDir = true; }
                    }
                    else
                    {
                        r.Details.Add($"[弱:Directory] フォルダのみ（実体未確認）: {dir}");
                        if (!countedDir) { r.Score += 1; countedDir = true; }
                    }
                }
            }
            catch (Exception ex) { r.Skipped.Add($"ディレクトリチェックスキップ: {ex.Message}"); }
        }
    }

    // ── 戦略5: 環境変数 PATH（+1 弱）──────────────────────────────────
    // Fix2: "sqlcmd.exe" を優先しつつ "SQL Server" も保持（両方あれば1点のみ）
    static void CheckPathEnv(DetectionResult r)
    {
        string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";

        if (pathEnv.IndexOf("sqlcmd.exe", StringComparison.OrdinalIgnoreCase) >= 0 ||
            pathEnv.IndexOf("SQL Server",  StringComparison.OrdinalIgnoreCase) >= 0)
        {
            r.Score += 1;
            r.Details.Add("[弱:EnvPath] SQL関連パスを検出");
        }
    }

    // ── UI表示（判定優先順位: NG > UNKNOWN > WARN > OK）──────────────────
    static void ShowResult(DetectionResult r)
    {
        string title    = "LinQ インストール可否チェック";
        string scoreInfo = $"確度: {r.Confidence} (Score: {r.Score})\n";

        if (r.IsInstalled)
        {
            MessageBox.Show(
                "⛔ インストール不可\n" +
                "SQL Server が既にインストールされているため、\n" +
                "LinQ と競合しインストールできません。\n\n" +
                scoreInfo +
                "\n【根拠】\n" + string.Join("\n", r.Details),
                title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        else if (r.LowConfidenceNoDetect)
        {
            MessageBox.Show(
                "⚠️ 判定不能\n" +
                "権限不足により主要な確認項目（サービス・レジストリ）を\n" +
                "スキャンできませんでした。\n" +
                "管理者権限での再実行を推奨します。\n\n" +
                "【スキップ内容】\n" + string.Join("\n", r.Skipped),
                title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else if (r.IsWarn)
        {
            MessageBox.Show(
                "⚠️ 警告：SQL Server の残骸を検出\n" +
                "構成の一部が残っています。実体がないため\n" +
                "LinQ のインストールは可能と判断しますが、\n" +
                "念のためシステム管理者への確認を推奨します。\n\n" +
                scoreInfo +
                "\n【検出内容】\n" + string.Join("\n", r.Details),
                title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else
        {
            string note = r.HasLocalDb
                ? "\n（LocalDB は検出されましたが競合なしと判断）"
                : "";

            MessageBox.Show(
                "✅ インストール可能\n" +
                "Microsoft SQL Server は検出されませんでした。" + note,
                title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // ── ログ出力（%TEMP%\SqlCheck_yyyyMMdd_HHmmss.log）──────────────────
    static void WriteLog(DetectionResult r)
    {
        try
        {
            string logPath = Path.Combine(
                Path.GetTempPath(),
                $"SqlCheck_{DateTime.Now:yyyyMMdd_HHmmss}.log");

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"[判定結果: {r.JudgeText}]");
            sb.AppendLine($"ExitCode: {(r.IsInstalled ? 1 : (r.LowConfidenceNoDetect ? 2 : 0))}");
            sb.AppendLine($"日時: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | PC: {Environment.MachineName} | User: {Environment.UserName}");
            sb.AppendLine($"Score={r.Score}, FullSQL={r.HasFullSqlServer}, LocalDB={r.HasLocalDb}");
            sb.AppendLine();
            sb.AppendLine("--- 検出ソース ---");
            r.Details.ForEach(d => sb.AppendLine(d));
            sb.AppendLine();
            sb.AppendLine("--- スキップ ---");
            r.Skipped.ForEach(s => sb.AppendLine(s));

            File.WriteAllText(logPath, sb.ToString(), System.Text.Encoding.UTF8);
        }
        catch { }
    }
}
