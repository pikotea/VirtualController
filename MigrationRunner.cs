using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public static class MigrationRunner
{
    // 履歴ファイル名（Application.StartupPath に作成）
    private const string HistoryFileName = "migrations_applied.txt";

    public static void RunMigrations(IEnumerable<IMigration> migrations)
    {
        if (migrations == null) return;

        var historyPath = Path.Combine(Application.StartupPath, HistoryFileName);
        var applied = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (File.Exists(historyPath))
        {
            foreach (var line in File.ReadAllLines(historyPath))
            {
                var t = line.Trim();
                if (!string.IsNullOrEmpty(t)) applied.Add(t);
            }
        }

        var pending = migrations.OrderBy(m => m.Id).Where(m => !applied.Contains(m.Id)).ToList();
        if (pending.Count == 0) return;

        foreach (var mig in pending)
        {
            try
            {
                mig.Up();
                File.AppendAllText(historyPath, mig.Id + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // 移行失敗時はユーザーに通知して処理を中断
                MessageBox.Show($"マイグレーション '{mig.Id}' に失敗しました: {mig.Description}\n\n{ex.Message}",
                    "マイグレーションエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // 失敗したら以降のマイグレーションは行わない
                break;
            }
        }
    }
}