using System;
using System.IO;

public class Migration_0001_MicrosToMacros : IMigration
{
    public string Id => "0001_micros_to_macros";
    public string Description => "Move legacy 'micros' folder to new 'macros' folder (CSV files).";

    public void Up()
    {
        var app = AppDomain.CurrentDomain.BaseDirectory;
        var oldDir = Path.Combine(app, "micros");
        var newDir = Path.Combine(app, "macros");

        if (!Directory.Exists(oldDir)) return; // 既に移行済みなら何もしない

        // 新フォルダが無ければ単純移動
        if (!Directory.Exists(newDir))
        {
            Directory.Move(oldDir, newDir);
            return;
        }

        // oldDir 内の CSV を列挙
        var files = Directory.GetFiles(oldDir, "*.csv");
        // フォルダが空ならそのまま削除して終了
        if (files.Length == 0)
        {
            try
            {
                Directory.Delete(oldDir, true);
            }
            catch (Exception)
            {
                // 削除失敗時は無視
            }
            return;
        }

        string backupDir = null;

        foreach (var src in files)
        {
            var name = Path.GetFileName(src);
            var dst = Path.Combine(newDir, name);
            if (!File.Exists(dst))
            {
                File.Copy(src, dst);
            }
            else
            {
                // 衝突が初めて発生した時点でバックアップフォルダを作成
                if (backupDir == null)
                {
                    backupDir = Path.Combine(app, $"micros_migration_backup_{DateTime.Now:yyyyMMdd_HHmmss}");
                    Directory.CreateDirectory(backupDir);
                }
                var b = Path.Combine(backupDir, name);
                File.Copy(src, b);
            }
        }

        // 旧フォルダを削除（内容をすべて削除してフォルダ自体を削除）
        try
        {
            if (Directory.Exists(oldDir))
            {
                Directory.Delete(oldDir, true); // true = 再帰的に削除
            }
        }
        catch (Exception)
        {
            // 削除失敗時は無視
        }
    }
}