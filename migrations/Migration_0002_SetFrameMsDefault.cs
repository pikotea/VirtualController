using System;
using System.IO;
using System.Text;

public class Migration_0002_SetFrameMsDefault : IMigration
{
    public string Id => "0002_set_frame_ms_default";
    public string Description => "既存の settings.ini があれば FrameMs の値を 16.666 に更新するマイグレーション。";

    public void Up()
    {
        var app = AppDomain.CurrentDomain.BaseDirectory;
        var settingsPath = Path.Combine(app, "settings.ini");
        if (!File.Exists(settingsPath)) return; // 設定ファイルがなければ何もしない

        try
        {
            var lines = File.ReadAllLines(settingsPath, Encoding.UTF8);
            bool changed = false;
            for (int i = 0; i < lines.Length; i++)
            {
                var trimmed = lines[i].TrimStart();
                if (trimmed.StartsWith("FrameMs=", StringComparison.OrdinalIgnoreCase))
                {
                    // 保存済みの FrameMs を新しい値に置換（1件目のみ）
                    lines[i] = "FrameMs=16.666";
                    changed = true;
                    break;
                }
            }
            if (changed)
            {
                File.WriteAllLines(settingsPath, lines, Encoding.UTF8);
            }
        }
        catch (Exception)
        {
            // マイグレーションは安全に無視できる失敗にする
        }
    }
}