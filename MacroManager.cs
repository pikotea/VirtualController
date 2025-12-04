using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace VirtualController
{
    public class MacroManager
    {
        private readonly string macroFolder;

        public MacroManager(string macroFolder)
        {
            this.macroFolder = macroFolder;
            if (!Directory.Exists(macroFolder))
            {
                Directory.CreateDirectory(macroFolder);
            }
        }

        // 新: マクロ一覧エントリ（ファイル or フォルダ）
        public class MacroEntry
        {
            public string Name { get; set; }
            public string RelativePath { get; set; } // relative to macroFolder, directories use relative path
            public bool IsFolder { get; set; }

            public override string ToString()
            {
                return Name;
            }
        }

        // ルート/指定フォルダ内のディレクトリ（先）→ファイル（後）を返す
        public List<MacroEntry> GetMacroEntries(string relativePath = "")
        {
            var entries = new List<MacroEntry>();
            try
            {
                string target = string.IsNullOrEmpty(relativePath) ? macroFolder : Path.Combine(macroFolder, relativePath);
                if (!Directory.Exists(target))
                    return entries;

                // directories first
                var dirs = Directory.GetDirectories(target);
                Array.Sort(dirs, StringComparer.CurrentCultureIgnoreCase);
                foreach (var d in dirs)
                {
                    var dirName = Path.GetFileName(d);
                    var rel = string.IsNullOrEmpty(relativePath) ? dirName : Path.Combine(relativePath, dirName);
                    entries.Add(new MacroEntry { Name = dirName, RelativePath = rel, IsFolder = true });
                }

                // then files
                var files = Directory.GetFiles(target, "*.csv");
                Array.Sort(files, StringComparer.CurrentCultureIgnoreCase);
                foreach (var f in files)
                {
                    var name = Path.GetFileNameWithoutExtension(f);
                    var rel = string.IsNullOrEmpty(relativePath) ? name : Path.Combine(relativePath, name);
                    entries.Add(new MacroEntry { Name = name, RelativePath = rel, IsFolder = false });
                }
            }
            catch
            {
                // IO エラーは空リストでフォールバック
            }
            return entries;
        }

        // 互換: 既存呼び出し向けにフォルダ指定版の名前取得
        public List<string> GetMacroNames(string relativePath = "")
        {
            var result = new List<string>();
            try
            {
                string target = string.IsNullOrEmpty(relativePath) ? macroFolder : Path.Combine(macroFolder, relativePath);
                if (!Directory.Exists(target))
                    return result;
                var files = Directory.GetFiles(target, "*.csv");
                foreach (var file in files)
                {
                    var name = Path.GetFileNameWithoutExtension(file);
                    result.Add(name);
                }
            }
            catch
            {
            }
            return result;
        }

        // 既存: ルートのマクロテキスト読み込み（互換）
        public string LoadMacroText(string macroName)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            if (File.Exists(path))
                return File.ReadAllText(path, Encoding.UTF8);
            return "";
        }

        // 新: 相対フォルダ内のマクロテキスト読み込み
        public string LoadMacroText(string relativePath, string macroName)
        {
            try
            {
                string target = string.IsNullOrEmpty(relativePath) ? macroFolder : Path.Combine(macroFolder, relativePath);
                string path = Path.Combine(target, macroName + ".csv");
                if (File.Exists(path))
                    return File.ReadAllText(path, Encoding.UTF8);
            }
            catch
            {
            }
            return "";
        }

        // 既存: ルートへ保存（互換）
        public void SaveMacro(string macroName, string text)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            File.WriteAllText(path, text, Encoding.UTF8);
        }

        // 新: 相対フォルダへ保存（存在しなければ作成）
        public void SaveMacroInFolder(string relativePath, string macroName, string text)
        {
            try
            {
                string target = string.IsNullOrEmpty(relativePath) ? macroFolder : Path.Combine(macroFolder, relativePath);
                if (!Directory.Exists(target))
                    Directory.CreateDirectory(target);
                string path = Path.Combine(target, macroName + ".csv");
                File.WriteAllText(path, text, Encoding.UTF8);
            }
            catch
            {
                // 失敗は呼び出し元で処理
                throw;
            }
        }

        // 既存: ルート削除
        public void DeleteMacro(string macroName)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            if (File.Exists(path))
                File.Delete(path);
        }

        // 新: 相対フォルダ内削除
        public void DeleteMacroInFolder(string relativePath, string macroName)
        {
            try
            {
                string target = string.IsNullOrEmpty(relativePath) ? macroFolder : Path.Combine(macroFolder, relativePath);
                string path = Path.Combine(target, macroName + ".csv");
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
            }
        }
    }
}