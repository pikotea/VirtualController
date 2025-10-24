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

        // マクロ一覧取得
        public List<string> GetMacroNames()
        {
            var result = new List<string>();
            var files = Directory.GetFiles(macroFolder, "*.csv");
            foreach (var file in files)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                result.Add(name);
            }
            return result;
        }

        // マクロファイルの内容取得
        public string LoadMacroText(string macroName)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            if (File.Exists(path))
                return File.ReadAllText(path, Encoding.UTF8);
            return "";
        }

        // マクロ保存（新規・上書き）
        public void SaveMacro(string macroName, string text)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            File.WriteAllText(path, text, Encoding.UTF8);
        }

        // マクロ削除
        public void DeleteMacro(string macroName)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}