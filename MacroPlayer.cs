using Nefarius.ViGEm.Client.Targets.Xbox360;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace VirtualController
{
    public class MacroPlayer
    {
        private readonly ControllerService controllerService;
        private readonly string macroFolder;

        public MacroPlayer(ControllerService controllerService, string macroFolder)
        {
            this.controllerService = controllerService;
            this.macroFolder = macroFolder;
        }

        // --- マクロファイルのパース処理を元の仕様に合わせて修正 ---
        public List<Dictionary<string, string>> ParseToFrameArray(string macroName)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            return MacroFrame.ParseToFrameArray(path);
        }

        // PlayAsync の Task.Run 内部を元の仕様に合わせて修正
        public async Task PlayAsync(
            List<string> macroNames,
            double frameMs,
            int playWaitFrames,
            bool isRepeat,
            bool isRandom,
            bool xAxisReverse,
            int frame,
            CancellationToken token)
        {
            await Task.Run(() =>
            {
                try
                {
                    do
                    {
                        double totalWaitMs = playWaitFrames * frameMs;
                        if (totalWaitMs > 0)
                        {
                            if (token.IsCancellationRequested) break;
                            Thread.Sleep((int)totalWaitMs);
                        }

                        string innerMacroName = null;
                        if (isRandom && macroNames.Count > 0)
                        {
                            var rand = new Random();
                            innerMacroName = macroNames[rand.Next(macroNames.Count)];
                        }
                        else if (macroNames.Count > 0)
                        {
                            innerMacroName = macroNames[0];
                        }

                        if (string.IsNullOrEmpty(innerMacroName)) break;

                        string path = Path.Combine(macroFolder, innerMacroName + ".csv");
                        var frameArray = MacroFrame.ParseToFrameArray(path);
                        var keyState = new MacroKeyState();

                        int i = 0;
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        double nextFrameTime = sw.Elapsed.TotalMilliseconds;
                        while (i < frameArray.Count)
                        {
                            if (token.IsCancellationRequested) break;

                            int waitCount = 0;
                            while (i < frameArray.Count && frameArray[i].Count == 0)
                            {
                                waitCount++;
                                i++;
                            }
                            if (waitCount > 0)
                            {
                                nextFrameTime += waitCount * frameMs;
                                while (sw.Elapsed.TotalMilliseconds < nextFrameTime)
                                {
                                    if (token.IsCancellationRequested) break;
                                    Thread.Sleep(1);
                                }
                                continue;
                            }

                            var frameDict = frameArray[i];
                            foreach (var kv in frameDict)
                            {
                                string key = kv.Key.ToUpper();
                                string val = kv.Value.ToUpper();
                                // --- エイリアス対応 ---
                                if (key == "LP") key = "X";
                                else if (key == "LK") key = "A";
                                else if (key == "MP") key = "Y";
                                else if (key == "MK") key = "B";
                                else if (key == "HP") key = "RB";
                                else if (key == "HK") key = "LB";
                                if (val == "ON")
                                    keyState.KeyStates[key] = true;
                                else if (val == "OFF")
                                    keyState.KeyStates[key] = false;
                            }

                            short y = 0, x = 0;
                            if (keyState.KeyStates["UP"] && !keyState.KeyStates["DOWN"]) y = short.MaxValue;
                            else if (keyState.KeyStates["DOWN"] && !keyState.KeyStates["UP"]) y = short.MinValue;

                            bool reverse = xAxisReverse;
                            // 必要ならInvokeでUIから取得
                            // this.Invoke((Action)(() => { reverse = XAxisReverseCheckBox.Checked; }));

                            if (!reverse)
                            {
                                if (keyState.KeyStates["LEFT"] && !keyState.KeyStates["RIGHT"]) x = short.MinValue;
                                else if (keyState.KeyStates["RIGHT"] && !keyState.KeyStates["LEFT"]) x = short.MaxValue;
                            }
                            else
                            {
                                if (keyState.KeyStates["LEFT"] && !keyState.KeyStates["RIGHT"]) x = short.MaxValue;
                                else if (keyState.KeyStates["RIGHT"] && !keyState.KeyStates["LEFT"]) x = short.MinValue;
                            }

                            // 軸・ボタンを直接コントローラーに反映
                            controllerService.Controller.SetAxisValue(Xbox360Axis.LeftThumbY, y);
                            controllerService.Controller.SetAxisValue(Xbox360Axis.LeftThumbX, x);

                            controllerService.Controller.SetButtonState(Xbox360Button.A, keyState.KeyStates["A"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.B, keyState.KeyStates["B"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.X, keyState.KeyStates["X"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.Y, keyState.KeyStates["Y"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.LeftShoulder, keyState.KeyStates["LB"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.RightShoulder, keyState.KeyStates["RB"]);

                            nextFrameTime += frameMs;
                            while (sw.Elapsed.TotalMilliseconds < nextFrameTime)
                            {
                                if (token.IsCancellationRequested) break;
                                Thread.Sleep(1);
                            }
                            i++;
                        }
                        if (i == frameArray.Count)
                        {
                            nextFrameTime += frameMs;
                            while (sw.Elapsed.TotalMilliseconds < nextFrameTime)
                            {
                                if (token.IsCancellationRequested) break;
                                Thread.Sleep(1);
                            }
                        }
                        controllerService.AllOff();
                    } while (isRepeat && !token.IsCancellationRequested);
                }
                finally
                {
                    // UIを元に戻す
                    // UIスレッドでInvokeする必要がある場合は呼び出し元で対応
                }
            }, token);
        }

        // MacroKeyStateを内包
        private class MacroKeyState
        {
            public Dictionary<string, bool> KeyStates = new Dictionary<string, bool>
            {
                { "UP", false }, { "DOWN", false }, { "LEFT", false }, { "RIGHT", false },
                { "A", false }, { "B", false }, { "X", false }, { "Y", false }, { "LB", false }, { "RB", false },
                // --- エイリアス追加 ---
                { "LP", false }, // Xのエイリアス
                { "LK", false }, // Aのエイリアス
                { "MP", false }, // Yのエイリアス
                { "MK", false }, // Bのエイリアス
                { "HP", false }, // RBのエイリアス
                { "HK", false }  // LBのエイリアス
            };
        }

        // MacroFrameクラスを元の仕様に合わせて修正 ---
        public class MacroFrame
        {
            public Dictionary<string, string> KeyOps = new Dictionary<string, string>();
            public int WaitFrames { get; set; } = 0;

            // 新しい実行形式パース
            public static List<Dictionary<string, string>> ParseToFrameArray(string path)
            {
                var lines = File.ReadAllLines(path);

                // --- コメント行（#から始まる行）を除去 ---
                var filteredLines = lines
                    .Where(line => !line.TrimStart().StartsWith("#"))
                    .ToArray();

                var frameCounts = new List<int>();
                var frameKeys = new List<List<string>>();
                int totalFrames = 0;

                foreach (var line in filteredLines)
                {
                    var trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed))
                        continue;

                    int frameCount = 1;
                    string keyPart = trimmed;

                    // 行頭のフレーム数を抽出
                    int colonIdx = trimmed.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var frameStr = trimmed.Substring(0, colonIdx).Trim();
                        if (!int.TryParse(frameStr, out frameCount) || frameCount < 1)
                            frameCount = 1;
                        keyPart = trimmed.Substring(colonIdx + 1).Trim();
                    }

                    // キーリストを抽出
                    var keys = keyPart.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(k => k.Trim().ToUpper())
                                      .Where(k => !string.IsNullOrEmpty(k))
                                      .ToList();

                    frameCounts.Add(frameCount);
                    frameKeys.Add(keys);
                    totalFrames += frameCount;
                }

                // 合計フレーム数で配列を作成
                var frameArray = new List<Dictionary<string, string>>();
                for (int i = 0; i < totalFrames; i++)
                    frameArray.Add(new Dictionary<string, string>());

                // フレーム操作カーソルで各行を反映
                int cursor = 0;
                var lastOnKeys = new HashSet<string>();
                for (int specIdx = 0; specIdx < frameCounts.Count; specIdx++)
                {
                    int frameCount = frameCounts[specIdx];
                    var keys = frameKeys[specIdx];

                    // 最初のフレームだけON
                    if (frameCount > 0)
                    {
                        foreach (var key in keys)
                        {
                            frameArray[cursor][key] = "ON";
                            lastOnKeys.Add(key);
                        }
                    }
                    // 2フレーム目以降は空（何もセットしない）

                    // OFFを次フレームにセット
                    int offFrameIdx = cursor + frameCount;
                    if (offFrameIdx < frameArray.Count)
                    {
                        foreach (var key in keys)
                        {
                            frameArray[offFrameIdx][key] = "OFF";
                            lastOnKeys.Remove(key);
                        }
                    }
                    cursor += frameCount;
                }

                // 最後にONになっているキーをOFFにするフレームを追加
                if (lastOnKeys.Count > 0)
                {
                    var offFrame = new Dictionary<string, string>();
                    foreach (var key in lastOnKeys)
                        offFrame[key] = "OFF";
                    frameArray.Add(offFrame);
                }

                return frameArray;
            }

            public static List<MacroFrame> LoadMacroFile(string path)
            {
                var result = new List<MacroFrame>();
                foreach (var line in File.ReadAllLines(path))
                {
                    var frame = Parse(line);
                    result.Add(frame);
                }
                return result;
            }

            public static MacroFrame Parse(string line)
            {
                var frame = new MacroFrame();
                if (string.IsNullOrWhiteSpace(line)) return frame;
                var items = line.Split(',');
                foreach (var item in items)
                {
                    var kv = item.Split('=');
                    if (kv.Length == 2)
                    {
                        string key = kv[0].Trim().ToUpper();
                        string val = kv[1].Trim().ToUpper();
                        if (key == "WAIT")
                        {
                            if (int.TryParse(val, out int wait))
                                frame.WaitFrames = wait;
                        }
                        else
                        {
                            frame.KeyOps[key] = val;
                        }
                    }
                    else if (kv.Length == 1)
                    {
                        string key = kv[0].Trim().ToUpper();
                        frame.KeyOps[key] = "ON";
                    }
                }
                return frame;
            }
        }
    }
}