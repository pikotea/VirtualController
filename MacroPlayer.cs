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

        // --- マクロファイルのパース処理を移動 ---
        public List<Dictionary<string, string>> ParseToFrameArray(string macroName)
        {
            string path = Path.Combine(macroFolder, macroName + ".csv");
            var macroFrames = MacroFrame.LoadMacroFile(path);

            var frameArray = new List<Dictionary<string, string>>();
            foreach (var frame in macroFrames)
            {
                // WAITフレームは空辞書
                if (frame.WaitFrames > 0)
                {
                    for (int i = 0; i < frame.WaitFrames; i++)
                        frameArray.Add(new Dictionary<string, string>());
                }
                else
                {
                    frameArray.Add(new Dictionary<string, string>(frame.KeyOps));
                }
            }
            return frameArray;
        }

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

                    var frameArray = ParseToFrameArray(innerMacroName);
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
                            if (val == "ON")
                                keyState.KeyStates[key] = true;
                            else if (val == "OFF")
                                keyState.KeyStates[key] = false;
                        }

                        short y = 0, x = 0;
                        if (keyState.KeyStates["UP"] && !keyState.KeyStates["DOWN"]) y = short.MaxValue;
                        else if (keyState.KeyStates["DOWN"] && !keyState.KeyStates["UP"]) y = short.MinValue;

                        if (!xAxisReverse)
                        {
                            if (keyState.KeyStates["LEFT"] && !keyState.KeyStates["RIGHT"]) x = short.MinValue;
                            else if (keyState.KeyStates["RIGHT"] && !keyState.KeyStates["LEFT"]) x = short.MaxValue;
                        }
                        else
                        {
                            if (keyState.KeyStates["LEFT"] && !keyState.KeyStates["RIGHT"]) x = short.MaxValue;
                            else if (keyState.KeyStates["RIGHT"] && !keyState.KeyStates["LEFT"]) x = short.MinValue;
                        }

                        controllerService.SetInputs(
                            new Dictionary<Xbox360Axis, short>
                            {
                                { Xbox360Axis.LeftThumbY, y },
                                { Xbox360Axis.LeftThumbX, x }
                            },
                            null,
                            frame
                        );

                        controllerService.SetInputs(
                            null,
                            new Dictionary<Xbox360Button, bool>
                            {
                                { Xbox360Button.A, keyState.KeyStates["A"] },
                                { Xbox360Button.B, keyState.KeyStates["B"] },
                                { Xbox360Button.X, keyState.KeyStates["X"] },
                                { Xbox360Button.Y, keyState.KeyStates["Y"] },
                                { Xbox360Button.LeftShoulder, keyState.KeyStates["LB"] },
                                { Xbox360Button.RightShoulder, keyState.KeyStates["RB"] }
                            },
                            frame
                        );

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
            }, token);
        }

        // MacroKeyStateを内包
        private class MacroKeyState
        {
            public Dictionary<string, bool> KeyStates = new Dictionary<string, bool>
            {
                { "UP", false }, { "DOWN", false }, { "LEFT", false }, { "RIGHT", false },
                { "A", false }, { "B", false }, { "X", false }, { "Y", false }, { "LB", false }, { "RB", false }
            };
        }

        // MacroFrameクラスをここに定義
        public class MacroFrame
        {
            public Dictionary<string, string> KeyOps = new Dictionary<string, string>();
            public int WaitFrames { get; set; } = 0;

            // 1行パース
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

            // ファイル全体パース
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
        }
    }
}