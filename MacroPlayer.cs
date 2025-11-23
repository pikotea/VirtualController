using Nefarius.ViGEm.Client.Targets.Xbox360;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

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
        // 先に全マクロを読み込み・パースしてから Task.Run に渡す（ネスト再生時に再読み込みしない）
        public async Task PlayAsync(
            List<string> macroNames,
            double frameMs,
            int playWaitFrames,
            bool isRepeat,
            bool isRandom,
            bool xAxisReverse,
            int frame,
            CancellationToken token,
            bool manageTiming = true,
            Dictionary<string, List<Dictionary<string, string>>> preParsed = null)
        {
            // 先に必要なマクロをすべて読み込み・パースしておく（重複はまとめる）
            var parsedMacros = preParsed ?? new Dictionary<string, List<Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
            if (preParsed == null)
            {
                foreach (var name in macroNames.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var path = Path.Combine(macroFolder, name + ".csv");
                        parsedMacros[name] = MacroFrame.ParseToFrameArray(path);
                    }
                    catch
                    {
                        // ファイル読み込み失敗は無視しておく（呼び出し元でログ化されている想定）
                    }
                }
            }

            await Task.Run(() =>
            {
                Thread currentThread = null;
                ThreadPriority originalPriority = ThreadPriority.Normal;
                if (manageTiming)
                {
                    // 高分解能タイマーを使用開始（内部で参照カウント管理）
                    TimingHelpers.BeginHighResolution();

                    // 再生中はスレッド優先度を一時的に上げる（注意: 短時間に限定）
                    currentThread = Thread.CurrentThread;
                    originalPriority = currentThread.Priority;
                    currentThread.Priority = ThreadPriority.Highest;
                }

                try
                {
                    try
                    {
                        do
                        {
                            double totalWaitMs = playWaitFrames * frameMs;
                            if (totalWaitMs > 0)
                            {
                                if (token.IsCancellationRequested) break;
                                // 従来互換で単純待機（短時間の待機はハイブリッド待機を使わない）
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

                            // ここでは事前にパース済みデータを使う（なければフォールバックしてパース）
                            List<Dictionary<string, string>> frameArray = null;
                            if (!parsedMacros.TryGetValue(innerMacroName, out frameArray))
                            {
                                try
                                {
                                    var path = Path.Combine(macroFolder, innerMacroName + ".csv");
                                    frameArray = MacroFrame.ParseToFrameArray(path);
                                    parsedMacros[innerMacroName] = frameArray;
                                }
                                catch
                                {
                                    frameArray = new List<Dictionary<string, string>>();
                                }
                            }

                            var keyState = new MacroKeyState();

                            int i = 0;
                            var sw = Stopwatch.StartNew();
                            double nextFrameTime = sw.Elapsed.TotalMilliseconds;

                            // 計測用
                            double maxJitterMs = 0;
                            double sumAbsJitter = 0;
                            long sampleCount = 0;

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
                                    TimingHelpers.WaitUntil(sw, nextFrameTime, token);
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

                                // 次フレーム予定時刻を更新してハイブリッドで待機
                                nextFrameTime += frameMs;
                                TimingHelpers.WaitUntil(sw, nextFrameTime, token);

                                // 計測: 実際到達時刻と予定時刻の差を記録（ミリ秒）
                                double actual = sw.Elapsed.TotalMilliseconds;
                                double jitter = actual - nextFrameTime;
                                sumAbsJitter += Math.Abs(jitter);
                                if (Math.Abs(jitter) > maxJitterMs) maxJitterMs = Math.Abs(jitter);
                                sampleCount++;

                                i++;
                            }
                            if (i == frameArray.Count)
                            {
                                nextFrameTime += frameMs;
                                TimingHelpers.WaitUntil(sw, nextFrameTime, token);
                            }
                            controllerService.AllOff();

                            // 計測結果をログ出力（デバッグ）
                            if (sampleCount > 0)
                                Debug.WriteLine($"[MacroPlayer] timing samples={sampleCount}, maxJitterMs={maxJitterMs:F3}, avgAbsJitterMs={(sumAbsJitter / sampleCount):F3}");
                        } while (isRepeat && !token.IsCancellationRequested);
                    }
                    finally
                    {
                        // UIを元に戻す
                        // UIスレッドでInvokeする必要がある場合は呼び出し元で対応
                    }
                }
                finally
                {
                    // 優先度と高分解能タイマーを元に戻す（manageTiming=true の場合のみ）
                    if (manageTiming)
                    {
                        try { if (currentThread != null) currentThread.Priority = originalPriority; } catch { }
                        TimingHelpers.EndHighResolution();
                    }
                }
            }, token);
        }

        // PlayAsyncの引数にジョイスティックと設定を追加
        // ここでも先にすべてのマクロを読み込み・パースしておく（トリガー解析と再生で再利用）
        public async Task PlayAsync(
            List<string> macroNames,
            double frameMs,
            int playWaitFrames,
            bool isRepeat,
            bool isRandom,
            bool xAxisReverse,
            int frame,
            CancellationToken token,
            SharpDX.DirectInput.Joystick joystick = null,
            RecordSettingsForm.RecordSettingsConfig recordConfig = null,
            bool manageTiming = true,
            Dictionary<string, List<Dictionary<string, string>>> preParsed = null)
        {
            var macroTriggers = new List<(string macroName, List<string> triggers, List<string> waitActions)>();

            // 事前に全マクロをパース（必要に応じてフォールバック）
            var parsedMacros = preParsed ?? new Dictionary<string, List<Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
            if (preParsed == null)
            {
                foreach (var name in macroNames.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var path = Path.Combine(macroFolder, name + ".csv");
                        parsedMacros[name] = MacroFrame.ParseToFrameArray(path);
                    }
                    catch
                    {
                        // 無視
                    }
                }
            }

            foreach (var macroName in macroNames)
            {
                var macroText = File.ReadAllText(Path.Combine(macroFolder, macroName + ".csv"));
                var lines = macroText.Split('\n');
                var triggers = new List<string>();
                var waitActions = new List<string>();
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("#TRIGGER:"))
                    {
                        triggers = trimmed.Substring("#TRIGGER:".Length)
                            .Split(',')
                            .Select(x =>
                            {
                                var key = x.Trim().ToUpper();
                                if (xAxisReverse)
                                {
                                    if (key == "LEFT") key = "RIGHT";
                                    else if (key == "RIGHT") key = "LEFT";
                                }
                                return key;
                            })
                            .Where(x => !string.IsNullOrEmpty(x))
                            .ToList();
                    }
                    if (trimmed.StartsWith("#TRIGGERWAIT:"))
                    {
                        waitActions = trimmed.Substring("#TRIGGERWAIT:".Length)
                            .Split(',')
                            .Select(x =>
                            {
                                var key = x.Trim().ToUpper();
                                if (xAxisReverse)
                                {
                                    if (key == "LEFT") key = "RIGHT";
                                    else if (key == "RIGHT") key = "LEFT";
                                }
                                return key;
                            })
                            .Where(x => !string.IsNullOrEmpty(x))
                            .ToList();
                    }
                    if (!trimmed.StartsWith("#") && (trimmed.Contains(":") || (trimmed.Length > 0 && !trimmed.StartsWith("#"))))
                        break;
                }
                macroTriggers.Add((macroName, triggers, waitActions));
            }

            var triggerMacros = macroTriggers.Where(mt => mt.triggers.Count > 0).ToList();
            System.Diagnostics.Debug.WriteLine("[MacroPlayer] triggerMacros:");
            foreach (var mt in triggerMacros)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"  macroName={mt.macroName}, triggers=[{string.Join(",", mt.triggers)}], waitActions=[{string.Join(",", mt.waitActions)}]"
                );
            }
            System.Diagnostics.Debug.WriteLine($"[MacroPlayer] joystick: {(joystick == null ? "null" : joystick.ToString())}");
            System.Diagnostics.Debug.WriteLine($"[MacroPlayer] recordConfig: {(recordConfig == null ? "null" : $"ControllerGuid={recordConfig.ControllerGuid}, ButtonIndices={string.Join(",", recordConfig.ButtonIndices.Select(kv => kv.Key + "=" + (kv.Value?.ToString() ?? "null")))} ZValues={string.Join(",", recordConfig.ZValues.Select(kv => kv.Key + "=" + (kv.Value?.ToString() ?? "null")))}")}");

            // 判定部分
            if (triggerMacros.Count == 0 || joystick == null || recordConfig == null)
            {
                System.Diagnostics.Debug.WriteLine("[MacroPlayer] PlayAsync(基本) 実行");
                await PlayAsync(
                    macroNames,
                    frameMs,
                    playWaitFrames,
                    isRepeat,
                    isRandom,
                    xAxisReverse,
                    frame,
                    token,
                    manageTiming,
                    parsedMacros // 事前パース済みを渡す
                );
                return;
            }

            // すべてのTRIGGERWAIT定義から最初のボタン群を抽出
            List<string> prioritizedWaitActions = null;
            foreach (var mt in triggerMacros)
            {
                if (mt.waitActions.Count > 0)
                {
                    prioritizedWaitActions = mt.waitActions;
                    break;
                }
            }

            // 高精度ループに変更：Task.Run 内で高分解能タイマーとスレッド優先度を使用
            await Task.Run(async () =>
            {
                Thread currentThread = null;
                ThreadPriority originalPriority = ThreadPriority.Normal;
                if (manageTiming)
                {
                    TimingHelpers.BeginHighResolution();
                    currentThread = Thread.CurrentThread;
                    originalPriority = currentThread.Priority;
                    currentThread.Priority = ThreadPriority.Highest;
                }

                try
                {
                    var sw = Stopwatch.StartNew();
                    double nextFrameTime = sw.Elapsed.TotalMilliseconds;

                    Random rand = new Random();
                    while (!token.IsCancellationRequested)
                    {
                        // トリガー待ちの間は prioritizedWaitActions のボタンを押す（毎フレーム更新）
                        if (prioritizedWaitActions != null && prioritizedWaitActions.Count > 0)
                        {
                            var keyState = new MacroKeyState();
                            foreach (var kv in prioritizedWaitActions)
                            {
                                string key = kv.ToUpper();
                                if (key == "LP") key = "X";
                                else if (key == "LK") key = "A";
                                else if (key == "MP") key = "Y";
                                else if (key == "MK") key = "B";
                                else if (key == "HP") key = "RB";
                                else if (key == "HK") key = "LB";
                                keyState.KeyStates[key] = true;
                            }

                            short y = 0, x = 0;
                            if (keyState.KeyStates["UP"] && !keyState.KeyStates["DOWN"]) y = short.MaxValue;
                            else if (keyState.KeyStates["DOWN"] && !keyState.KeyStates["UP"]) y = short.MinValue;

                            bool reverse = xAxisReverse;
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

                            controllerService.Controller.SetAxisValue(Xbox360Axis.LeftThumbY, y);
                            controllerService.Controller.SetAxisValue(Xbox360Axis.LeftThumbX, x);

                            controllerService.Controller.SetButtonState(Xbox360Button.A, keyState.KeyStates["A"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.B, keyState.KeyStates["B"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.X, keyState.KeyStates["X"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.Y, keyState.KeyStates["Y"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.LeftShoulder, keyState.KeyStates["LB"]);
                            controllerService.Controller.SetButtonState(Xbox360Button.RightShoulder, keyState.KeyStates["RB"]);
                        }

                        bool triggered = false;
                        while (!triggered && !token.IsCancellationRequested)
                        {
                            var state = joystick.GetCurrentState();
                            var pressedKeys = new List<string>();

                            foreach (var kv in recordConfig.ButtonIndices)
                            {
                                if (kv.Value != null && state.Buttons.Length > kv.Value.Value && state.Buttons[kv.Value.Value])
                                    pressedKeys.Add(kv.Key.Replace("Button", ""));
                            }
                            foreach (var kv in recordConfig.ZValues)
                            {
                                if (kv.Value != null && state.Z == kv.Value.Value)
                                    pressedKeys.Add(kv.Key.Replace("Button", ""));
                            }
                            if (state.PointOfViewControllers != null && state.PointOfViewControllers.Length > 0)
                            {
                                int pov = state.PointOfViewControllers[0];
                                if (pov == 0) pressedKeys.Add("UP");
                                else if (pov == 9000) pressedKeys.Add("RIGHT");
                                else if (pov == 18000) pressedKeys.Add("DOWN");
                                else if (pov == 27000) pressedKeys.Add("LEFT");
                                else if (pov == 4500) { pressedKeys.Add("UP"); pressedKeys.Add("RIGHT"); }
                                else if (pov == 13500) { pressedKeys.Add("DOWN"); pressedKeys.Add("RIGHT"); }
                                else if (pov == 22500) { pressedKeys.Add("DOWN"); pressedKeys.Add("LEFT"); }
                                else if (pov == 31500) { pressedKeys.Add("UP"); pressedKeys.Add("LEFT"); }
                            }

                            var triggerdMacros = new List<string>();
                            foreach (var mt in triggerMacros)
                            {
                                if (mt.triggers.All(t => pressedKeys.Contains(t)))
                                {
                                    triggerdMacros.Add(mt.macroName);
                                }
                            }

                            if (triggerdMacros.Count > 0)
                            {
                                controllerService.AllOff();
                                System.Diagnostics.Debug.WriteLine("[MacroPlayer] PlayAsync(ジョイスティック/設定付き) 実行");

                                // parsedMacros からトリガー対象のパース済みデータを抽出して渡す（ネスト再生で再読み込みしない）
                                var parsedSubset = new Dictionary<string, List<Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
                                foreach (var nm in triggerdMacros)
                                {
                                    if (parsedMacros.TryGetValue(nm, out var arr))
                                        parsedSubset[nm] = arr;
                                }

                                // ネストして再生（manageTiming=false を渡して Begin/End/Priority の重複を回避）
                                await PlayAsync(
                                    triggerdMacros,
                                    frameMs, 0, false, isRandom, xAxisReverse, frame, token, false, parsedSubset
                                );
                                triggered = true;
                            }

                            // 次チェックまで高精度で待機（frameMs に同期）
                            nextFrameTime += frameMs / 3;
                            TimingHelpers.WaitUntil(sw, nextFrameTime, token);
                        }
                    }
                }
                finally
                {
                    if (manageTiming)
                    {
                        try { if (currentThread != null) currentThread.Priority = originalPriority; } catch { }
                        TimingHelpers.EndHighResolution();
                    }
                }
            }, token);
            return;
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