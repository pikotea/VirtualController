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

        // 新: マクロのメタデータ + フレーム配列を保持するコンテナ
        public class MacroData
        {
            public Dictionary<string, object> PropsRaw { get; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            public List<string> Triggers { get; set; }
            public List<string> TriggerWaits { get; set; }
            public int? Interval { get; set; }
            public double? Frame { get; set; }
            public int? Random { get; set; }
            public int? Loop { get; set; }
            public List<Dictionary<string, string>> Frames { get; } = new List<Dictionary<string, string>>();

            // 追加: パースエラー情報
            public bool HasParseError { get; set; } = false;
            public string ParseErrorMessage { get; set; }
        }

        // 追加: マクロファイルを読み、props（コメント行）と frames を返すヘルパー
        private MacroData ParseMacroFileWithProps(string path, bool applyAxisReverse = false)
        {
            var data = new MacroData();

            string[] lines;
            try
            {
                lines = File.ReadAllLines(path);
            }
            catch (Exception ex)
            {
                data.HasParseError = true;
                data.ParseErrorMessage = $"マクロファイルを開けませんでした: {Path.GetFileName(path)}\n{ex.Message}";
                return data;
            }

            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (!line.StartsWith("#")) continue;
                var idx = line.IndexOf(':');
                if (idx < 0) continue;
                var key = line.Substring(1, idx - 1).Trim().ToUpper();
                var val = line.Substring(idx + 1).Trim();

                data.PropsRaw[key] = val;

                if (key == "TRIGGER")
                {
                    var items = val.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(s => s.Trim().ToUpper())
                                   .Where(s => !string.IsNullOrEmpty(s))
                                   .ToList();
                    if (applyAxisReverse)
                    {
                        for (int i = 0; i < items.Count; i++)
                        {
                            if (items[i] == "LEFT") items[i] = "RIGHT";
                            else if (items[i] == "RIGHT") items[i] = "LEFT";
                        }
                    }
                    data.Triggers = items;
                }
                else if (key == "TRIGGERWAIT")
                {
                    var items = val.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(s => s.Trim().ToUpper())
                                   .Where(s => !string.IsNullOrEmpty(s))
                                   .ToList();
                    if (applyAxisReverse)
                    {
                        for (int i = 0; i < items.Count; i++)
                        {
                            if (items[i] == "LEFT") items[i] = "RIGHT";
                            else if (items[i] == "RIGHT") items[i] = "LEFT";
                        }
                    }
                    data.TriggerWaits = items;
                }
                else if (key == "INTERVAL")
                {
                    if (int.TryParse(val, out int v)) data.Interval = v;
                }
                else if (key == "FRAME")
                {
                    if (double.TryParse(val, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double dv)) data.Frame = dv;
                }
                else if (key == "RANDOM")
                {
                    if (int.TryParse(val, out int v)) data.Random = v;
                }
                else if (key == "LOOP")
                {
                    if (int.TryParse(val, out int v)) data.Loop = v;
                }
            }

            // frames のパースは検証付きで行う
            if (!MacroFrame.TryParseToFrameArray(path, out var frames, out var error))
            {
                data.HasParseError = true;
                data.ParseErrorMessage = $"マクロのパースに失敗しました: {Path.GetFileName(path)}\n{error}";
                return data;
            }

            data.Frames.AddRange(frames);
            return data;
        }

        // PlayAsync (基本) - preParsed の型を MacroData に変更
        public async Task PlayAsync(
            List<string> macroNames,
            double frameMs,
            int playWaitFrames,
            bool isRepeat,
            bool isRandom,
            bool xAxisReverse,
            CancellationToken token,
            bool manageTiming = true,
            Dictionary<string, MacroData> preParsed = null)
        {
            var parsedMacros = preParsed ?? new Dictionary<string, MacroData>(StringComparer.OrdinalIgnoreCase);
            if (preParsed == null)
            {
                foreach (var name in macroNames.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var path = Path.Combine(macroFolder, name + ".csv");
                        parsedMacros[name] = ParseMacroFileWithProps(path, applyAxisReverse: xAxisReverse);
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            // 事前パースでエラーがあれば例外を投げて呼び出し元に処理を任せる
            foreach (var kv in parsedMacros)
            {
                if (kv.Value != null && kv.Value.HasParseError)
                {
                    throw new InvalidDataException(kv.Value.ParseErrorMessage ?? $"マクロのパースエラー: {kv.Key}")
;
                }
            }

            await Task.Run(() =>
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

                            // ランダム再生の重み付き選択をサポート
                            // Random の値を使い、デフォルトは 1（1 未満は 1 として扱う）
                            Random rand = null;
                            if (isRandom && macroNames.Count > 0)
                            {
                                rand = new Random();
                                // 重みを計算
                                var weights = new int[macroNames.Count];
                                int totalWeight = 0;
                                for (int wi = 0; wi < macroNames.Count; wi++)
                                {
                                    int w = 1;
                                    if (parsedMacros.TryGetValue(macroNames[wi], out var md2) && md2 != null && md2.Random.HasValue)
                                    {
                                        if (md2.Random.Value >= 1) w = md2.Random.Value;
                                        else w = 1;
                                    }
                                    weights[wi] = w;
                                    totalWeight += w;
                                }

                                if (totalWeight <= 0)
                                {
                                    innerMacroName = macroNames[0];
                                }
                                else
                                {
                                    int r = rand.Next(totalWeight);
                                    int acc = 0;
                                    for (int wi = 0; wi < macroNames.Count; wi++)
                                    {
                                        acc += weights[wi];
                                        if (r < acc)
                                        {
                                            innerMacroName = macroNames[wi];
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (macroNames.Count > 0)
                            {
                                innerMacroName = macroNames[0];
                            }

                            if (string.IsNullOrEmpty(innerMacroName)) break;

                            MacroData md = null;
                            if (!parsedMacros.TryGetValue(innerMacroName, out md))
                            {
                                try
                                {
                                    var path = Path.Combine(macroFolder, innerMacroName + ".csv");
                                    md = ParseMacroFileWithProps(path, applyAxisReverse: xAxisReverse);
                                    parsedMacros[innerMacroName] = md;
                                }
                                catch
                                {
                                    md = new MacroData();
                                }

                                if (md.HasParseError)
                                {
                                    // パースエラーは例外を投げる
                                    throw new InvalidDataException(md.ParseErrorMessage ?? $"マクロのパースエラー: {innerMacroName}");
                                }
                            }

                            var frameArray = md?.Frames ?? new List<Dictionary<string, string>>();

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
                                    else if (key == "HK") key = "RT";
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

                                // 再生ループでボタン/軸をコントローラに反映する箇所に追加
                                // (buttonStates の反映のあと、トリガーを反映)
                                bool ltOn = keyState.KeyStates.ContainsKey("LT") && keyState.KeyStates["LT"];
                                bool rtOn = keyState.KeyStates.ContainsKey("RT") && keyState.KeyStates["RT"];
                                controllerService.Controller.SetSliderValue(Xbox360Slider.LeftTrigger, ltOn ? byte.MaxValue : (byte)0);
                                controllerService.Controller.SetSliderValue(Xbox360Slider.RightTrigger, rtOn ? byte.MaxValue : (byte)0);

                                controllerService.Controller.SubmitReport();

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

        // PlayAsync (joystick-aware) - preParsed 型を MacroData に変更
        public async Task PlayAsync(
            List<string> macroNames,
            double frameMs,
            int playWaitFrames,
            bool isRepeat,
            bool isRandom,
            bool xAxisReverse,
            CancellationToken token,
            SharpDX.DirectInput.Joystick joystick = null,
            RecordSettingsForm.RecordSettingsConfig recordConfig = null,
            bool manageTiming = true,
            Dictionary<string, MacroData> preParsed = null)
        {
            var macroTriggers = new List<(string macroName, List<string> triggers, List<string> waitActions)>();

            var parsedMacros = preParsed ?? new Dictionary<string, MacroData>(StringComparer.OrdinalIgnoreCase);
            if (preParsed == null)
            {
                foreach (var name in macroNames.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var path = Path.Combine(macroFolder, name + ".csv");
                        parsedMacros[name] = ParseMacroFileWithProps(path, applyAxisReverse: xAxisReverse);
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            // 事前パースでエラーがあればユーザーへ通知して再生を中止
            foreach (var kv in parsedMacros)
            {
                if (kv.Value != null && kv.Value.HasParseError)
                {
                    // エラー表示は呼び出し元の UI (MainForm) で行うため、ここでは例外を投げる
                    throw new InvalidDataException(kv.Value.ParseErrorMessage ?? $"マクロのパースエラー: {kv.Key}");
                }
            }

            foreach (var macroName in macroNames)
            {
                MacroData md = null;
                parsedMacros.TryGetValue(macroName, out md);
                var triggers = md?.Triggers ?? new List<string>();
                var waitActions = md?.TriggerWaits ?? new List<string>();
                macroTriggers.Add((macroName, triggers, waitActions));
            }

            var triggerMacros = macroTriggers.Where(mt => mt.triggers != null && mt.triggers.Count > 0).ToList();
            Debug.WriteLine("[MacroPlayer] triggerMacros:");
            foreach (var mt in triggerMacros)
            {
                Debug.WriteLine($"  macroName={mt.macroName}, triggers=[{string.Join(",", mt.triggers)}], waitActions=[{string.Join(",", mt.waitActions)}]");
            }
            Debug.WriteLine($"[MacroPlayer] joystick: {(joystick == null ? "null" : joystick.ToString())}");
            Debug.WriteLine($"[MacroPlayer] recordConfig: {(recordConfig == null ? "null" : $"ControllerGuid={recordConfig.ControllerGuid}, ButtonIndices={string.Join(",", recordConfig.ButtonIndices.Select(kv => kv.Key + "=" + (kv.Value?.ToString() ?? "null")))} ZValues={string.Join(",", recordConfig.ZValues.Select(kv => kv.Key + "=" + (kv.Value?.ToString() ?? "null")))} ")} ");

            if (triggerMacros.Count == 0 || joystick == null || recordConfig == null)
            {
                Debug.WriteLine("[MacroPlayer] PlayAsync(基本) 実行");
                await PlayAsync(
                    macroNames,
                    frameMs,
                    playWaitFrames,
                    isRepeat,
                    isRandom,
                    xAxisReverse,
                    token,
                    manageTiming,
                    parsedMacros
                );
                return;
            }

            List<string> prioritizedWaitActions = null;
            foreach (var mt in triggerMacros)
            {
                if (mt.waitActions != null && mt.waitActions.Count > 0)
                {
                    prioritizedWaitActions = mt.waitActions;
                    break;
                }
            }

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
                   Debug.WriteLine("[MacroPlayer] TRIGGERキー待ち状態");
                   var sw = Stopwatch.StartNew();
                   double nextFrameTime = sw.Elapsed.TotalMilliseconds;

                   while (!token.IsCancellationRequested)
                   {
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
                               else if (key == "HK") key = "RT";
                               keyState.KeyStates[key] = true;
                           }

                           short y = 0, x = 0;
                           if (keyState.KeyStates["UP"] && !keyState.KeyStates["DOWN"]) y = short.MaxValue;
                           else if (keyState.KeyStates["DOWN"] && !keyState.KeyStates["UP"]) y = short.MinValue;

                           if (keyState.KeyStates["LEFT"] && !keyState.KeyStates["RIGHT"]) x = short.MinValue;
                           else if (keyState.KeyStates["RIGHT"] && !keyState.KeyStates["LEFT"]) x = short.MaxValue;


                           controllerService.Controller.SetAxisValue(Xbox360Axis.LeftThumbY, y);
                           controllerService.Controller.SetAxisValue(Xbox360Axis.LeftThumbX, x);

                           controllerService.Controller.SetButtonState(Xbox360Button.A, keyState.KeyStates["A"]);
                           controllerService.Controller.SetButtonState(Xbox360Button.B, keyState.KeyStates["B"]);
                           controllerService.Controller.SetButtonState(Xbox360Button.X, keyState.KeyStates["X"]);
                           controllerService.Controller.SetButtonState(Xbox360Button.Y, keyState.KeyStates["Y"]);
                           controllerService.Controller.SetButtonState(Xbox360Button.LeftShoulder, keyState.KeyStates["LB"]);
                           controllerService.Controller.SetButtonState(Xbox360Button.RightShoulder, keyState.KeyStates["RB"]);

                           bool ltOn = keyState.KeyStates.ContainsKey("LT") && keyState.KeyStates["LT"];
                           bool rtOn = keyState.KeyStates.ContainsKey("RT") && keyState.KeyStates["RT"];
                           controllerService.Controller.SetSliderValue(Xbox360Slider.LeftTrigger, ltOn ? byte.MaxValue : (byte)0);
                           controllerService.Controller.SetSliderValue(Xbox360Slider.RightTrigger, rtOn ? byte.MaxValue : (byte)0);
                           controllerService.Controller.SubmitReport();
                        }

                       bool triggered = false;
                       while (!triggered && !token.IsCancellationRequested)
                       {
                           try
                           {
                               var state = joystick.GetCurrentState();
                               var pressedKeys = new List<string>();

                               foreach (var kv in recordConfig.ButtonIndices)
                               {
                                   if (kv.Value != null && state.Buttons != null && state.Buttons.Length > kv.Value.Value && state.Buttons[kv.Value.Value])
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
                                   Debug.WriteLine("[MacroPlayer] PlayAsync(ジョイスティック/設定付き) 実行");

                                   var parsedSubset = new Dictionary<string, MacroData>(StringComparer.OrdinalIgnoreCase);
                                   foreach (var nm in triggerdMacros)
                                   {
                                       if (parsedMacros.TryGetValue(nm, out var md2))
                                           parsedSubset[nm] = md2;
                                   }

                                   await PlayAsync(
                                       triggerdMacros,
                                       frameMs, 0, false, isRandom, xAxisReverse, token, false, parsedSubset
                                   );
                                   triggered = true;
                               }
                           }
                           catch (Exception ex)
                           {
                               Debug.WriteLine($"[MacroPlayer] joystick 取得/判定で例外: {ex}");
                           }

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
            // KeyStates 辞書に LT, RT を追加
            public Dictionary<string, bool> KeyStates = new Dictionary<string, bool>
            {
                { "UP", false }, { "DOWN", false }, { "LEFT", false }, { "RIGHT", false },
                { "A", false }, { "B", false }, { "X", false }, { "Y", false }, { "LT", false }, { "RT", false }, { "LB", false }, { "RB", false },
                { "LP", false }, { "LK", false }, { "MP", false }, { "MK", false }, { "HP", false },{ "HK", false } 
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
                    else
                    {
                        // 行が数字のみなら待機指定 (例: "5" -> 5 フレームの待機)
                        if (int.TryParse(trimmed, out var numericCount) && numericCount >= 1)
                        {
                            frameCount = numericCount;
                            keyPart = "";
                        }
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

            // 追加: 検証付きパース。失敗時は false を返し error にメッセージを設定する
            public static bool TryParseToFrameArray(string path, out List<Dictionary<string, string>> frames, out string error)
            {
                frames = new List<Dictionary<string, string>>();
                error = null;
                string[] lines;
                try
                {
                    lines = File.ReadAllLines(path);
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }

                var filteredLines = lines
                    .Where(line => !line.TrimStart().StartsWith("#"))
                    .ToArray();

                var frameCounts = new List<int>();
                var frameKeys = new List<List<string>>();
                int totalFrames = 0;

                int lineNo = 0;
                foreach (var line in filteredLines)
                {
                    lineNo++;
                    var trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed))
                        continue;

                    int frameCount = 1;
                    string keyPart = trimmed;

                    int colonIdx = trimmed.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var frameStr = trimmed.Substring(0, colonIdx).Trim();
                        if (!int.TryParse(frameStr, out frameCount) || frameCount < 1)
                        {
                            error = $"無効なフレーム数指定（行 {lineNo}）: '{frameStr}'";
                            return false;
                        }
                        keyPart = trimmed.Substring(colonIdx + 1).Trim();
                    }
                    else
                    {
                        // 行が数字のみなら待機指定 (例: "5" -> 5 フレームの待機)
                        if (int.TryParse(trimmed, out var numericCount) && numericCount >= 1)
                        {
                            frameCount = numericCount;
                            keyPart = "";
                        }
                    }

                    // キーリストを抽出
                    var keys = keyPart.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(k => k.Trim().ToUpper())
                                      .Where(k => !string.IsNullOrEmpty(k))
                                      .ToList();

                    // キー指定は任意にする（空行や待機のみの行を許容）
                    // もしキーが無い場合は空のキーリストとして扱い、エラーにしない。
                    // 検証: 指定されたキーが MacroKeyState に定義されているか、および矛盾がないかをチェック
                    var allowedKeys = new HashSet<string>(new MacroKeyState().KeyStates.Keys, StringComparer.OrdinalIgnoreCase);
                    if (keys.Count > 0)
                    {
                        // 左右同時指定禁止
                        if (keys.Contains("LEFT") && keys.Contains("RIGHT"))
                        {
                            error = $"左右が同時に指定されています（行 {lineNo}）: '{line}'";
                            return false;
                        }
                        // 上下同時指定禁止
                        if (keys.Contains("UP") && keys.Contains("DOWN"))
                        {
                            error = $"上下が同時に指定されています（行 {lineNo}）: '{line}'";
                            return false;
                        }

                        // 未知のキー禁止
                        foreach (var k in keys)
                        {
                            if (!allowedKeys.Contains(k))
                            {
                                error = $"不明なキーが指定されています（行 {lineNo}）: '{k}'";
                                return false;
                            }
                        }
                    }

                    frameCounts.Add(frameCount);
                    frameKeys.Add(keys);
                    totalFrames += frameCount;
                }

                for (int i = 0; i < totalFrames; i++)
                    frames.Add(new Dictionary<string, string>());

                int cursor = 0;
                var lastOnKeys = new HashSet<string>();
                for (int specIdx = 0; specIdx < frameCounts.Count; specIdx++)
                {
                    int frameCount = frameCounts[specIdx];
                    var keys = frameKeys[specIdx];
                    if (frameCount > 0)
                    {
                        foreach (var key in keys)
                        {
                            frames[cursor][key] = "ON";
                            lastOnKeys.Add(key);
                        }
                    }

                    int offFrameIdx = cursor + frameCount;
                    if (offFrameIdx < frames.Count)
                    {
                        foreach (var key in keys)
                        {
                            frames[offFrameIdx][key] = "OFF";
                            lastOnKeys.Remove(key);
                        }
                    }
                    cursor += frameCount;
                }

                if (lastOnKeys.Count > 0)
                {
                    var offFrame = new Dictionary<string, string>();
                    foreach (var key in lastOnKeys)
                        offFrame[key] = "OFF";
                    frames.Add(offFrame);
                }

                return true;
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