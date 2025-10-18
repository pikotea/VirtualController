using Nefarius.ViGEm.Client;
using Nefarius.ViGEm.Client.Targets;
using Nefarius.ViGEm.Client.Targets.Xbox360;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace VirtualController
{
    public partial class Form1 : Form
    {
        ViGEmClient client;
        IXbox360Controller controller;
        int frame = 16;
        string macroFolder = Path.Combine(Application.StartupPath, "micros");
        List<MacroFrame> loadedMacro = new List<MacroFrame>();
        FileSystemWatcher macroWatcher;

        // 再生・停止制御用
        private volatile bool isMacroPlaying = false;
        private CancellationTokenSource macroCancelSource;

        private string editingMacroName = null;
        private string lastLoadedMacroText = "";

        private const string SettingsFile = "settings.ini";

        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }


        public Form1()
        {
            InitializeComponent();
            client = new ViGEmClient();
            controller = this.client.CreateXbox360Controller();
            if (!Directory.Exists(macroFolder)) Directory.CreateDirectory(macroFolder);
            StartMacroFolderWatcher();
            LoadMacroList();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.SetControllerButtonsEnabled(false);
            PlayMacroButton.Text = "再生";
            LoadMacroSettings();
        }

        // マクロ一覧のリロード
        private void LoadMacroList(bool stopMacro = true)
        {
            // マクロ一覧更新時は自動停止（条件付き）
            if (stopMacro)
                StopMacroIfPlaying();

            // 現在の選択状態を保存
            var selectedNames = MacroListBox.SelectedItems.Cast<string>().ToList();

            MacroListBox.Items.Clear();
            var files = Directory.GetFiles(macroFolder, "*.csv");
            foreach (var file in files)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                MacroListBox.Items.Add(name);
            }

            // 選択状態を復元
            MacroListBox.ClearSelected();
            foreach (var name in selectedNames)
            {
                int idx = MacroListBox.Items.IndexOf(name);
                if (idx >= 0)
                    MacroListBox.SetSelected(idx, true);
            }
        }


        // マクロ再生
        private void PlayMacroButton_Click(object sender, EventArgs e)
        {
            if (isMacroPlaying) return;
            if (MacroListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("マクロが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // --- 追加: マクロ編集内容に変更があれば上書き保存 ---
            if (!string.IsNullOrEmpty(editingMacroName) && MacroEditTextBox.Text != lastLoadedMacroText)
            {
                OverwriteMacro();
            }

            // UI制御
            PlayMacroButton.Enabled = false;
            PlayMacroButton.Text = "再生中...";
            StopMacroButton.Enabled = true;
            isMacroPlaying = true;
            macroCancelSource = new CancellationTokenSource();
            var token = macroCancelSource.Token;

            double frameMs = 1000 / 60;
            int playWaitFrames = 0;
            bool isRepeat = RepeatCheckBox.Checked;
            bool isRandom = RandomCheckBox.Checked;
            double.TryParse(FrameMsTextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out frameMs);
            int.TryParse(PlayWaitTextBox.Text, out playWaitFrames);

            // 選択マクロ名リスト
            var macroNames = MacroListBox.SelectedItems.Cast<string>().ToList();

            // ここで全選択状態を解除
            this.Invoke((Action)(() => {
                MacroEditTextBox.SelectionLength = 0;
                MacroEditTextBox.SelectionStart = MacroEditTextBox.TextLength;
            }));

            Task.Run(() =>
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
                                if (val == "ON")
                                    keyState.KeyStates[key] = true;
                                else if (val == "OFF")
                                    keyState.KeyStates[key] = false;
                            }

                            short y = 0, x = 0;
                            if (keyState.KeyStates["UP"] && !keyState.KeyStates["DOWN"]) y = short.MaxValue;
                            else if (keyState.KeyStates["DOWN"] && !keyState.KeyStates["UP"]) y = short.MinValue;

                            bool reverse = false;
                            this.Invoke((Action)(() => { reverse = XAxisReverseCheckBox.Checked; }));

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

                            controller.SetAxisValue(Xbox360Axis.LeftThumbY, y);
                            controller.SetAxisValue(Xbox360Axis.LeftThumbX, x);

                            controller.SetButtonState(Xbox360Button.A, keyState.KeyStates["A"]);
                            controller.SetButtonState(Xbox360Button.B, keyState.KeyStates["B"]);
                            controller.SetButtonState(Xbox360Button.X, keyState.KeyStates["X"]);
                            controller.SetButtonState(Xbox360Button.Y, keyState.KeyStates["Y"]);
                            controller.SetButtonState(Xbox360Button.LeftShoulder, keyState.KeyStates["LB"]);
                            controller.SetButtonState(Xbox360Button.RightShoulder, keyState.KeyStates["RB"]);

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
                        MacroFrame.AllOff(controller);

                    } while (isRepeat && !token.IsCancellationRequested);
                }
                finally
                {
                    // UIを元に戻す
                    this.Invoke((Action)(() =>
                    {
                        PlayMacroButton.Enabled = true;
                        PlayMacroButton.Text = "再生";
                        StopMacroButton.Enabled = false;
                        isMacroPlaying = false;
                    }));
                }
            });
        }

        private void StopMacroButton_Click(object sender, EventArgs e)
        {
            if (!isMacroPlaying) return;
            macroCancelSource?.Cancel();
            MacroFrame.AllOff(controller);
        }

        // フォルダ監視でマクロ一覧リロード
        private void StartMacroFolderWatcher()
        {
            macroWatcher = new FileSystemWatcher(macroFolder, "*.csv");
            macroWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            macroWatcher.Changed += (s, e) => this.Invoke((Action)(() => LoadMacroList()));
            macroWatcher.Created += (s, e) => this.Invoke((Action)(() => LoadMacroList()));
            macroWatcher.Deleted += (s, e) => this.Invoke((Action)(() => LoadMacroList()));
            macroWatcher.Renamed += (s, e) => this.Invoke((Action)(() => LoadMacroList()));
            macroWatcher.EnableRaisingEvents = true;
        }

        // 接続
        private void ConnectButton_Click(object sender, EventArgs e)
        {
            this.controller.Connect();
            this.SetControllerButtonsEnabled(true);
        }

        // 切断
        private void DisconnectButton_Click(object sender, EventArgs e)
        {
            this.controller.Disconnect();
            this.SetControllerButtonsEnabled(false);
        }

        /// <summary>
        /// 接続・切断以外のコントローラ操作ボタンを有効/無効化
        /// </summary>
        private void SetControllerButtonsEnabled(bool enabled)
        {
            // 接続・切断ボタン以外を対象
            foreach (var btn in this.Controls.OfType<Button>())
            {
                if (btn.Name == "ConnectButton")
                {
                    btn.Enabled = !enabled;
                }
                else if (btn.Name == "NewMacroButton" || btn.Name == "SaveAsButton" || btn.Name == "OverwriteSaveButton")
                {
                    // 新規作成・保存系ボタンは独立
                }
                else if (btn.Name == "PlayMacroButton")
                {
                    // 再生ボタンは「接続済み」かつ「マクロ選択済み」の場合のみ有効
                    btn.Enabled = enabled && MacroListBox.SelectedItems.Count > 0;
                }
                else if (btn.Name == "StopMacroButton")
                {
                    // 停止ボタンは再生中のみ有効
                    btn.Enabled = isMacroPlaying;
                }
                else if (btn.Name == "DebugMacroButton")
                {
                    // デバッグボタンはマクロ選択状態でのみ有効
                    btn.Enabled = MacroListBox.SelectedItems.Count > 0;
                }
                else if (btn.Name == "OpenMacroFolderButton")
                {
                    // フォルダを開くボタンは常に有効
                    btn.Enabled = true;
                }
                else
                {
                    btn.Enabled = enabled;
                }
            }
        }

        // 上
        private void UpButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbY, short.MaxValue },
                },
                null,
                frame
            );
        }

        // 下
        private void DownButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbY, short.MinValue },
                },
                null,
                frame
            );
        }

        // 左
        private void LeftButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MinValue },
                },
                null,
                frame
            );
        }

        // 右
        private void RightButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MaxValue },
                },
                null,
                frame
            );
        }


        

        // Aボタン
        private void AButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.A, true }
                },
                frame
            );
        }

        // Bボタン
        private void BButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.B, true }
                },
                frame
            );
        }

        //　Xボタン
        private void XButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.X, true }
                },
                frame
            );
        }

        // Yボタン
        private void YButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.Y, true }
                },
                frame
            );
        }

        // RBボタン
        private void RBButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.RightShoulder, true }
                },
                frame
            );
        }


        // LBボタン
        private void LBButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.LeftShoulder, true }
                },
                frame
            );
        }


        // STARTボタン
        private void StartButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.Start, true }
                },
                frame
            );
        }

        // BACKボタン
        private void BackButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                null,
                new Dictionary<Xbox360Button, bool>
                {
                    { Xbox360Button.Back, true }
                },
                frame
            );
        }

        // 右下
        private void DownRightButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MaxValue },
                    { Xbox360Axis.LeftThumbY, short.MinValue },
                },
                null,
                frame
            );
        }

        // 左上
        private void UpLeftButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MinValue },
                    { Xbox360Axis.LeftThumbY, short.MaxValue },
                },
                null,
                frame
            );
        }

        // 右上
        private void UpRightButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MaxValue },
                    { Xbox360Axis.LeftThumbY, short.MaxValue },
                },
                null,
                frame
            );
        }

        // 左下
        private void DownLeftButton_Click(object sender, EventArgs e)
        {
            ControllerExtension.SetInputs(
                controller,
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MinValue },
                    { Xbox360Axis.LeftThumbY, short.MinValue },
                },
                null,
                frame
            );
        }

        private void Macro1Button_Click(object sender, EventArgs e)
        {
            var macroButtons = new[]
            {
                Xbox360Button.X,
                Xbox360Button.A,
                Xbox360Button.Y,
                Xbox360Button.B,
                Xbox360Button.RightShoulder,
                Xbox360Button.LeftShoulder
            };

            foreach (var btn in macroButtons)
            {
                ControllerExtension.SetInputs(
                    controller,
                    null,
                    new Dictionary<Xbox360Button, bool> { { btn, true } },
                    frame
                );
            }
        }

        // マクロ選択時のロード（複数選択対応・編集エリア表示）
        // --- マクロ選択時にラベル表示・保存 ---
        private void MacroListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // マクロ選択変更時は自動停止
            StopMacroIfPlaying();

            loadedMacro.Clear();
            string lastMacroName = null;
            bool macroLoaded = false;
            foreach (var macroNameObj in MacroListBox.SelectedItems)
            {
                string macroName = macroNameObj as string;
                if (string.IsNullOrEmpty(macroName)) continue;
                string path = Path.Combine(macroFolder, macroName + ".csv");
                var macroFrames = MacroFrame.LoadMacroFile(path);
                if (macroFrames != null && macroFrames.Count > 0)
                {
                    loadedMacro.AddRange(macroFrames);
                    macroLoaded = true;
                }
                lastMacroName = macroName;
            }
            // 編集エリアに内容表示
            if (lastMacroName != null)
            {
                string path = Path.Combine(macroFolder, lastMacroName + ".csv");
                MacroEditTextBox.Text = File.ReadAllText(path);
                MacroEditTextBox.Enabled = true;
                editingMacroName = lastMacroName;
                lastLoadedMacroText = MacroEditTextBox.Text;
                MacroNameLabel.Text = lastMacroName; // ラベルに表示
            }
            else
            {
                MacroEditTextBox.Text = "";
                MacroEditTextBox.Enabled = true;
                editingMacroName = null;
                lastLoadedMacroText = "";
                MacroNameLabel.Text = "";
            }
            // 名前を付けて保存は常に有効
            SaveAsButton.Enabled = true;
            // 上書き保存は初期状態では無効
            OverwriteSaveButton.Enabled = false;
            // 再生ボタンの有効化条件を変更
            PlayMacroButton.Enabled = macroLoaded && controller != null && DisconnectButton.Enabled;

            DeleteMicroButton.Enabled = MacroListBox.SelectedItems.Count > 0;

            // デバッグボタンの有効化条件を変更
            DebugMacroButton.Enabled = MacroListBox.SelectedItems.Count > 0;

            SaveMacroSettings();
        }

        // 編集エリア変更時に上書き保存ボタンの有効/無効を制御
        private void MacroEditTextBox_TextChanged(object sender, EventArgs e)
        {
            // テキスト変更時は自動停止
            StopMacroIfPlaying();

            SaveAsButton.Enabled = true;
            if (!string.IsNullOrEmpty(editingMacroName) && MacroEditTextBox.Text != lastLoadedMacroText)
                OverwriteSaveButton.Enabled = true;
            else
                OverwriteSaveButton.Enabled = false;
        }

        // 保存ボタン押下時の処理
        private void SaveMacroButton_Click(object sender, EventArgs e)
        {
            string macroName = null;
            if (MacroListBox.SelectedItems.Count > 0)
            {
                macroName = MacroListBox.SelectedItems[MacroListBox.SelectedItems.Count - 1] as string;
            }
            if (string.IsNullOrEmpty(macroName))
            {
                // 新規マクロ名生成
                int idx = 1;
                do
                {
                    macroName = $"新規マクロ{idx}";
                    idx++;
                } while (File.Exists(Path.Combine(macroFolder, macroName + ".csv")));
            }
            string defaultPath = Path.Combine(macroFolder, macroName + ".csv");

            using (var sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = macroFolder;
                sfd.FileName = macroName + ".csv";
                sfd.Filter = "CSVファイル (*.csv)|*.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, MacroEditTextBox.Text, Encoding.UTF8);
                    LoadMacroList();
                    SaveAsButton.Enabled = false;
                }
            }
        }

        private void NewMacroButton_Click(object sender, EventArgs e)
        {
            if (MacroEditTextBox.Enabled && MacroEditTextBox.Text != lastLoadedMacroText && !string.IsNullOrEmpty(lastLoadedMacroText))
            {
                var result = MessageBox.Show("編集中のマクロがあります。上書き保存しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    OverwriteMacro();
                }
            }
            MacroEditTextBox.Text = "";
            MacroEditTextBox.Enabled = true;
            editingMacroName = null;
            lastLoadedMacroText = "";
            OverwriteSaveButton.Enabled = false;
            SaveAsButton.Enabled = true;
        }

        private void OverwriteSaveButton_Click(object sender, EventArgs e)
        {
            OverwriteMacro();
        }

        // --- 上書き保存処理 ---
        private void OverwriteMacro()
        {
            if (!string.IsNullOrEmpty(editingMacroName))
            {
                string path = Path.Combine(macroFolder, editingMacroName + ".csv");
                File.WriteAllText(path, MacroEditTextBox.Text, Encoding.UTF8);
                lastLoadedMacroText = MacroEditTextBox.Text;
                LoadMacroList(false); // 停止せず一覧更新
                OverwriteSaveButton.Enabled = false;

                this.Invoke((Action)(() => {
                    MacroEditTextBox.SelectionLength = 0;
                    MacroEditTextBox.SelectionStart = MacroEditTextBox.TextLength;
                }));
            }
        }

        // --- 設定保存: 複数選択対応（カンマ区切りで保存） ---
        private void SaveMacroSettings()
        {
            // 選択されているマクロ名を取得
            var selectedMacros = new List<string>();
            foreach (int idx in MacroListBox.SelectedIndices)
            {
                selectedMacros.Add(MacroListBox.Items[idx].ToString());
            }
            var selectedMacroStr = string.Join(",", selectedMacros);
            var lines = new[]
            {
                $"FrameMs={FrameMsTextBox.Text}",
                $"PlayWait={PlayWaitTextBox.Text}",
                $"Repeat={RepeatCheckBox.Checked}",
                $"Random={RandomCheckBox.Checked}",
                $"XAxisReverse={XAxisReverseCheckBox.Checked}", // 追加
                $"SelectedMacros={selectedMacroStr}"
            };
            File.WriteAllLines(Path.Combine(Application.StartupPath, SettingsFile), lines, Encoding.UTF8);
        }

        // --- 設定ロード: 複数選択復元 ---
        private void LoadMacroSettings()
        {
            string path = Path.Combine(Application.StartupPath, SettingsFile);
            if (!File.Exists(path)) return;
            string[] selectedMacros = null;
            foreach (var line in File.ReadAllLines(path))
            {
                var kv = line.Split('=');
                if (kv.Length != 2) continue;
                string key = kv[0].Trim();
                string val = kv[1].Trim();
                switch (key)
                {
                    case "FrameMs":
                        FrameMsTextBox.Text = val;
                        break;
                    case "PlayWait":
                        PlayWaitTextBox.Text = val;
                        break;
                    case "Repeat":
                        RepeatCheckBox.Checked = val == "True";
                        break;
                    case "Random":
                        RandomCheckBox.Checked = val == "True";
                        break;
                    case "XAxisReverse": // 追加
                        XAxisReverseCheckBox.Checked = val == "True";
                        break;
                    case "SelectedMacros":
                        selectedMacros = val.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        break;
                }
            }
            // マクロ一覧ロード後に複数選択復元
            if (selectedMacros != null && selectedMacros.Length > 0)
            {
                MacroListBox.ClearSelected();
                foreach (var macro in selectedMacros)
                {
                    int idx = MacroListBox.Items.IndexOf(macro);
                    if (idx >= 0)
                        MacroListBox.SetSelected(idx, true);
                }
            }
        }


        // --- 各コントロールの変更イベントで保存を呼び出す ---
        private void FrameMsTextBox_TextChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();
            SaveMacroSettings();
        }
        private void PlayWaitTextBox_TextChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();
            SaveMacroSettings();
        }
        private void RepeatCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();
            SaveMacroSettings();
        }
        private void RandomCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();
            SaveMacroSettings();
        }
        private void XAxisReverseCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();
            SaveMacroSettings();
        }

        private void DebugMacroButton_Click(object sender, EventArgs e)
        {
            if (MacroListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("マクロが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string macroName = MacroListBox.SelectedItems[0] as string;
            if (string.IsNullOrEmpty(macroName)) return;

            string path = Path.Combine(macroFolder, macroName + ".csv");
            var frameArray = MacroFrame.ParseToFrameArray(path);

            var sb = new StringBuilder();
            sb.AppendLine($"マクロ名: {macroName}");
            sb.AppendLine($"フレーム数: {frameArray.Count}");
            for (int i = 0; i < frameArray.Count; i++)
            {
                sb.Append($"[{i + 1}] ");
                if (frameArray[i].Count == 0)
                {
                    sb.AppendLine("(WAIT)");
                }
                else
                {
                    foreach (var kv in frameArray[i])
                    {
                        sb.Append($"{kv.Key}={kv.Value} ");
                    }
                    sb.AppendLine();
                }
            }

            // 結果表示
            var debugForm = new Form
            {
                Text = "フレームデータ",
                Width = 600,
                Height = 400
            };
            var textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Both,
                Text = sb.ToString()
            };
            debugForm.Controls.Add(textBox);

            var copyButton = new Button
            {
                Text = "コピー",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            copyButton.Click += (s, ev) =>
            {
                Clipboard.SetText(textBox.Text);
                MessageBox.Show("クリップボードにコピーしました。", "コピー", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            debugForm.Controls.Add(copyButton);

            debugForm.ShowDialog();
        }

        private void DeleteMicroButton_Click(object sender, EventArgs e)
        {
            if (MacroListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("削除するマクロが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string macroName = MacroListBox.SelectedItems[0] as string;
            if (string.IsNullOrEmpty(macroName)) return;

            var result = MessageBox.Show(
                $"マクロ「{macroName}」を削除しますか？\nこの操作は元に戻せません。",
                "確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                string path = Path.Combine(macroFolder, macroName + ".csv");
                if (File.Exists(path))
                {
                    File.Delete(path);
                    LoadMacroList();
                    // 選択状態が外れるので削除ボタンを無効化
                    DeleteMicroButton.Enabled = false;
                    // 削除されたマクロの内容をクリア
                    MacroEditTextBox.Text = "";
                    MacroNameLabel.Text = "";
                    editingMacroName = null;
                    lastLoadedMacroText = "";
                    MessageBox.Show("マクロを削除しました。", "削除完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// マクロ再生中なら停止する
        /// </summary>
        private void StopMacroIfPlaying()
        {
            if (isMacroPlaying)
            {
                macroCancelSource?.Cancel();
                MacroFrame.AllOff(controller);
                isMacroPlaying = false;
                this.Invoke((Action)(() =>
                {
                    PlayMacroButton.Enabled = true;
                    PlayMacroButton.Text = "再生";
                    StopMacroButton.Enabled = false;
                }));
            }
        }

        private void OpenMacroFolderButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(macroFolder))
                {
                    Directory.CreateDirectory(macroFolder);
                }
                System.Diagnostics.Process.Start("explorer.exe", macroFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show("フォルダを開けませんでした。\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

public static class ControllerExtension
{

    /// <summary>
    /// 複数の軸・ボタンを同時に押す（duration後に全て離す）
    /// </summary>
    public static void SetInputs(
        IXbox360Controller controller,
        Dictionary<Xbox360Axis, short> axisValues,
        Dictionary<Xbox360Button, bool> buttonStates,
        int duration)
    {
        // 軸を同時にセット
        if (axisValues != null)
        {
            foreach (var kvp in axisValues)
            {
                controller.SetAxisValue(kvp.Key, kvp.Value);
            }
        }
        // ボタンを同時にセット
        if (buttonStates != null)
        {
            foreach (var kvp in buttonStates)
            {
                controller.SetButtonState(kvp.Key, kvp.Value);
            }
        }
        Thread.Sleep(duration);
        // 軸を全て0に戻す
        if (axisValues != null)
        {
            foreach (var kvp in axisValues)
            {
                controller.SetAxisValue(kvp.Key, 0);
            }
        }
        // ボタンを全て離す
        if (buttonStates != null)
        {
            foreach (var kvp in buttonStates)
            {
                controller.SetButtonState(kvp.Key, false);
            }
        }
    }
}

// キーの状態管理（継続ON/OFF用）
public class MacroKeyState
{
    public Dictionary<string, bool> KeyStates = new Dictionary<string, bool>
    {
        { "UP", false }, { "DOWN", false }, { "LEFT", false }, { "RIGHT", false },
        { "A", false }, { "B", false }, { "X", false }, { "Y", false }, { "LB", false }, { "RB", false }
    };
}
// 1フレーム分の入力情報
public class MacroFrame
{
    // 各フレームのキー状態（ON/OFFのみ。WAITはブランクで表現）
    public Dictionary<string, string> KeyOps = new Dictionary<string, string>();

    // 追加: WaitFrames プロパティ
    public int WaitFrames { get; set; } = 0;

    // --- 新しい実行形式パース ---
    public static List<Dictionary<string, string>> ParseToFrameArray(string path)
    {
        var lines = File.ReadAllLines(path);

        var frameCounts = new List<int>();
        var frameKeys = new List<List<string>>();
        int totalFrames = 0;

        foreach (var line in lines)
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
                string key = kv[0].Trim().ToUpper(); // キーを大文字化
                string val = kv[1].Trim().ToUpper(); // 値も大文字化
                if (key == "WAIT")
                {
                    if (int.TryParse(val, out int wait))
                        frame.WaitFrames = wait;
                }
                else
                {
                    frame.KeyOps[key] = val; // ON/OFF
                }
            }
            else if (kv.Length == 1)
            {
                string key = kv[0].Trim().ToUpper(); // キーを大文字化
                frame.KeyOps[key] = "ON"; // 省略形はON
            }
        }
        return frame;
    }

    public static void AllOff(IXbox360Controller controller)
    {
        controller.SetAxisValue(Xbox360Axis.LeftThumbY, 0);
        controller.SetAxisValue(Xbox360Axis.LeftThumbX, 0);
        controller.SetButtonState(Xbox360Button.A, false);
        controller.SetButtonState(Xbox360Button.B, false);
        controller.SetButtonState(Xbox360Button.X, false);
        controller.SetButtonState(Xbox360Button.Y, false);
        controller.SetButtonState(Xbox360Button.LeftShoulder, false);
        controller.SetButtonState(Xbox360Button.RightShoulder, false);
    }

    // FrameSpecクラスを外に定義
    private class FrameSpec
    {
        public Dictionary<string, int> KeyDurations;
        public int WaitFrames;
        public FrameSpec(Dictionary<string, int> keyDurations, int waitFrames)
        {
            KeyDurations = keyDurations;
            WaitFrames = waitFrames;
        }
    }
}

