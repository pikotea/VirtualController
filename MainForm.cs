using Nefarius.ViGEm.Client.Targets.Xbox360;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VirtualController;


namespace VirtualController
{
    public partial class MainForm : Form
    {
        int frame = 16;
        string macroFolder = Path.Combine(Application.StartupPath, "micros");
        List<MacroPlayer.MacroFrame> loadedMacro = new List<MacroPlayer.MacroFrame>();
        FileSystemWatcher macroWatcher;

        // 再生・停止制御用
        private volatile bool isMacroPlaying = false;
        private CancellationTokenSource macroCancelSource;

        private string editingMacroName = null;
        private string lastLoadedMacroText = "";

        private const string SettingsFile = "settings.ini";

        // 1. フィールド追加
        private MacroManager macroManager;
        private ControllerService controllerService;
        private MacroPlayer macroPlayer;
        private SharpDX.DirectInput.DirectInput directInput;
        private SharpDX.DirectInput.Joystick currentJoystick;
        private RecordSettingsForm.RecordSettingsConfig recordConfig;

        // MainForm クラス内フィールド追加
        private bool isRecording = false;
        private List<Dictionary<string, string>> recordedFrames = new List<Dictionary<string, string>>();
        private DateTime recordStartTime;

        // 追加: 保存後に優先選択するマクロ名（拡張子なし）
        private string pendingSelectMacroName = null;

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


        // 2. コンストラクタで初期化
        // --- コンストラクタで自動接続 ---
        public MainForm()
        {
            InitializeComponent();
            controllerService = new ControllerService();
            controllerService.Connect(); // ← 起動時に自動接続
            macroManager = new MacroManager(macroFolder);
            macroPlayer = new MacroPlayer(controllerService, macroFolder);
            StartMacroFolderWatcher();
            LoadMacroList();
        }

        // --- 1. 起動時（Form1_Load）で設定ファイルを読み込む ---
        private void Form1_Load(object sender, EventArgs e)
        {
            // --- Form1_Loadでコントローラ操作ボタンを有効化 ---
            this.SetControllerButtonsEnabled(true); // ← 常に有効化
            PlayMacroButton.Text = "再生";
            LoadMacroSettings();
            recordConfig = LoadRecordSettingsConfig(); // ← 追加
            if (IsRecordSettingsValid())
            {
                ConnectRecordController();
            }
            UpdateRecButtonEnabled();
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
            var macroNames = macroManager.GetMacroNames(); // 変更
            foreach (var name in macroNames)
            {
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

            // 追加: pendingSelectMacroName があれば優先して選択（保存直後の非同期更新対策）
            if (!string.IsNullOrEmpty(pendingSelectMacroName))
            {
                int idx = MacroListBox.Items.IndexOf(pendingSelectMacroName);
                if (idx >= 0)
                {
                    MacroListBox.ClearSelected();
                    MacroListBox.SetSelected(idx, true);
                }
                pendingSelectMacroName = null;
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
            double.TryParse(FrameMsTextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out frameMs);

            int playWaitFrames = 0;
            int.TryParse(PlayWaitTextBox.Text, out playWaitFrames);
            bool isRepeat = RepeatCheckBox.Checked;
            bool isRandom = RandomCheckBox.Checked;


            // 選択マクロ名リスト
            var macroNames = MacroListBox.SelectedItems.Cast<string>().ToList();

            // ここで全選択状態を解除
            this.Invoke((Action)(() =>
            {
                MacroEditTextBox.SelectionLength = 0;
                MacroEditTextBox.SelectionStart = MacroEditTextBox.TextLength;
            }));

            // --- MacroPlayerを利用 ---
            Task.Run(async () =>
            {
                try
                {
                    await macroPlayer.PlayAsync(
                        macroNames,
                        frameMs,
                        playWaitFrames,
                        isRepeat,
                        isRandom,
                        XAxisReverseCheckBox.Checked,
                        frame,
                        token,
                        currentJoystick,
                        recordConfig
                    );
                }
                finally
                {
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
            controllerService.AllOff(); // 変更
        }

        // フォルダ監視でマクロ一覧リロード
        private void StartMacroFolderWatcher()
        {
            macroWatcher = new FileSystemWatcher(macroFolder, "*.csv");
            macroWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            macroWatcher.Changed += MacroFolderChanged;
            macroWatcher.Created += MacroFolderChanged;
            macroWatcher.Deleted += MacroFolderChanged;
            macroWatcher.Renamed += MacroFolderChanged;
            macroWatcher.EnableRaisingEvents = true;
        }

        // ファイル変更時のイベントハンドラ
        private void MacroFolderChanged(object sender, FileSystemEventArgs e)
        {
            // マクロ再生中なら停止
            if (isMacroPlaying)
            {
                this.Invoke((Action)(() => StopMacroIfPlaying()));
                // マクロ一覧は停止後に更新
                this.Invoke((Action)(() => LoadMacroList(false)));
            }
            else
            {
                this.Invoke((Action)(() => LoadMacroList(false)));
            }
        }

        /// <summary>
        /// コントローラ操作ボタンの有効/無効化（簡略化）
        /// </summary>
        private void SetControllerButtonsEnabled(bool enabled)
        {
            foreach (var btn in this.Controls.OfType<Button>())
            {
                if (btn.Name == "NewMacroButton" || btn.Name == "SaveAsButton" || btn.Name == "OverwriteSaveButton")
                {
                    // 新規作成・保存系ボタンは独立
                }
                else if (btn.Name == "PlayMacroButton")
                {
                    btn.Enabled = enabled && MacroListBox.SelectedItems.Count > 0;
                }
                else if (btn.Name == "StopMacroButton")
                {
                    btn.Enabled = isMacroPlaying;
                }
                else if (btn.Name == "DebugMacroButton")
                {
                    btn.Enabled = MacroListBox.SelectedItems.Count > 0;
                }
                else if (btn.Name == "OpenMacroFolderButton")
                {
                    btn.Enabled = true;
                }
                else if (btn.Name == "RecSettingButton")
                {
                    btn.Enabled = true;
                }
                else
                {
                    btn.Enabled = enabled;
                }
            }
        }


        // マクロ選択時のロード（複数選択対応・編集エリア表示）
        // --- マクロ選択時にラベル表示・保存 ---
        private void MacroListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();

            loadedMacro.Clear();
            string lastMacroName = null;
            foreach (var macroNameObj in MacroListBox.SelectedItems)
            {
                string macroName = macroNameObj as string;
                if (string.IsNullOrEmpty(macroName)) continue;
                // ここで自作MacroFrameのLoadMacroFileを使う
                var macroFrames = MacroPlayer.MacroFrame.LoadMacroFile(Path.Combine(macroFolder, macroName + ".csv"));
                if (macroFrames != null && macroFrames.Count > 0)
                {
                    loadedMacro.AddRange(macroFrames);
                }
                lastMacroName = macroName;
            }
            // 編集エリアに内容表示
            if (lastMacroName != null)
            {
                MacroEditTextBox.Text = macroManager.LoadMacroText(lastMacroName); // 変更
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

            PlayMacroButton.Enabled = MacroListBox.SelectedItems.Count > 0;

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

        // 保存ボタン押下時の処理（変更部分のみ）
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
                } while (macroManager.GetMacroNames().Contains(macroName));
            }
            using (var sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = macroFolder;
                sfd.FileName = macroName + ".csv";
                sfd.Filter = "CSVファイル (*.csv)|*.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var savedName = Path.GetFileNameWithoutExtension(sfd.FileName);
                    macroManager.SaveMacro(savedName, MacroEditTextBox.Text);


                    // マクロ一覧再読み込み
                    pendingSelectMacroName = savedName;
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

            // マクロ選択を解除
            MacroListBox.ClearSelected();

            // 編集エリアを新規状態に
            MacroEditTextBox.Text = "";
            MacroEditTextBox.Enabled = true;
            editingMacroName = null;
            lastLoadedMacroText = "";
            OverwriteSaveButton.Enabled = false;
            SaveAsButton.Enabled = true;

            // フォーカスを編集エリアへ移動し、キャレットを先頭に配置
            this.ActiveControl = MacroEditTextBox;
            MacroEditTextBox.SelectionStart = 0;
            MacroEditTextBox.SelectionLength = 0;
            MacroEditTextBox.Focus();

            // 設定保存（選択状態が変わったため必要に応じて保存）
            SaveMacroSettings();
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
                loadedMacro = MacroPlayer.MacroFrame.LoadMacroFile(path);
                LoadMacroList(false);
                OverwriteSaveButton.Enabled = false;
                this.Invoke((Action)(() =>
                {
                    MacroEditTextBox.SelectionLength = 0;
                    MacroEditTextBox.SelectionStart = MacroEditTextBox.TextLength;
                }));

                // 追加: 上書き保存後に優先選択するマクロ名を設定
                pendingSelectMacroName = editingMacroName;
            }
        }

        // 置き換え対象: 現在の SaveMacroSettings(...)
        // 戻す内容（引数なし）
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

        // --- DebugMacroButton_Click 内のマクロデータ表示処理を MacroPlayer に統一 ---
        private void DebugMacroButton_Click(object sender, EventArgs e)
        {
            if (MacroListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("マクロが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string macroName = MacroListBox.SelectedItems[0] as string;
            if (string.IsNullOrEmpty(macroName)) return;

            // MacroPlayerのParseToFrameArrayを利用
            var frameArray = macroPlayer.ParseToFrameArray(macroName);

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

        // 6. マクロ削除
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
                macroManager.DeleteMacro(macroName); // 変更
                LoadMacroList();
                DeleteMicroButton.Enabled = false;
                // 削除されたマクロの内容をクリア
                MacroEditTextBox.Text = "";
                MacroNameLabel.Text = "";
                editingMacroName = null;
                lastLoadedMacroText = "";
                MessageBox.Show("マクロを削除しました。", "削除完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                controllerService.AllOff(); // 変更
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

        // 上
        private void UpButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbY, short.MaxValue } },
                null);
        }
        private void UpButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbY, 0 } },
                null);
        }

        // 下
        private void DownButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbY, short.MinValue } },
                null);
        }
        private void DownButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbY, 0 } },
                null);
        }

        // 左
        private void LeftButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbX, short.MinValue } },
                null);
        }
        private void LeftButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbX, 0 } },
                null);
        }

        // 右
        private void RightButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbX, short.MaxValue } },
                null);
        }
        private void RightButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short> { { Xbox360Axis.LeftThumbX, 0 } },
                null);
        }

        // 右下
        private void DownRightButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MaxValue },
                    { Xbox360Axis.LeftThumbY, short.MinValue }
                },
                null);
        }
        private void DownRightButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, 0 },
                    { Xbox360Axis.LeftThumbY, 0 }
                },
                null);
        }

        // 左上
        private void UpLeftButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MinValue },
                    { Xbox360Axis.LeftThumbY, short.MaxValue }
                },
                null);
        }
        private void UpLeftButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, 0 },
                    { Xbox360Axis.LeftThumbY, 0 }
                },
                null);
        }

        // 右上
        private void UpRightButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MaxValue },
                    { Xbox360Axis.LeftThumbY, short.MaxValue }
                },
                null);
        }
        private void UpRightButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, 0 },
                    { Xbox360Axis.LeftThumbY, 0 }
                },
                null);
        }

        // 左下
        private void DownLeftButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, short.MinValue },
                    { Xbox360Axis.LeftThumbY, short.MinValue }
                },
                null);
        }
        private void DownLeftButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                new Dictionary<Xbox360Axis, short>
                {
                    { Xbox360Axis.LeftThumbX, 0 },
                    { Xbox360Axis.LeftThumbY, 0 }
                },
                null);
        }

        // Aボタン
        private void AButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.A, true } });
        }
        private void AButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.A, false } });
        }

        // Bボタン
        private void BButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.B, true } });
        }
        private void BButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.B, false } });
        }

        // Xボタン
        private void XButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.X, true } });
        }
        private void XButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.X, false } });
        }

        // Yボタン
        private void YButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.Y, true } });
        }
        private void YButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.Y, false } });
        }

        // RBボタン
        private void RBButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.RightShoulder, true } });
        }
        private void RBButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.RightShoulder, false } });
        }

        // LBボタン
        private void LBButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.LeftShoulder, true } });
        }
        private void LBButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.LeftShoulder, false } });
        }

        // STARTボタン
        private void StartButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.Start, true } });
        }
        private void StartButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.Start, false } });
        }

        // BACKボタン
        private void BackButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.Back, true } });
        }
        private void BackButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.SetInputs(
                null,
                new Dictionary<Xbox360Button, bool> { { Xbox360Button.Back, false } });
        }

        private void RecSettingButton_Click(object sender, EventArgs e)
        {
            using (var dlg = new RecordSettingsForm())
            {
                var result = dlg.ShowDialog(this);
                if (result == DialogResult.OK)
                {
                    recordConfig = LoadRecordSettingsConfig(); // ← 追加
                    if (IsRecordSettingsValid())
                    {
                        ConnectRecordController();
                    }
                    UpdateRecButtonEnabled();
                }
            }
        }


        // 記録設定が揃っているか判定するメソッド
        private bool IsRecordSettingsValid()
        {
            // RecordSettingsConfig.json の存在と内容チェック
            var configPath = Path.Combine(Application.StartupPath, "RecordSettingsConfig.json");
            if (!File.Exists(configPath)) return false;

            try
            {
                using (var fs = new FileStream(configPath, FileMode.Open))
                {
                    var serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(
                        typeof(RecordSettingsForm.RecordSettingsConfig),
                        new System.Runtime.Serialization.Json.DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true }
                    );
                    var config = (RecordSettingsForm.RecordSettingsConfig)serializer.ReadObject(fs);

                    // コントローラーGUIDとボタン割り当てが全て揃っているか
                    if (config.ControllerGuid == Guid.Empty) return false;
                    var btns = new[] { "LPButton", "MPButton", "HPButton", "LKButton", "MKButton", "HKButton" };
                    foreach (var btn in btns)
                    {
                        if ((config.ButtonIndices == null || !config.ButtonIndices.ContainsKey(btn) || config.ButtonIndices[btn] == null)
                            && (config.ZValues == null || !config.ZValues.ContainsKey(btn) || config.ZValues[btn] == null))
                            return false;
                    }
                    return true;
                }
            }
            catch (System.Runtime.Serialization.SerializationException)
            {
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // コントローラー接続処理
        private void ConnectRecordController()
        {
            // 既存のジョイスティックを解放
            currentJoystick?.Unacquire();
            currentJoystick?.Dispose();

            var configPath = Path.Combine(Application.StartupPath, "RecordSettingsConfig.json");
            if (!File.Exists(configPath)) return;

            using (var fs = new FileStream(configPath, FileMode.Open))
            {
                var serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(
                    typeof(RecordSettingsForm.RecordSettingsConfig),
                    new System.Runtime.Serialization.Json.DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true }
                );
                var config = (RecordSettingsForm.RecordSettingsConfig)serializer.ReadObject(fs);
                if (config.ControllerGuid == Guid.Empty) return;

                directInput = new SharpDX.DirectInput.DirectInput();
                var devices = directInput.GetDevices(SharpDX.DirectInput.DeviceType.Gamepad, SharpDX.DirectInput.DeviceEnumerationFlags.AllDevices)
                    .Concat(directInput.GetDevices(SharpDX.DirectInput.DeviceType.Joystick, SharpDX.DirectInput.DeviceEnumerationFlags.AllDevices))
                    .Concat(directInput.GetDevices(SharpDX.DirectInput.DeviceClass.GameControl, SharpDX.DirectInput.DeviceEnumerationFlags.AllDevices))
                    .ToList();

                var device = devices.FirstOrDefault(d => d.InstanceGuid == config.ControllerGuid);
                if (device == null) return;

                currentJoystick = new SharpDX.DirectInput.Joystick(directInput, device.InstanceGuid);
                currentJoystick.Acquire();
            }
        }

        // 記録スタートボタンの有効化
        private void UpdateRecButtonEnabled()
        {
            RecButton.Enabled = currentJoystick != null && IsRecordSettingsValid();
        }

        private async void RecButton_Click(object sender, EventArgs e)
        {
            if (!isRecording)
            {
                if (!IsRecordSettingsValid() || currentJoystick == null)
                {
                    MessageBox.Show("記録設定が不正です。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // recordConfig = LoadRecordSettingsConfig(); ← 削除
                recordedFrames.Clear();
                recordStartTime = DateTime.Now;
                int frameMs = 16;
                double.TryParse(FrameMsTextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double frameMsDouble);
                frameMs = (int)frameMsDouble;

                RecButton.Text = "記録中（停止する）";
                isRecording = true;

                await Task.Run(async () =>
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    long nextTick = sw.ElapsedMilliseconds;
                    while (isRecording)
                    {
                        var state = currentJoystick.GetCurrentState();
                        var frame = new Dictionary<string, string>();

                        // ボタン記録（従来通り）
                        var btnMap = recordConfig.ButtonIndices;
                        var zMap = recordConfig.ZValues;
                        var btnNames = new[] { "LPButton", "MPButton", "HPButton", "LKButton", "MKButton", "HKButton" };
                        for (int i = 0; i < btnNames.Length; i++)
                        {
                            string logical = btnNames[i].Replace("Button", "");
                            bool pressed = false;
                            if (btnMap != null && btnMap.ContainsKey(btnNames[i]) && btnMap[btnNames[i]] != null)
                            {
                                int idx = btnMap[btnNames[i]].Value;
                                if (idx >= 0 && idx < state.Buttons.Length)
                                    pressed = state.Buttons[idx];
                            }
                            else if (zMap != null && zMap.ContainsKey(btnNames[i]) && zMap[btnNames[i]] != null)
                            {
                                int z = zMap[btnNames[i]].Value;
                                if (state.Z == z)
                                    pressed = true;
                            }
                            if (pressed)
                                frame[logical] = "1";
                        }

                        // POV（方向キー）記録（斜め対応）
                        var povs = state.PointOfViewControllers;
                        if (povs != null && povs.Length > 0)
                        {
                            int pov = povs[0];
                            if (pov >= 0)
                            {
                                // POV値は36000で一周、9000単位で方向
                                // 斜め方向
                                if (pov == 4500) { frame["UP"] = "1"; frame["RIGHT"] = "1"; }
                                else if (pov == 13500) { frame["DOWN"] = "1"; frame["RIGHT"] = "1"; }
                                else if (pov == 22500) { frame["DOWN"] = "1"; frame["LEFT"] = "1"; }
                                else if (pov == 31500) { frame["UP"] = "1"; frame["LEFT"] = "1"; }
                                // 単方向
                                else if (pov == 0) frame["UP"] = "1";
                                else if (pov == 9000) frame["RIGHT"] = "1";
                                else if (pov == 18000) frame["DOWN"] = "1";
                                else if (pov == 27000) frame["LEFT"] = "1";
                            }
                        }
                        recordedFrames.Add(frame);

                        nextTick += frameMs;
                        long sleep = nextTick - sw.ElapsedMilliseconds;
                        if (sleep > 0)
                            await Task.Delay((int)sleep);
                        else
                            await Task.Yield();
                    }
                });
            }
            else
            {
                isRecording = false;
                RecButton.Text = "記録スタート";
                RemoveEmptyFrames(recordedFrames);
                ShowRecordSaveDialog();
            }
        }

        // 空フレーム除去
        private void RemoveEmptyFrames(List<Dictionary<string, string>> frames)
        {
            // 先頭
            while (frames.Count > 0 && frames[0].Count == 0)
                frames.RemoveAt(0);
            // 末尾
            while (frames.Count > 0 && frames[frames.Count - 1].Count == 0)
                frames.RemoveAt(frames.Count - 1);
        }

        // 記録後の保存ダイアログ
        private void ShowRecordSaveDialog()
        {
            var dialog = new Form
            {
                Text = "記録データの保存",
                Width = 400,
                Height = 200,
                StartPosition = FormStartPosition.CenterParent
            };

            var label = new Label
            {
                Text = "記録したデータをどのように保存しますか？",
                Dock = DockStyle.Top,
                Height = 40
            };
            dialog.Controls.Add(label);

            var overwriteButton = new Button
            {
                Text = "選択中マクロに上書き保存",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            overwriteButton.Click += (s, e) =>
            {
                SaveRecordedMacro(true);
                dialog.Close();
            };
            dialog.Controls.Add(overwriteButton);

            var saveAsButton = new Button
            {
                Text = "名前を付けて保存",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            saveAsButton.Click += (s, e) =>
            {
                SaveRecordedMacro(false);
                dialog.Close();
            };
            dialog.Controls.Add(saveAsButton);

            dialog.ShowDialog();
        }

        // マクロファイル保存処理
        private void SaveRecordedMacro(bool overwrite)
        {
            // 優先順位リスト
            var keyOrder = new[] { "UP", "DOWN", "LEFT", "RIGHT", "LP", "MP", "HP", "LK", "MK", "HK" };
            // 方向キーのキャメルケース変換用
            var directionCamel = new Dictionary<string, string>
            {
                { "UP", "Up" }, { "DOWN", "Down" }, { "LEFT", "Left" }, { "RIGHT", "Right" }
            };

            var sb = new StringBuilder();
            int waitCount = 0;
            int opCount = 0;
            List<string> lastOps = null;

            for (int i = 0; i < recordedFrames.Count; i++)
            {
                var frame = recordedFrames[i];
                if (frame.Count == 0)
                {
                    if (lastOps != null)
                    {
                        sb.AppendLine($"{opCount}: {string.Join(", ", lastOps)}");
                        lastOps = null;
                        opCount = 0;
                    }
                    waitCount++;
                }
                else
                {
                    // 優先順位で並べ替え＋方向キーはキャメルケース
                    var ops = frame.Keys
                        .Select(k => {
                            var upper = k.ToUpper();
                            return directionCamel.ContainsKey(upper) ? directionCamel[upper] : upper;
                        })
                        .Where(k => keyOrder.Contains(k.ToUpper()) || directionCamel.Values.Contains(k))
                        .OrderBy(k => Array.IndexOf(keyOrder, directionCamel.ContainsValue(k) ? keyOrder.First(x => directionCamel[x] == k) : k.ToUpper()))
                        .ToList();

                    if (waitCount > 0)
                    {
                        sb.AppendLine($"{waitCount}: ");
                        waitCount = 0;
                    }
                    if (lastOps != null && ops.SequenceEqual(lastOps))
                    {
                        opCount++;
                    }
                    else
                    {
                        if (lastOps != null)
                        {
                            sb.AppendLine($"{opCount}: {string.Join(", ", lastOps)}");
                        }
                        lastOps = ops;
                        opCount = 1;
                    }
                }
            }
            if (lastOps != null)
            {
                sb.AppendLine($"{opCount}: {string.Join(", ", lastOps)}");
            }
            if (waitCount > 0)
            {
                sb.AppendLine($"{waitCount}: ");
            }

            if (overwrite && MacroListBox.SelectedItems.Count > 0)
            {
                string macroName = MacroListBox.SelectedItems[MacroListBox.SelectedItems.Count - 1] as string;
                if (!string.IsNullOrEmpty(macroName))
                {
                    macroManager.SaveMacro(macroName, sb.ToString());
                    LoadMacroList();
                    MessageBox.Show("上書き保存しました。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = macroFolder;
                sfd.Filter = "CSVファイル (*.csv)|*.csv";
                sfd.FileName = $"新規記録_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string macroName = Path.GetFileNameWithoutExtension(sfd.FileName);
                    macroManager.SaveMacro(macroName, sb.ToString());
                    LoadMacroList();
                    MessageBox.Show("保存しました。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // --- 1. 設定ファイル読み込み関数を追加 ---
        private RecordSettingsForm.RecordSettingsConfig LoadRecordSettingsConfig()
        {
            var configPath = Path.Combine(Application.StartupPath, "RecordSettingsConfig.json");
            if (!File.Exists(configPath)) return null;
            using (var fs = new FileStream(configPath, FileMode.Open))
            {
                var serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(
                    typeof(RecordSettingsForm.RecordSettingsConfig),
                    new System.Runtime.Serialization.Json.DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true }
                );
                return (RecordSettingsForm.RecordSettingsConfig)serializer.ReadObject(fs);
            }
        }
    }
}