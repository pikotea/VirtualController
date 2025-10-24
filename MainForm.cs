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
        public MainForm()
        {
            InitializeComponent();
            controllerService = new ControllerService();
            macroManager = new MacroManager(macroFolder);
            macroPlayer = new MacroPlayer(controllerService, macroFolder); // 追加
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
                        token
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

        // 接続
        private void ConnectButton_Click(object sender, EventArgs e)
        {
            controllerService.Connect();
            this.SetControllerButtonsEnabled(true);
        }

        // 切断
        private void DisconnectButton_Click(object sender, EventArgs e)
        {
            controllerService.Disconnect();
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


        // マクロ選択時のロード（複数選択対応・編集エリア表示）
        // --- マクロ選択時にラベル表示・保存 ---
        private void MacroListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            StopMacroIfPlaying();

            loadedMacro.Clear();
            string lastMacroName = null;
            bool macroLoaded = false;
            foreach (var macroNameObj in MacroListBox.SelectedItems)
            {
                string macroName = macroNameObj as string;
                if (string.IsNullOrEmpty(macroName)) continue;
                // ここで自作MacroFrameのLoadMacroFileを使う
                var macroFrames = MacroPlayer.MacroFrame.LoadMacroFile(Path.Combine(macroFolder, macroName + ".csv"));
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
            // 再生ボタンの有効化条件を変更
            PlayMacroButton.Enabled = macroLoaded && controllerService.IsConnected && DisconnectButton.Enabled;

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
                } while (macroManager.GetMacroNames().Contains(macroName)); // 変更
            }
            using (var sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = macroFolder;
                sfd.FileName = macroName + ".csv";
                sfd.Filter = "CSVファイル (*.csv)|*.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    macroManager.SaveMacro(Path.GetFileNameWithoutExtension(sfd.FileName), MacroEditTextBox.Text); // 変更
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
                loadedMacro = MacroPlayer.MacroFrame.LoadMacroFile(path);
                LoadMacroList(false);
                OverwriteSaveButton.Enabled = false;
                this.Invoke((Action)(() =>
                {
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
    }
}
