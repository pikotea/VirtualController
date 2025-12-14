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
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace VirtualController
{
    public partial class MainForm : Form
    {
        string macroFolder = Path.Combine(Application.StartupPath, "macros");
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

        // 追加フィールド: デバウンス用タイマーとロック
        private System.Timers.Timer macroWatcherDebounceTimer;
        private readonly object macroWatcherLock = new object();
        private int macroWatcherDebounceMs = 200;

        // 新: 現在の相対フォルダ（"" = ルート）
        private string currentRelativePath = "";

        // アイコンキャッシュ
        private Image folderIcon;
        // fileIcon は廃止（ファイルはアイコンなしで表示）

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
            // 起動時にマイグレーションを実行
            RunMigrationsIfNeeded();

            InitializeComponent();
            controllerService = new ControllerService();
            controllerService.Connect(); // ← 起動時に自動接続
            macroManager = new MacroManager(macroFolder);
            macroPlayer = new MacroPlayer(controllerService, macroFolder);
            StartMacroFolderWatcher();
            // 初期表示は currentRelativePath (LoadMacroSettings が起動時に上書きする場合がある)
            LoadMacroList();
        }

        private void RunMigrationsIfNeeded()
        {
            try
            {
                // 現在の実行アセンブリとロード済みアセンブリを検索して IMigration を実装する型を見つける
                var types = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a =>
                    {
                        try { return a.GetTypes(); }
                        catch (ReflectionTypeLoadException ex) { return ex.Types.Where(t => t != null); }
                        catch { return Array.Empty<Type>(); }
                    });

                var migrations = types
                    .Where(t => typeof(IMigration).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract)
                    .Select(t => (IMigration)Activator.CreateInstance(t))
                    .OrderBy(m => m.Id)
                    .ToList();

                if (migrations.Count > 0)
                {
                    MigrationRunner.RunMigrations(migrations);
                }
            }
            catch (Exception ex)
            {
                // 任意: ログ出力やユーザー通知を行う（ここでは例外を無視して起動継続）
                System.Diagnostics.Trace.TraceWarning("マイグレーション実行中に例外: " + ex.Message);
            }
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


        // --- アイコン読み込みヘルパー ---
        private void EnsureIconsLoaded()
        {
            if (folderIcon != null) return;

            var info = new SHSTOCKICONINFO();
            info.cbSize = (uint)Marshal.SizeOf(typeof(SHSTOCKICONINFO));
            const uint SIID_FOLDER = 3;
            const uint SHGSI_ICON = 0x000000100;
            const uint SHGSI_SMALLICON = 0x000000001;

            int hr = SHGetStockIconInfo(SIID_FOLDER, SHGSI_ICON | SHGSI_SMALLICON, ref info);
            if (hr == 0 && info.hIcon != IntPtr.Zero)
            {
                try
                {
                    using (var ico = Icon.FromHandle(info.hIcon))
                    {
                        folderIcon = ico.ToBitmap();
                    }
                }
                finally
                {
                    DestroyIcon(info.hIcon);
                }
            }

            // フォールバック: 取得できなければ組み込みの SystemIcons を使う
            if (folderIcon == null)
            {
                folderIcon = SystemIcons.WinLogo.ToBitmap();
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHSTOCKICONINFO
        {
            public uint cbSize;
            public IntPtr hIcon;
            public int iSysImageIndex;
            public int iIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szPath;
        }

        [DllImport("shell32.dll", SetLastError = true)]
        private static extern int SHGetStockIconInfo(uint siid, uint uFlags, ref SHSTOCKICONINFO psii);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        // マクロ一覧のリロード（相対フォルダ対応）
        private void LoadMacroList(bool stopMacro = true, string relativePath = null)
        {
            if (relativePath == null) relativePath = currentRelativePath ?? "";
            // 現在の相対パスを更新
            currentRelativePath = relativePath ?? "";

            // マクロ一覧更新時は自動停止（条件付き）
            if (stopMacro)
                StopMacroIfPlaying();

            // 現在の選択状態を保存（名前ベース）
            var selectedNames = MacroListBox.SelectedItems.Cast<object>()
                .Select(o => o?.ToString() ?? "")
                .ToList();

            MacroListBox.Items.Clear();

            // 取得: フォルダ/ファイルエントリ
            var entries = macroManager.GetMacroEntries(currentRelativePath);

            // ルートでなければ「…」を先頭に追加（親へ戻る）
            if (!string.IsNullOrEmpty(currentRelativePath))
            {
                var parentEntry = new MacroManager.MacroEntry { Name = "…", RelativePath = "..", IsFolder = true };
                MacroListBox.Items.Add(parentEntry);
            }

            // ディレクトリ（先）とファイル（後）は MacroManager が順序を返す
            foreach (var entry in entries)
            {
                MacroListBox.Items.Add(entry);
            }

            // 修正: 常に複数選択を許可（サブフォルダ内でも複数選択できるようにする）
            MacroListBox.SelectionMode = SelectionMode.MultiExtended;

            // 選択状態を復元（名前が一致するアイテムを再選択）
            MacroListBox.ClearSelected();
            foreach (var name in selectedNames)
            {
                for (int i = 0; i < MacroListBox.Items.Count; i++)
                {
                    if (MacroListBox.Items[i].ToString().Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        MacroListBox.SetSelected(i, true);
                        break;
                    }
                }
            }

            // pendingSelectMacroName があれば優先選択（保存直後の非同期更新対策）
            if (!string.IsNullOrEmpty(pendingSelectMacroName))
            {
                for (int i = 0; i < MacroListBox.Items.Count; i++)
                {
                    if (MacroListBox.Items[i].ToString().Equals(pendingSelectMacroName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        MacroListBox.ClearSelected();
                        MacroListBox.SetSelected(i, true);
                        break;
                    }
                }
                pendingSelectMacroName = null;
            }
        }

        // DrawItem ハンドラ（Designer で接続済み）
        private void MacroListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();
            EnsureIconsLoaded();

            var item = MacroListBox.Items[e.Index];
            string text = item?.ToString() ?? "";
            bool isFolder = false;

            if (item is MacroManager.MacroEntry me)
                isFolder = me.IsFolder;
            else if (!string.IsNullOrEmpty(text) && text.StartsWith("…"))
                isFolder = true;

            var icon = isFolder ? folderIcon : null;
            Rectangle iconRect = new Rectangle(e.Bounds.Left + 2, e.Bounds.Top + 2, 16, 16);

            if (icon != null)
            {
                e.Graphics.DrawImage(icon, iconRect);
            }

            Rectangle textRect = new Rectangle(e.Bounds.Left + 22, e.Bounds.Top + 2, e.Bounds.Width - 24, e.Bounds.Height - 4);
            var fore = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? SystemColors.HighlightText : e.ForeColor;
            TextRenderer.DrawText(e.Graphics, text, e.Font, textRect, fore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);

            e.DrawFocusRectangle();
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
            var macroNames = GetSelectedMacroNames(includeRelativePath: true);

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
                        token,
                        currentJoystick,
                        recordConfig
                    );
                }
                catch (InvalidDataException ex)
                {
                    // マクロのパースエラーをユーザーに表示
                    try
                    {
                        this.Invoke((Action)(() =>
                        {
                            MessageBox.Show(this, ex.Message, "マクロのパースエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                    catch
                    {
                        // フォームが既に破棄されている可能性があるので無視
                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        this.Invoke((Action)(() =>
                        {
                            MessageBox.Show(this, "再生中にエラーが発生しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                    catch
                    {
                    }
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
            // ルートフォルダを監視。フィルタは "*.*" にしてフォルダイベントも取得する
            macroWatcher = new FileSystemWatcher(macroFolder)
            {
                Filter = "*.*",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite,
                IncludeSubdirectories = true
            };

            // バッファを増やして大量イベント時のオーバーフローを緩和
            try
            {
                macroWatcher.InternalBufferSize = 64 * 1024; // 64KB
            }
            catch
            {
                // 一部環境で例外が出る可能性があるので無視して続行
            }

            // 既存イベントは保持するが、実処理はデバウンスタイマへ委譲する
            macroWatcher.Created += MacroFolderChanged;
            macroWatcher.Deleted += MacroFolderChanged;
            macroWatcher.Renamed += MacroFolderChanged;

            // デバウンスタイマー初期化（AutoReset=false で手動再起動）
            macroWatcherDebounceTimer = new System.Timers.Timer(macroWatcherDebounceMs) { AutoReset = false };
            macroWatcherDebounceTimer.Elapsed += (s, e) =>
            {
                // タイマーから来るスレッドは ThreadPool。ここで UI スレッドへディスパッチする。
                try
                {
                    this.Invoke((Action)(() =>
                    {
                        System.Diagnostics.Debug.WriteLine("[MacroPlayer] watcher debounce elapsed - reloading macro list");
                        try
                        {
                            LoadMacroList(false);
                        }
                        catch { }
                    }));
                }
                catch (ObjectDisposedException) { /* フォーム破棄時の安全処理 */ }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[MacroPlayer] debounce handler exception: " + ex);
                }
            };

            macroWatcher.EnableRaisingEvents = true;
        }

        // ファイル変更時のイベントハンドラ（簡素化してデバウンスに集約）
        private void MacroFolderChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                // デバッグログ（発生を把握する）
                // 除外: 変更（Changed）はファイルの上書き等で大量発生するため再読み込み対象外とする
                if (e.ChangeType == WatcherChangeTypes.Changed)
                {
                    System.Diagnostics.Debug.WriteLine($"[MacroPlayer] Ignored Changed event: {e.FullPath}");
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"[MacroPlayer] MacroFolderChanged: {e.ChangeType} {e.FullPath}");

                // デバウンス: イベントが来るたびにタイマーを再起動してまとめる
                lock (macroWatcherLock)
                {
                    // 再起動（既に動いていれば停止してからスタート）
                    macroWatcherDebounceTimer.Stop();
                    macroWatcherDebounceTimer.Start();
                }

                // マクロ再生中なら停止要求だけは即時行う（安全のため）
                if (isMacroPlaying)
                {
                    this.Invoke((Action)(() => StopMacroIfPlaying()));
                    // LoadMacroList はデバウンス後に行う
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[MacroPlayer] MacroFolderChanged exception: " + ex);
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


        // マクロ選択時のロード（フォルダ遷移対応）
        // --- マクロ選択時にラベル表示・保存 ---
        private void MacroListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // クリックでフォルダに入る仕様: 最初に選んだアイテムがフォルダなら遷移
            if (MacroListBox.SelectedItems.Count > 0)
            {
                var first = MacroListBox.SelectedItems[0];
                if (first is MacroManager.MacroEntry me && me.IsFolder)
                {
                    // 追加: フォルダ移動時はマクロを停止し、選択状態をクリアする
                    StopMacroIfPlaying();
                    MacroListBox.ClearSelected();

                    // parent entry ("..")
                    if (me.RelativePath == "..")
                    {
                        // 親へ戻る
                        if (string.IsNullOrEmpty(currentRelativePath))
                        {
                            // already root -> nothing
                        }
                        else
                        {
                            var parent = Path.GetDirectoryName(currentRelativePath);
                            // Path.GetDirectoryName returns null for single-level; normalize to empty string
                            currentRelativePath = parent ?? "";
                            LoadMacroList(false, currentRelativePath);
                            // 重要: フォルダ移動直後に設定を保存して起動時の復元を最新にする
                            SaveMacroSettings();
                        }
                        return;
                    }
                    else
                    {
                        // サブフォルダへ移動（relative path stored in entry)
                        currentRelativePath = me.RelativePath ?? "";
                        LoadMacroList(false, currentRelativePath);
                        // 重要: フォルダ移動直後に設定を保存して起動時の復元を最新にする
                        SaveMacroSettings();
                        return;
                    }
                }
            }

            // 通常のファイル選択処理（複数選択対応）
            StopMacroIfPlaying();

            loadedMacro.Clear();
            string lastMacroName = null;
            foreach (var macroNameObj in MacroListBox.SelectedItems)
            {
                string macroName = null;
                if (macroNameObj is MacroManager.MacroEntry me2 && !me2.IsFolder)
                    macroName = me2.Name;
                else if (macroNameObj is string s)
                    macroName = s;

                if (string.IsNullOrEmpty(macroName)) continue;

                // サブフォルダ対応のロード
                var macroFrames = MacroPlayer.MacroFrame.LoadMacroFile(Path.Combine(macroFolder, string.IsNullOrEmpty(currentRelativePath) ? macroName + ".csv" : Path.Combine(currentRelativePath, macroName + ".csv")));
                if (macroFrames != null && macroFrames.Count > 0)
                {
                    loadedMacro.AddRange(macroFrames);
                }
                lastMacroName = macroName;
            }

            // 編集エリアに内容表示
            if (lastMacroName != null)
            {
                // LoadMacroText は相対フォルダ対応を使う
                MacroEditTextBox.Text = macroManager.LoadMacroText(currentRelativePath, lastMacroName);
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

        // 保存ボタン押下時の処理（相対フォルダ対応）
        private void SaveMacroButton_Click(object sender, EventArgs e)
        {
            string macroName = null;
            if (MacroListBox.SelectedItems.Count > 0)
            {
                var sel = MacroListBox.SelectedItems[MacroListBox.SelectedItems.Count - 1];
                macroName = sel is MacroManager.MacroEntry me && !me.IsFolder ? me.Name : sel?.ToString();
            }
            if (string.IsNullOrEmpty(macroName))
            {
                // 新規マクロ名生成（現在フォルダ内の重複チェック）
                int idx = 1;
                List<string> existing = macroManager.GetMacroNames(currentRelativePath);
                do
                {
                    macroName = $"新規マクロ{idx}";
                    idx++;
                } while (existing.Contains(macroName));
            }
            using (var sfd = new SaveFileDialog())
            {
                string initialDir = string.IsNullOrEmpty(currentRelativePath) ? macroFolder : Path.Combine(macroFolder, currentRelativePath);
                sfd.InitialDirectory = initialDir;
                sfd.FileName = macroName + ".csv";
                sfd.Filter = "CSVファイル (*.csv)|*.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var savedName = Path.GetFileNameWithoutExtension(sfd.FileName);
                    // 相対フォルダに保存
                    macroManager.SaveMacroInFolder(currentRelativePath, savedName, MacroEditTextBox.Text);


                    // マクロは保存済み。FileSystemWatcher のデバウンスで再読み込みされるのでここでは明示呼び出ししない
                    pendingSelectMacroName = savedName;
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

        // --- 上書き保存処理（相対フォルダ対応） ---
        private void OverwriteMacro()
        {
            if (!string.IsNullOrEmpty(editingMacroName))
            {
                try
                {
                    pendingSelectMacroName = editingMacroName;  // 上書き保存後に優先選択するマクロ名を設定

                    macroManager.SaveMacroInFolder(currentRelativePath, editingMacroName, MacroEditTextBox.Text);
                    lastLoadedMacroText = MacroEditTextBox.Text;
                    loadedMacro = MacroPlayer.MacroFrame.LoadMacroFile(Path.Combine(macroFolder, string.IsNullOrEmpty(currentRelativePath) ? editingMacroName + ".csv" : Path.Combine(currentRelativePath, editingMacroName + ".csv")));
                    OverwriteSaveButton.Enabled = false;
                    this.Invoke((Action)(() =>
                    {
                        MacroEditTextBox.SelectionLength = 0;
                        MacroEditTextBox.SelectionStart = MacroEditTextBox.TextLength;
                    }));
                }
                finally
                {
                }
            }
        }

        // 置き換え対象: 現在の SaveMacroSettings(...)
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
        $"SelectedMacros={selectedMacroStr}",
        $"CurrentFolder={currentRelativePath}"
    };
            File.WriteAllLines(Path.Combine(Application.StartupPath, SettingsFile), lines, Encoding.UTF8);
        }

        // --- 設定ロード: 複数選択復元（CurrentFolder 復元含む） ---
        private void LoadMacroSettings()
        {
            string path = Path.Combine(Application.StartupPath, SettingsFile);
            if (!File.Exists(path)) return;
            string[] selectedMacros = null;
            string loadedCurrentFolder = null;
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
                    case "CurrentFolder":
                        loadedCurrentFolder = val;
                        break;
                }
            }

            // currentRelativePath を復元（存在確認）
            if (!string.IsNullOrEmpty(loadedCurrentFolder))
            {
                string full = Path.Combine(macroFolder, loadedCurrentFolder);
                if (Directory.Exists(full))
                {
                    currentRelativePath = loadedCurrentFolder;
                }
                else
                {
                    currentRelativePath = "";
                }
            }

            // マクロ一覧を currentRelativePath に合わせて再構築
            LoadMacroList(false, currentRelativePath);

            // マクロ一覧ロード後に複数選択復元
            if (selectedMacros != null && selectedMacros.Length > 0)
            {
                MacroListBox.ClearSelected();
                foreach (var macro in selectedMacros)
                {
                    for (int i = 0; i < MacroListBox.Items.Count; i++)
                    {
                        if (MacroListBox.Items[i].ToString().Equals(macro, StringComparison.CurrentCultureIgnoreCase))
                        {
                            MacroListBox.SetSelected(i, true);
                            break;
                        }
                    }
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
        // LT（左トリガー）
        private void LTButton_MouseDown(object sender, MouseEventArgs e)
        {
            // 0〜255。押下で最大値にする
            controllerService.Controller.SetSliderValue(Xbox360Slider.LeftTrigger, byte.MaxValue);
        }
        private void LTButton_MouseUp(object sender, MouseEventArgs e)
        {
            // 離したら0に戻す
            controllerService.Controller.SetSliderValue(Xbox360Slider.LeftTrigger, 0);
        }

        // RT（右トリガー）
        private void RTButton_MouseDown(object sender, MouseEventArgs e)
        {
            controllerService.Controller.SetSliderValue(Xbox360Slider.RightTrigger, byte.MaxValue);
        }
        private void RTButton_MouseUp(object sender, MouseEventArgs e)
        {
            controllerService.Controller.SetSliderValue(Xbox360Slider.RightTrigger, 0);
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
            RecCautionLabel.Visible = !(currentJoystick != null && IsRecordSettingsValid());
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
                        .Select(k =>
                        {
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
                string macroName = GetSelectedMacroNames().LastOrDefault();
                if (!string.IsNullOrEmpty(macroName))
                {
                    // 相対フォルダ対応で保存する（currentRelativePath を使う）
                    macroManager.SaveMacroInFolder(currentRelativePath, macroName, sb.ToString());
                    // FileSystemWatcher will trigger reload — remember selection
                    pendingSelectMacroName = macroName;
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
                    // Let watcher reload and restore selection
                    pendingSelectMacroName = macroName;
                 }
             }
         }

        // --- 設定ファイル読み込み関数を追加 ---
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

        // 追加: 選択中のマクロ名を安全に取得するヘルパー
        private List<string> GetSelectedMacroNames(bool includeRelativePath = false)
        {
            var names = new List<string>();
            foreach (var item in MacroListBox.SelectedItems)
            {
                if (item is MacroManager.MacroEntry me)
                {
                    if (me.IsFolder) continue;
                    if (string.IsNullOrEmpty(me.Name)) continue;
                    if (includeRelativePath && !string.IsNullOrEmpty(currentRelativePath))
                        names.Add(Path.Combine(currentRelativePath, me.Name));
                    else
                        names.Add(me.Name);
                }
                else
                {
                    var s = item?.ToString();
                    if (string.IsNullOrEmpty(s)) continue;
                    if (includeRelativePath && !string.IsNullOrEmpty(currentRelativePath))
                        names.Add(Path.Combine(currentRelativePath, s));
                    else
                        names.Add(s);
                }
            }
            return names;
        }

    }
}