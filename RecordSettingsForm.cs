using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpDX.DirectInput;

namespace VirtualController
{
    public partial class RecordSettingsForm : Form
    {
        private DirectInput directInput;
        private List<DeviceInstance> uniqueDevices;
        private Joystick currentJoystick;

        private bool isListening = false;
        private Button listeningButton = null;
        private Label listeningLabel = null;
        private int? assignedZValue = null;

        // ボタンごとの割り当て値保存用プロパティ
        private Dictionary<Button, int?> assignedButtonIndices = new Dictionary<Button, int?>();
        private Dictionary<Button, int?> assignedZValues = new Dictionary<Button, int?>();

        private const string ConfigFilePath = "RecordSettingsConfig.json";

        // 変更後: private class RecordSettingsConfig
        public class RecordSettingsConfig
        {
            public Guid ControllerGuid;
            public Dictionary<string, int?> ButtonIndices = new Dictionary<string, int?>();
            public Dictionary<string, int?> ZValues = new Dictionary<string, int?>();
        }

        public RecordSettingsForm()
        {
            InitializeComponent();
            SetActionButtonsEnabled(false);
            OKButton.Enabled = false;
            LoadPhysicalControllers();
            ControllerComboBox.SelectedIndexChanged += ControllerComboBox_SelectedIndexChanged;
            OKButton.Click += OKButton_Click;
            LoadConfig(); // 追加: 設定ファイルから復元
        }

        private void SetActionButtonsEnabled(bool enabled)
        {
            LPButton.Enabled = enabled;
            MPButton.Enabled = enabled;
            HPButton.Enabled = enabled;
            LKButton.Enabled = enabled;
            MKButton.Enabled = enabled;
            HKButton.Enabled = enabled;
        }

        private void LoadPhysicalControllers()
        {
            directInput = new DirectInput();
            var devices = directInput.GetDevices(DeviceType.Gamepad, DeviceEnumerationFlags.AllDevices)
                .Concat(directInput.GetDevices(DeviceType.Joystick, DeviceEnumerationFlags.AllDevices))
                .Concat(directInput.GetDevices(DeviceClass.GameControl, DeviceEnumerationFlags.AllDevices))
                .ToList();

            // InstanceGuidで重複排除
            uniqueDevices = devices
                .GroupBy(d => d.InstanceGuid)
                .Select(g => g.First())
                .ToList();

            ControllerComboBox.Items.Clear();
            foreach (var device in uniqueDevices)
            {
                ControllerComboBox.Items.Add($"{device.InstanceName} ({device.InstanceGuid})");
            }
        }

        private void ControllerComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ControllerComboBox.SelectedIndex < 0)
            {
                SetActionButtonsEnabled(false);
                return;
            }

            SetActionButtonsEnabled(true);

            var device = uniqueDevices[ControllerComboBox.SelectedIndex];

            // 既存のジョイスティックを解放
            currentJoystick?.Unacquire();
            currentJoystick?.Dispose();

            // 新しいジョイスティックを接続
            currentJoystick = new Joystick(directInput, device.InstanceGuid);
            currentJoystick.Acquire();

            LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] コントローラー「{device.InstanceName}」に接続しました。\r\n");

            var timer = new Timer();
            timer.Interval = 100;
            timer.Tick += (s, ev) =>
            {
                try
                {
                    var state = currentJoystick.GetCurrentState();

                    if (isListening)
                    {
                        // ボタン判定
                        int pressedIndex = Array.FindIndex(state.Buttons, b => b);
                        if (pressedIndex >= 0)
                        {
                            EndListening($"{pressedIndex}ボタン", pressedIndex, null);
                            LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {pressedIndex}ボタン を割り当てました\r\n");
                            return;
                        }
                        // Z軸判定
                        int zValueListening = state.Z;
                        if (assignedZValue == null || assignedZValue != zValueListening)
                        {
                            if (zValueListening == 128)
                            {
                                EndListening("Z1ボタン", null, zValueListening);
                                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] Z1ボタン を割り当てました\r\n");
                            }
                            else if (zValueListening == 65408)
                            {
                                EndListening("Z2ボタン", null, zValueListening);
                                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] Z2ボタン を割り当てました\r\n");
                            }
                            return;
                        }
                    }

                    // Z軸の値を取得し、変化があればログ出力
                    int zValue = state.Z;
                    if (zValue != 32767)
                    {
                        string zLog = $"[{DateTime.Now:HH:mm:ss}] Z軸: {zValue}\r\n";
                        if (zValue == 128)
                            zLog = $"[{DateTime.Now:HH:mm:ss}] Z1ボタン\r\n";
                        else if (zValue == 65408)
                            zLog = $"[{DateTime.Now:HH:mm:ss}] Z2ボタン\r\n";
                        LogTextBox.AppendText(zLog);
                    }

                    // ボタン・方向キーのログ
                    var pressed = state.Buttons.Select((b, i) => b ? $"{i}ボタン" : null).Where(x => x != null).ToList();

                    var povs = state.PointOfViewControllers;
                    var directions = new List<string>();
                    foreach (var pov in povs)
                    {
                        if (pov >= 0)
                        {
                            switch (pov)
                            {
                                case 0: directions.Add("上"); break;
                                case 4500: directions.Add("右上"); break;
                                case 9000: directions.Add("右"); break;
                                case 13500: directions.Add("右下"); break;
                                case 18000: directions.Add("下"); break;
                                case 22500: directions.Add("左下"); break;
                                case 27000: directions.Add("左"); break;
                                case 31500: directions.Add("左上"); break;
                                default: directions.Add($"POV:{pov}"); break;
                            }
                        }
                    }

                    if (pressed.Count > 0 || directions.Count > 0)
                    {
                        LogTextBox.AppendText(
                            $"[{DateTime.Now:HH:mm:ss}] {string.Join(", ", pressed.Concat(directions))}\r\n"
                        );
                    }
                }
                catch { /* デバイス切断時など */ }
            };
            timer.Start();
        }

        // リスニングモード開始
        private void StartListening(Button button, Label label)
        {
            if (isListening) return;
            isListening = true;
            listeningButton = button;
            listeningLabel = label;
            button.BackColor = Color.Orange;
            label.Text = "リスニング中...";
            LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {button.Text} リスニングモード開始\r\n");
        }

        // リスニングモード終了
        private void EndListening(string displayText, int? buttonIndex, int? zValue)
        {
            // 既存の割り当て解除処理
            if (buttonIndex != null)
            {
                // 他のボタンに同じbuttonIndexが割り当てられていたら解除
                foreach (var kvp in assignedButtonIndices.ToList())
                {
                    if (kvp.Value == buttonIndex && kvp.Key != listeningButton)
                    {
                        assignedButtonIndices[kvp.Key] = null;
                        assignedZValues[kvp.Key] = null;
                        // ラベルも未割り当て表示に
                        var label = Controls.OfType<Label>().FirstOrDefault(l => l.Name == kvp.Key.Name.Replace("Button", "Label"));
                        if (label != null) label.Text = "未割り当て";
                    }
                }
            }
            if (zValue != null)
            {
                // 他のボタンに同じzValueが割り当てられていたら解除
                foreach (var kvp in assignedZValues.ToList())
                {
                    if (kvp.Value == zValue && kvp.Key != listeningButton)
                    {
                        assignedButtonIndices[kvp.Key] = null;
                        assignedZValues[kvp.Key] = null;
                        var label = Controls.OfType<Label>().FirstOrDefault(l => l.Name == kvp.Key.Name.Replace("Button", "Label"));
                        if (label != null) label.Text = "未割り当て";
                    }
                }
            }

            if (listeningButton != null)
            {
                listeningButton.BackColor = SystemColors.Control;
                listeningLabel.Text = displayText;
                assignedButtonIndices[listeningButton] = buttonIndex;
                assignedZValues[listeningButton] = zValue;
            }
            isListening = false;
            listeningButton = null;
            listeningLabel = null;

            // 6つのボタンがすべて割り当て済みならOKボタンを有効化
            OKButton.Enabled = IsAllActionButtonsAssigned();
        }

        private bool IsAllActionButtonsAssigned()
        {
            var buttons = new[] { LPButton, MPButton, HPButton, LKButton, MKButton, HKButton };
            return buttons.All(btn =>
                (assignedButtonIndices.ContainsKey(btn) && assignedButtonIndices[btn] != null) ||
                (assignedZValues.ContainsKey(btn) && assignedZValues[btn] != null)
            );
        }

        // 各ボタンのクリックイベント
        private void LPButton_Click(object sender, EventArgs e) => StartListening(LPButton, LPLabel);
        private void MPButton_Click(object sender, EventArgs e) => StartListening(MPButton, MPLabel);
        private void HPButton_Click(object sender, EventArgs e) => StartListening(HPButton, HPLabel);
        private void LKButton_Click(object sender, EventArgs e) => StartListening(LKButton, LKLabel);
        private void MKButton_Click(object sender, EventArgs e) => StartListening(MKButton, MKLabel);
        private void HKButton_Click(object sender, EventArgs e) => StartListening(HKButton, HKLabel);

        private void OKButton_Click(object sender, EventArgs e)
        {
            SaveConfig();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void SaveConfig()
        {
            var config = new RecordSettingsConfig();
            if (ControllerComboBox.SelectedIndex >= 0)
                config.ControllerGuid = uniqueDevices[ControllerComboBox.SelectedIndex].InstanceGuid;

            var buttons = new[] { LPButton, MPButton, HPButton, LKButton, MKButton, HKButton };
            foreach (var btn in buttons)
            {
                config.ButtonIndices[btn.Name] = (assignedButtonIndices.ContainsKey(btn) && assignedButtonIndices[btn] != null)
                    ? assignedButtonIndices[btn]
                    : null;
                config.ZValues[btn.Name] = (assignedZValues.ContainsKey(btn) && assignedZValues[btn] != null)
                    ? assignedZValues[btn]
                    : null;
            }

            using (var fs = new FileStream(ConfigFilePath, FileMode.Create))
            {
                var serializer = new DataContractJsonSerializer(
                    typeof(RecordSettingsConfig),
                    new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true }
                );
                serializer.WriteObject(fs, config);
            }
        }

        private void LoadConfig()
        {
            if (!File.Exists(ConfigFilePath)) return;

            RecordSettingsConfig config;
            using (var fs = new FileStream(ConfigFilePath, FileMode.Open))
            {
                var serializer = new DataContractJsonSerializer(typeof(RecordSettingsConfig));
                config = (RecordSettingsConfig)serializer.ReadObject(fs);
            }

            // コントローラー自動選択
            int idx = uniqueDevices.FindIndex(d => d.InstanceGuid == config.ControllerGuid);
            if (idx >= 0)
            {
                ControllerComboBox.SelectedIndex = idx;
                SetActionButtonsEnabled(true);
            }

            // 割り当て復元
            var buttons = new[] { LPButton, MPButton, HPButton, LKButton, MKButton, HKButton };
            foreach (var btn in buttons)
            {
                assignedButtonIndices[btn] = config.ButtonIndices.ContainsKey(btn.Name) ? config.ButtonIndices[btn.Name] : null;
                assignedZValues[btn] = config.ZValues.ContainsKey(btn.Name) ? config.ZValues[btn.Name] : null;
                var label = Controls.OfType<Label>().FirstOrDefault(l => l.Name == btn.Name.Replace("Button", "Label"));
                if (label != null)
                {
                    if (assignedButtonIndices[btn] != null)
                        label.Text = $"{assignedButtonIndices[btn]}ボタン";
                    else if (assignedZValues[btn] != null)
                        label.Text = assignedZValues[btn] == 128 ? "Z1ボタン" : assignedZValues[btn] == 65408 ? "Z2ボタン" : $"Z軸:{assignedZValues[btn]}";
                    else
                        label.Text = "未割り当て";
                }
            }
            OKButton.Enabled = IsAllActionButtonsAssigned();
        }
    }
}
