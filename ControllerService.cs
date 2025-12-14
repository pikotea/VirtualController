using Nefarius.ViGEm.Client;
using Nefarius.ViGEm.Client.Targets;
using Nefarius.ViGEm.Client.Targets.Xbox360;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;

namespace VirtualController
{
    public class ControllerService
    {
        private readonly ViGEmClient client;
        private IXbox360Controller controller;

        public IXbox360Controller Controller => controller;

        // 追加: IsConnected プロパティ
        public bool IsConnected => controller != null;

        public ControllerService()
        {
            client = new ViGEmClient();
            controller = client.CreateXbox360Controller();
            // 自動送信を無効にして明示的に SubmitReport() を呼ぶ方式にする
            controller.AutoSubmitReport = false;
        }

        public void Connect()
        {
            // 実装例
            if (controller == null)
            {
                // コントローラの初期化処理
                controller = client.CreateXbox360Controller();
                controller.AutoSubmitReport = false;
            }
            controller.Connect(); // ←必ずConnect()を呼ぶ
        }

        public void Disconnect()
        {
            // 実装例
            if (controller != null)
            {
                // コントローラの切断処理
                controller.Disconnect();
                controller = null;
            }
        }

        public void SetInputs(Dictionary<Xbox360Axis, short> axisValues, Dictionary<Xbox360Button, bool> buttonStates)
        {
            Debug.WriteLine("SetInputs called");
            if (axisValues != null)
            {
                foreach (var kvp in axisValues)
                {
                    controller.SetAxisValue(kvp.Key, kvp.Value);
                }
            }
            if (buttonStates != null)
            {
                foreach (var kvp in buttonStates)
                {
                    controller.SetButtonState(kvp.Key, kvp.Value);
                }
            }
            // 変更: 一連の更新後にまとめてレポート送信
            controller.SubmitReport();
        }

        public void AllOff()
        {
            controller.SetAxisValue(Xbox360Axis.LeftThumbY, 0);
            controller.SetAxisValue(Xbox360Axis.LeftThumbX, 0);
            controller.SetButtonState(Xbox360Button.A, false);
            controller.SetButtonState(Xbox360Button.B, false);
            controller.SetButtonState(Xbox360Button.X, false);
            controller.SetButtonState(Xbox360Button.Y, false);
            controller.SetButtonState(Xbox360Button.LeftShoulder, false);
            controller.SetButtonState(Xbox360Button.RightShoulder, false);
            controller.SetSliderValue(Xbox360Slider.LeftTrigger, 0);
            controller.SetSliderValue(Xbox360Slider.RightTrigger, 0);
            controller.SubmitReport();
        }
    }
}