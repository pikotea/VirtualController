using Nefarius.ViGEm.Client;
using Nefarius.ViGEm.Client.Targets;
using Nefarius.ViGEm.Client.Targets.Xbox360;
using System.Collections.Generic;
using System.Threading;

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
        }

        public void Connect()
        {
            // 実装例
            if (controller == null)
            {
                // コントローラの初期化処理
                controller = client.CreateXbox360Controller();
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

        public void SetInputs(Dictionary<Xbox360Axis, short> axisValues, Dictionary<Xbox360Button, bool> buttonStates, int duration)
        {
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
            Thread.Sleep(duration);
            if (axisValues != null)
            {
                foreach (var kvp in axisValues)
                {
                    controller.SetAxisValue(kvp.Key, 0);
                }
            }
            if (buttonStates != null)
            {
                foreach (var kvp in buttonStates)
                {
                    controller.SetButtonState(kvp.Key, false);
                }
            }
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
        }
    }
}