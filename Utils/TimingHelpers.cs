using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace VirtualController
{
    internal static class TimingHelpers
    {
        [DllImport("winmm.dll", SetLastError = true)]
        private static extern uint timeBeginPeriod(uint ms);

        [DllImport("winmm.dll", SetLastError = true)]
        private static extern uint timeEndPeriod(uint ms);

        private static int refCount = 0;
        private const uint PERIOD_MS = 1;

        // アプリ起動時に1回呼ぶのが理想だが、ここでは再生直前に参照カウントで管理する
        public static void BeginHighResolution()
        {
            if (Interlocked.Increment(ref refCount) == 1)
            {
                try { timeBeginPeriod(PERIOD_MS); } catch { /* 無視 */ }
            }
        }

        public static void EndHighResolution()
        {
            if (Interlocked.Decrement(ref refCount) == 0)
            {
                try { timeEndPeriod(PERIOD_MS); } catch { /* 無視 */ }
            }
        }

        // ハイブリッド待機。targetMs は Stopwatch.Elapsed.TotalMilliseconds ベースの目標時刻
        public static void WaitUntil(Stopwatch sw, double targetMs, CancellationToken token)
        {
            const double SPIN_THRESHOLD_MS = 2.5;   // スピンに切り替える閾値（実測で調整）
            const double SLEEP_MARGIN_MS = 8.0;     // Sleep からスピンへ切替えるマージン（目標: 最大遅延 <= 8ms）

            while (!token.IsCancellationRequested)
            {
                double remaining = targetMs - sw.Elapsed.TotalMilliseconds;
                if (remaining <= 0) break;

                if (remaining > (SLEEP_MARGIN_MS + SPIN_THRESHOLD_MS))
                {
                    int sleepMs = Math.Max(1, (int)(remaining - SLEEP_MARGIN_MS));
                    Thread.Sleep(sleepMs);
                }
                else if (remaining > SPIN_THRESHOLD_MS)
                {
                    Thread.Sleep(1);
                }
                else
                {
                    // 最終局面：短い busy-wait で正確に待つ
                    while (sw.Elapsed.TotalMilliseconds < targetMs)
                    {
                        Thread.SpinWait(10);
                        if (token.IsCancellationRequested) return;
                    }
                    break;
                }
            }
        }
    }
}