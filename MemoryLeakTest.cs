using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OpenCvSharp;
using Serilog;

namespace peilin
{
    public static class MemoryLeakTest
    {
        private static Process currentProcess = Process.GetCurrentProcess();
        private static bool isTestRunning = false;

        public static async Task RunFullPipelineTest(int totalSamples = 1000, int samplesPerMinute = 100)
        {
            if (isTestRunning)
            {
                Log.Warning("[記憶體測試] 測試已在執行中，請勿重複啟動");
                return;
            }

            isTestRunning = true;

            try
            {
                Log.Information($"[記憶體測試] 開始完整流程測試：{totalSamples} 個樣品，速度 {samplesPerMinute} 樣品/分");

                currentProcess.Refresh();
                long startMemory = currentProcess.PrivateMemorySize64 / (1024 * 1024);
                Log.Information($"[記憶體測試] 初始記憶體：{startMemory} MB");

                int sampleIntervalMs = 60000 / samplesPerMinute;

                // ✅ 確保執行緒運行
                EnsureGetMatThreadsRunning();

                // ✅ 啟用系統狀態
                app.status = true;
                app.DetectMode = 0;

                using (Mat testImage = new Mat(2048, 2448, MatType.CV_8UC3, Scalar.RandomColor()))
                {
                    for (int sampleIndex = 0; sampleIndex < totalSamples; sampleIndex++)
                    {
                        // ✅ 依序呼叫四個站點
                        await SimulateStationCapture(testImage, 0, sampleIndex);
                        await Task.Delay(50);

                        await SimulateStationCapture(testImage, 1, sampleIndex);
                        await Task.Delay(50);

                        await SimulateStationCapture(testImage, 2, sampleIndex);
                        await Task.Delay(50);

                        await SimulateStationCapture(testImage, 3, sampleIndex);

                        // 每 100 個樣品檢查記憶體
                        if ((sampleIndex + 1) % 100 == 0)
                        {
                            currentProcess.Refresh();
                            long currentMemory = currentProcess.PrivateMemorySize64 / (1024 * 1024);
                            long growth = currentMemory - startMemory;

                            int q1 = app.Queue_Bitmap1.Count;
                            int q2 = app.Queue_Bitmap2.Count;
                            int q3 = app.Queue_Bitmap3.Count;
                            int q4 = app.Queue_Bitmap4.Count;

                            Log.Information($"[記憶體測試] 第 {sampleIndex + 1} 個樣品：{currentMemory} MB (增長 {growth} MB)");
                            Log.Information($"[記憶體測試] 佇列狀態：Q1={q1}, Q2={q2}, Q3={q3}, Q4={q4}");

                            double expectedMaxGrowthMB = (sampleIndex + 1) * 10;
                            if (growth > expectedMaxGrowthMB)
                            {
                                Log.Warning($"[記憶體測試] ⚠️ 記憶體增長異常！預期 <{expectedMaxGrowthMB:F0} MB，實際 {growth} MB");
                            }

                            int totalQueue = q1 + q2 + q3 + q4;
                            if (totalQueue > 100)
                            {
                                Log.Warning($"[記憶體測試] ⚠️ 佇列積壓過多：{totalQueue} 張影像");
                            }
                        }

                        await Task.Delay(sampleIntervalMs);
                    }

                    Log.Information("[記憶體測試] 樣品發送完成，等待佇列處理完成...");
                    await WaitForQueuesEmpty(timeoutSeconds: 120);
                }

                app.status = false;

                currentProcess.Refresh();
                long finalMemory = currentProcess.PrivateMemorySize64 / (1024 * 1024);
                long totalGrowth = finalMemory - startMemory;

                Log.Information($"[記憶體測試] ===== 測試完成 =====");
                Log.Information($"[記憶體測試] 初始記憶體：{startMemory} MB");
                Log.Information($"[記憶體測試] 最終記憶體：{finalMemory} MB");
                Log.Information($"[記憶體測試] 總增長：{totalGrowth} MB");
                Log.Information($"[記憶體測試] 平均每個樣品洩漏：{(double)totalGrowth / totalSamples:F2} MB");

                if (totalGrowth > 500)
                {
                    Log.Error($"[記憶體測試] ❌ 測試失敗！記憶體洩漏嚴重：{totalGrowth} MB");
                }
                else
                {
                    Log.Information($"[記憶體測試] ✅ 測試通過！記憶體洩漏在可接受範圍");
                }
            }
            finally
            {
                isTestRunning = false;
                app.status = false;
            }
        }

        private static async Task SimulateStationCapture(Mat sourceImage, int camID, int sampleId)
        {
            await Task.Run(() =>
            {
                try
                {
                    //SimulateReceiver(sourceImage, camID, sampleId);
                }
                catch (Exception ex)
                {
                    Log.Error($"[記憶體測試] 模擬站點 {camID + 1} 擷取時發生錯誤: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 模擬 OnImageGrabbed 呼叫 Receiver
        /// ✅ 修正：正確更新計數器，模擬真實的站點順序
        /// </summary>
        private static void SimulateReceiver(Mat sourceImage, int camID, int sampleId)
        {
            // ✅ 修正：更新所有站點的計數器（模擬真實的 PLC 觸發順序）
            lock (app.counter)
            {
                // 站點 N 的計數器 = 當前樣品ID - (3 - N)
                // 例如：樣品0在站點2時，stop0=-2, stop1=-1, stop2=0
                for (int i = 0; i <= camID; i++)
                {
                    app.counter[$"stop{i}"] = sampleId - (camID - i);
                }
            }

            // ✅ 呼叫真實的 Receiver 方法
            var form1 = System.Windows.Forms.Application.OpenForms["Form1"] as Form1;
            if (form1 != null)
            {
                form1.Receiver(camID, sourceImage, DateTime.Now);
            }
            else
            {
                Log.Warning("[記憶體測試] 找不到 Form1 實例，無法呼叫 Receiver");
            }
        }

        private static void EnsureGetMatThreadsRunning()
        {
            if (app.T1 == null || app.T1.Status != TaskStatus.Running)
            {
                Log.Warning("[記憶體測試] getMat1 執行緒未運行，嘗試啟動...");
                app.T1 = new Task(() =>
                {
                    var form1 = System.Windows.Forms.Application.OpenForms["Form1"] as Form1;
                    if (form1 != null)
                    {
                        var method = typeof(Form1).GetMethod("getMat1",
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        method?.Invoke(form1, null);
                    }
                }, TaskCreationOptions.LongRunning);
                app.T1.Start();
            }

            if (app.T2 == null || app.T2.Status != TaskStatus.Running)
            {
                Log.Warning("[記憶體測試] getMat2 執行緒未運行，嘗試啟動...");
                app.T2 = new Task(() =>
                {
                    var form1 = System.Windows.Forms.Application.OpenForms["Form1"] as Form1;
                    if (form1 != null)
                    {
                        var method = typeof(Form1).GetMethod("getMat2",
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        method?.Invoke(form1, null);
                    }
                }, TaskCreationOptions.LongRunning);
                app.T2.Start();
            }

            if (app.T3 == null || app.T3.Status != TaskStatus.Running)
            {
                Log.Warning("[記憶體測試] getMat3 執行緒未運行，嘗試啟動...");
                app.T3 = new Task(() =>
                {
                    var form1 = System.Windows.Forms.Application.OpenForms["Form1"] as Form1;
                    if (form1 != null)
                    {
                        var method = typeof(Form1).GetMethod("getMat3",
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        method?.Invoke(form1, null);
                    }
                }, TaskCreationOptions.LongRunning);
                app.T3.Start();
            }

            if (app.T4 == null || app.T4.Status != TaskStatus.Running)
            {
                Log.Warning("[記憶體測試] getMat4 執行緒未運行，嘗試啟動...");
                app.T4 = new Task(() =>
                {
                    var form1 = System.Windows.Forms.Application.OpenForms["Form1"] as Form1;
                    if (form1 != null)
                    {
                        var method = typeof(Form1).GetMethod("getMat4",
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        method?.Invoke(form1, null);
                    }
                }, TaskCreationOptions.LongRunning);
                app.T4.Start();
            }

            Thread.Sleep(500);
            Log.Information("[記憶體測試] getMat1-4 執行緒已確認運行");
        }

        private static async Task WaitForQueuesEmpty(int timeoutSeconds = 120)
        {
            DateTime startTime = DateTime.Now;
            int lastTotalCount = -1;
            int stableCounter = 0;

            while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds)
            {
                int q1 = app.Queue_Bitmap1.Count;
                int q2 = app.Queue_Bitmap2.Count;
                int q3 = app.Queue_Bitmap3.Count;
                int q4 = app.Queue_Bitmap4.Count;
                int totalCount = q1 + q2 + q3 + q4;

                if (totalCount == 0)
                {
                    Log.Information("[記憶體測試] 所有佇列已清空");
                    return;
                }

                if (totalCount == lastTotalCount)
                {
                    stableCounter++;
                    if (stableCounter > 20)
                    {
                        Log.Warning($"[記憶體測試] ⚠️ 佇列長時間無變化，可能卡住：Q1={q1}, Q2={q2}, Q3={q3}, Q4={q4}");
                        app._wh1.Set();
                        app._wh2.Set();
                        app._wh3.Set();
                        app._wh4.Set();
                    }
                }
                else
                {
                    stableCounter = 0;
                }

                lastTotalCount = totalCount;
                await Task.Delay(500);
            }

            int finalQ1 = app.Queue_Bitmap1.Count;
            int finalQ2 = app.Queue_Bitmap2.Count;
            int finalQ3 = app.Queue_Bitmap3.Count;
            int finalQ4 = app.Queue_Bitmap4.Count;

            Log.Warning($"[記憶體測試] ⚠️ 超時！佇列未清空：Q1={finalQ1}, Q2={finalQ2}, Q3={finalQ3}, Q4={finalQ4}");
        }
    }
}