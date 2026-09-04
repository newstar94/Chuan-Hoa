using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading.Tasks;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests
{
    public sealed class VietnameseEngineEndToEndIpcTests
    {
        [Fact]
        public async Task EngineProcess_StartsAndServesIpcRequests()
        {
            var exePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "tools", "VietnameseEngine", "bin", "Debug", "net10.0", "VietnameseEngine.exe");
            exePath = Path.GetFullPath(exePath);

            if (!File.Exists(exePath))
            {
                // If not built yet or different configuration, skip
                return;
            }

            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                CreateNoWindow = true,
                UseShellExecute = false,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            Process? engineProcess = null;
            try
            {
                engineProcess = Process.Start(psi);
                Assert.NotNull(engineProcess);

                var client = new VietnameseEngineIpcClient(exePath);

                // 1. Wait for engine to become alive and ping
                var isAlive = false;
                for (var attempt = 0; attempt < 20; attempt++)
                {
                    client.ResetCircuitBreaker();
                    isAlive = await client.PingAsync(timeoutMs: 300);
                    if (isAlive) break;
                    await Task.Delay(100);
                }

                Assert.True(isAlive, "VietnameseEngine should respond to ping over Named Pipe.");

                // 2. Send text containing real-world error "bàn dao"
                const string sampleText = "Đồng chí gửi bàn dao hồ sơ trước ngày mai.";
                var findings = await client.CheckAsync(sampleText, paragraphIndex: 1, timeoutMs: 2000);

                Assert.NotEmpty(findings);
                var banDaoFinding = Assert.Single(findings, f => f.Original == "bàn dao");
                Assert.Equal(3, banDaoFinding.Level);
                Assert.Equal("bàn giao", banDaoFinding.BestSuggestion);
                Assert.Equal(sampleText.IndexOf("bàn dao", StringComparison.Ordinal), banDaoFinding.StartOffset);
                Assert.Equal("bàn dao".Length, banDaoFinding.Length);

                // 3. Stop the engine gracefully via pipe
                try
                {
                    using (var pipe = new NamedPipeClientStream(".", "ChuanHoa_VietnameseEngine", PipeDirection.InOut))
                    {
                        await pipe.ConnectAsync(1000);
                        using (var writer = new StreamWriter(pipe, Encoding.UTF8) { AutoFlush = true })
                        {
                            await writer.WriteLineAsync("{\"action\":\"stop\"}");
                        }
                    }
                }
                catch { }

                if (!engineProcess.WaitForExit(3000))
                {
                    engineProcess.Kill();
                }
            }
            finally
            {
                if (engineProcess != null && !engineProcess.HasExited)
                {
                    try { engineProcess.Kill(); } catch { }
                }
            }
        }
    }
}
