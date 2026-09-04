using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ChuanHoa.Client.Core.Scanning
{
    public sealed class EngineFindingResult
    {
        public int StartOffset { get; set; }
        public int Length { get; set; }
        public string Original { get; set; } = string.Empty;
        public int Level { get; set; }
        public double Confidence { get; set; }
        public string Description { get; set; } = string.Empty;
        public string BestSuggestion { get; set; } = string.Empty;
    }

    /// <summary>
    /// Out-of-process client connecting to VietnameseEngine.exe over local Named Pipe.
    /// Operates asynchronously, completely offline, with strict timeouts to prevent freezing Word.
    /// </summary>
    public sealed class VietnameseEngineIpcClient
    {
        private const string PipeName = "ChuanHoa_VietnameseEngine";
        private readonly string? _engineExecutablePath;
        private int _consecutiveFailures;
        private const int MaxConsecutiveFailures = 3;

        public VietnameseEngineIpcClient(string? engineExecutablePath = null)
        {
            _engineExecutablePath = engineExecutablePath;
        }

        public bool IsCircuitBroken => _consecutiveFailures >= MaxConsecutiveFailures;

        public bool Ping(int timeoutMs = 500)
        {
            if (IsCircuitBroken || !IsEngineProcessRunning()) return false;

            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut))
                {
                    client.Connect(timeoutMs);
                    using (var reader = new StreamReader(client, Encoding.UTF8))
                    using (var writer = new StreamWriter(client, Encoding.UTF8) { AutoFlush = true })
                    {
                        var requestId = Guid.NewGuid().ToString("N");
                        writer.WriteLine("{\"action\":\"ping\",\"requestId\":\"" + requestId + "\"}");
                        writer.Flush();
                        var line = reader.ReadLine();
                        var ok = !string.IsNullOrEmpty(line) && line.Contains(requestId);
                        if (ok) _consecutiveFailures = 0;
                        return ok;
                    }
                }
            }
            catch
            {
                _consecutiveFailures++;
                return false;
            }
        }

        public Task<bool> PingAsync(int timeoutMs = 500) => Task.Run(() => Ping(timeoutMs));

        public void EnsureEngineRunning()
        {
            if (IsEngineProcessRunning()) return;
            Task.Run(() => EnsureProcessStarted());
        }

        public IReadOnlyList<EngineFindingResult> Check(
            string text,
            int paragraphIndex = 0,
            string documentHash = "",
            int timeoutMs = 250)
        {
            if (string.IsNullOrWhiteSpace(text) || IsCircuitBroken || !IsEngineProcessRunning())
                return Array.Empty<EngineFindingResult>();

            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut))
                {
                    client.Connect(timeoutMs);
                    using (var reader = new StreamReader(client, Encoding.UTF8))
                    using (var writer = new StreamWriter(client, Encoding.UTF8) { AutoFlush = true })
                    {
                        var requestId = Guid.NewGuid().ToString("N");
                        var jsonRequest = "{\"action\":\"check\",\"requestId\":\"" + requestId +
                                          "\",\"paragraphIndex\":" + paragraphIndex +
                                          ",\"documentHash\":\"" + JsonEscape(documentHash) +
                                          "\",\"text\":\"" + JsonEscape(text) + "\"}";

                        writer.WriteLine(jsonRequest);
                        writer.Flush();
                        var responseLine = reader.ReadLine();

                        if (string.IsNullOrEmpty(responseLine))
                            return Array.Empty<EngineFindingResult>();

                        _consecutiveFailures = 0;
                        return ParseFindings(responseLine);
                    }
                }
            }
            catch
            {
                _consecutiveFailures++;
                return Array.Empty<EngineFindingResult>();
            }
        }

        public Task<IReadOnlyList<EngineFindingResult>> CheckAsync(
            string text,
            int paragraphIndex = 0,
            string documentHash = "",
            int timeoutMs = 2000)
        {
            return Task.Run(() => Check(text, paragraphIndex, documentHash, timeoutMs));
        }

        public void ResetCircuitBreaker()
        {
            _consecutiveFailures = 0;
        }

        private string? ResolveEngineExecutablePath()
        {
            if (!string.IsNullOrEmpty(_engineExecutablePath) && File.Exists(_engineExecutablePath))
                return _engineExecutablePath;

            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrEmpty(baseDir))
                {
                    var p1 = Path.Combine(baseDir, "VietnameseEngine.exe");
                    if (File.Exists(p1)) return p1;
                    var p2 = Path.Combine(baseDir, "Engine", "VietnameseEngine.exe");
                    if (File.Exists(p2)) return p2;
                }
            }
            catch { }

            try
            {
                var localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var pInstalledEngine = Path.Combine(localApp, "ChuanHoa", "DevelopmentInstaller", "Current", "Engine", "VietnameseEngine.exe");
                if (File.Exists(pInstalledEngine)) return pInstalledEngine;

                var pInstalledRoot = Path.Combine(localApp, "ChuanHoa", "DevelopmentInstaller", "Current", "VietnameseEngine.exe");
                if (File.Exists(pInstalledRoot)) return pInstalledRoot;
            }
            catch { }

            return null;
        }

        private void EnsureProcessStarted()
        {
            var exePath = ResolveEngineExecutablePath();
            if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath))
                return;

            try
            {
                var processName = Path.GetFileNameWithoutExtension(exePath);
                var existing = Process.GetProcessesByName(processName);
                if (existing.Length == 0)
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = exePath,
                        WorkingDirectory = Path.GetDirectoryName(exePath) ?? string.Empty,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        WindowStyle = ProcessWindowStyle.Hidden
                    };
                    Process.Start(psi);
                }
            }
            catch
            {
                // Fallback gracefully without throwing
            }
        }

        private static bool IsEngineProcessRunning()
        {
            try
            {
                return Process.GetProcessesByName("VietnameseEngine").Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static string JsonEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }

        private static IReadOnlyList<EngineFindingResult> ParseFindings(string json)
        {
            var results = new List<EngineFindingResult>();
            if (string.IsNullOrEmpty(json)) return results;

            // Lightweight parsing without depending on external JSON packages in netstandard2.0
            var findingsIdx = json.IndexOf("\"findings\":", StringComparison.OrdinalIgnoreCase);
            if (findingsIdx < 0) return results;

            var arrayStart = json.IndexOf('[', findingsIdx);
            if (arrayStart < 0) return results;

            var arrayEnd = json.LastIndexOf(']');
            if (arrayEnd <= arrayStart) return results;

            var findingsContent = json.Substring(arrayStart + 1, arrayEnd - arrayStart - 1);
            var items = findingsContent.Split(new[] { "},{" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var item in items)
            {
                var start = ExtractInt(item, "startOffset");
                var len = ExtractInt(item, "length");
                var level = ExtractInt(item, "level");
                var original = ExtractString(item, "original");
                var desc = ExtractString(item, "issueDescription");
                var suggestion = ExtractFirstSuggestion(item);

                if (len > 0)
                {
                    results.Add(new EngineFindingResult
                    {
                        StartOffset = start,
                        Length = len,
                        Level = level,
                        Original = original,
                        Description = desc,
                        BestSuggestion = suggestion
                    });
                }
            }

            return results;
        }

        private static int ExtractInt(string json, string propertyName)
        {
            var key = "\"" + propertyName + "\":";
            var idx = json.IndexOf(key, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return 0;
            var numStart = idx + key.Length;
            var numEnd = numStart;
            while (numEnd < json.Length && (char.IsDigit(json[numEnd]) || json[numEnd] == '-')) numEnd++;
            var str = json.Substring(numStart, numEnd - numStart).Trim();
            int.TryParse(str, out var val);
            return val;
        }

        private static string ExtractString(string json, string propertyName)
        {
            var key = "\"" + propertyName + "\":\"";
            var idx = json.IndexOf(key, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return string.Empty;
            var strStart = idx + key.Length;
            var strEnd = json.IndexOf('"', strStart);
            if (strEnd < 0) return string.Empty;
            return json.Substring(strStart, strEnd - strStart);
        }

        private static string ExtractFirstSuggestion(string json)
        {
            var key = "\"text\":\"";
            var idx = json.IndexOf(key, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return string.Empty;
            var strStart = idx + key.Length;
            var strEnd = json.IndexOf('"', strStart);
            if (strEnd < 0) return string.Empty;
            return json.Substring(strStart, strEnd - strStart);
        }
    }
}
