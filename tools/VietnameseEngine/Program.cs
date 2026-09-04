using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ChuanHoa.Client.Core.Lexicon;

namespace ChuanHoa.VietnameseEngine
{
    public static class Program
    {
        private const string PipeName = "ChuanHoa_VietnameseEngine";
        private static readonly FastCandidateGenerator CandidateGenerator = new FastCandidateGenerator();
        private static readonly PersonalDictionaryManager DictionaryManager = PersonalDictionaryManager.Instance;
        private static readonly CancellationTokenSource ShutdownCts = new CancellationTokenSource();
        private static readonly TimeSpan IdleTimeout = TimeSpan.FromMinutes(15);
        private static DateTime _lastActivityUtc = DateTime.UtcNow;

        public static async Task<int> Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine($"[VietnameseEngine] Starting background server on \\\\.\\pipe\\{PipeName}...");

            // Start idle watchdog
            _ = Task.Run(async () =>
            {
                while (!ShutdownCts.Token.IsCancellationRequested)
                {
                    await Task.Delay(TimeSpan.FromSeconds(30), ShutdownCts.Token).ConfigureAwait(false);
                    if (DateTime.UtcNow - _lastActivityUtc > IdleTimeout)
                    {
                        Console.WriteLine("[VietnameseEngine] Idle timeout reached. Shutting down gracefully.");
                        ShutdownCts.Cancel();
                        break;
                    }
                }
            }, ShutdownCts.Token);

            try
            {
                while (!ShutdownCts.Token.IsCancellationRequested)
                {
                    var pipeServer = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.InOut,
                        NamedPipeServerStream.MaxAllowedServerInstances,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    try
                    {
                        await pipeServer.WaitForConnectionAsync(ShutdownCts.Token).ConfigureAwait(false);
                        _lastActivityUtc = DateTime.UtcNow;
                        _ = Task.Run(() => HandleClientAsync(pipeServer, ShutdownCts.Token));
                    }
                    catch (OperationCanceledException)
                    {
                        pipeServer.Dispose();
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[VietnameseEngine] Pipe connection error: {ex.Message}");
                        pipeServer.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[VietnameseEngine] Fatal error: {ex.Message}");
                return 1;
            }

            Console.WriteLine("[VietnameseEngine] Stopped.");
            return 0;
        }

        private static async Task HandleClientAsync(NamedPipeServerStream stream, CancellationToken cancellationToken)
        {
            using (stream)
            using (var reader = new StreamReader(stream, Encoding.UTF8, false, 4096, leaveOpen: true))
            using (var writer = new StreamWriter(stream, Encoding.UTF8, 4096, leaveOpen: true) { AutoFlush = true })
            {
                while (stream.IsConnected && !cancellationToken.IsCancellationRequested)
                {
                    string? line;
                    try
                    {
                        line = await reader.ReadLineAsync().ConfigureAwait(false);
                    }
                    catch
                    {
                        break;
                    }

                    if (string.IsNullOrEmpty(line)) break;
                    _lastActivityUtc = DateTime.UtcNow;

                    EngineResponse response;
                    try
                    {
                        var request = JsonSerializer.Deserialize<EngineRequest>(line, new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true
                        });

                        if (request == null)
                        {
                            response = new EngineResponse { IsSuccess = false, ErrorMessage = "Invalid JSON payload" };
                        }
                        else if (string.Equals(request.Action, "stop", StringComparison.OrdinalIgnoreCase))
                        {
                            response = new EngineResponse { RequestId = request.RequestId, IsSuccess = true };
                            var jsonStop = JsonSerializer.Serialize(response);
                            await writer.WriteLineAsync(jsonStop).ConfigureAwait(false);
                            ShutdownCts.Cancel();
                            break;
                        }
                        else if (string.Equals(request.Action, "ping", StringComparison.OrdinalIgnoreCase))
                        {
                            response = new EngineResponse { RequestId = request.RequestId, IsSuccess = true };
                        }
                        else
                        {
                            response = ProcessCheckRequest(request);
                        }
                    }
                    catch (Exception ex)
                    {
                        response = new EngineResponse
                        {
                            IsSuccess = false,
                            ErrorMessage = ex.Message
                        };
                    }

                    var responseJson = JsonSerializer.Serialize(response);
                    try
                    {
                        await writer.WriteLineAsync(responseJson).ConfigureAwait(false);
                    }
                    catch
                    {
                        break;
                    }
                }
            }
        }

        private static EngineResponse ProcessCheckRequest(EngineRequest request)
        {
            var response = new EngineResponse
            {
                RequestId = request.RequestId,
                IsSuccess = true
            };

            if (string.IsNullOrWhiteSpace(request.Text))
                return response;

            var tokens = VietnameseWordTokenizer.TokenizeParagraph(request.Text, request.ParagraphIndex);

            // 1. Check administrative multi-word confusion pairs first
            foreach (var pair in VietnameseConfusionSets.AdministrativeConfusionPairs)
            {
                var idx = request.Text.IndexOf(pair.Key, StringComparison.OrdinalIgnoreCase);
                while (idx >= 0)
                {
                    response.Findings.Add(new EngineFinding
                    {
                        StartOffset = idx,
                        Length = pair.Key.Length,
                        Original = pair.Key,
                        Level = 3, // Real-word contextual error
                        Confidence = 0.95,
                        IssueDescription = $"Cụm từ “{pair.Key}” có thể sai ngữ cảnh trong văn bản hành chính.",
                        Suggestions = new List<EngineSuggestion>
                        {
                            new EngineSuggestion { Text = pair.Value, Score = 0.95, Source = "AdministrativePair" }
                        }
                    });

                    idx = request.Text.IndexOf(pair.Key, idx + pair.Key.Length, StringComparison.OrdinalIgnoreCase);
                }
            }

            // 2. Check individual tokens for telex / phonetic typos
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token.Kind != VietnameseTokenKind.Word) continue;

                // Skip if token is in user's personal dictionary or ignored
                if (DictionaryManager.IsKnownOrIgnored(token.Text, request.DocumentHash))
                    continue;

                // Generate candidates
                var candidates = CandidateGenerator.GenerateCandidates(token.Text, 4);
                if (candidates.Count > 1)
                {
                    var bestAlternative = candidates[1];
                    // If candidate score is high and not equal to original
                    if (bestAlternative.BaseScore >= 0.85 &&
                        !string.Equals(bestAlternative.Text, token.Text, StringComparison.OrdinalIgnoreCase))
                    {
                        var finding = new EngineFinding
                        {
                            StartOffset = token.StartOffset,
                            Length = token.Length,
                            Original = token.Text,
                            Level = 2,
                            Confidence = bestAlternative.BaseScore,
                            IssueDescription = $"Từ “{token.Text}” có thể viết chưa chính xác.",
                            Suggestions = new List<EngineSuggestion>()
                        };

                        foreach (var cand in candidates)
                        {
                            finding.Suggestions.Add(new EngineSuggestion
                            {
                                Text = cand.Text,
                                Score = cand.BaseScore,
                                Source = cand.Source
                            });
                        }

                        response.Findings.Add(finding);
                    }
                }
            }

            return response;
        }
    }
}
