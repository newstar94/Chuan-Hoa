using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ChuanHoa.Client.Core.Lexicon
{
    public sealed class LocalFeedbackEntry
    {
        public string Original { get; set; } = string.Empty;
        public string Suggestion { get; set; } = string.Empty;
        public int AcceptedCount { get; set; }
        public int RejectedCount { get; set; }
        public DateTime LastUpdatedUtc { get; set; }
    }

    /// <summary>
    /// Stores local user feedback (accepted vs rejected suggestions) strictly on-device.
    /// Used by candidate ranker to adjust suggestion weights according to user preferences.
    /// </summary>
    public sealed class LocalFeedbackStore
    {
        private static readonly object SyncLock = new object();
        private readonly string _filePath;
        private readonly Dictionary<string, LocalFeedbackEntry> _entries =
            new Dictionary<string, LocalFeedbackEntry>(StringComparer.OrdinalIgnoreCase);

        public LocalFeedbackStore(string storageDirectory)
        {
            if (storageDirectory == null) throw new ArgumentNullException(nameof(storageDirectory));
            _filePath = Path.Combine(storageDirectory, "spelling_feedback.tsv");
            Load();
        }

        public void RecordFeedback(string original, string suggestion, bool accepted)
        {
            if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(suggestion)) return;

            var key = MakeKey(original, suggestion);
            lock (SyncLock)
            {
                if (!_entries.TryGetValue(key, out var entry))
                {
                    entry = new LocalFeedbackEntry
                    {
                        Original = original.Trim(),
                        Suggestion = suggestion.Trim()
                    };
                    _entries[key] = entry;
                }

                if (accepted) entry.AcceptedCount++;
                else entry.RejectedCount++;
                entry.LastUpdatedUtc = DateTime.UtcNow;

                Save();
            }
        }

        public double GetPreferenceAdjustment(string original, string suggestion)
        {
            if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(suggestion)) return 0.0;

            var key = MakeKey(original, suggestion);
            lock (SyncLock)
            {
                if (_entries.TryGetValue(key, out var entry))
                {
                    var total = entry.AcceptedCount + entry.RejectedCount;
                    if (total >= 2)
                    {
                        var acceptanceRate = (double)entry.AcceptedCount / total;
                        // Range: -0.15 (if heavily rejected) to +0.15 (if heavily accepted)
                        return (acceptanceRate - 0.5) * 0.3;
                    }
                }
            }
            return 0.0;
        }

        private static string MakeKey(string original, string suggestion) =>
            $"{original.Trim().ToLowerInvariant()}==>{suggestion.Trim().ToLowerInvariant()}";

        private void Load()
        {
            try
            {
                if (File.Exists(_filePath))
                {
                    var lines = File.ReadAllLines(_filePath, Encoding.UTF8);
                    lock (SyncLock)
                    {
                        _entries.Clear();
                        foreach (var line in lines)
                        {
                            var parts = line.Split('\t');
                            if (parts.Length >= 4 &&
                                int.TryParse(parts[2], out var accepted) &&
                                int.TryParse(parts[3], out var rejected))
                            {
                                var orig = parts[0];
                                var sugg = parts[1];
                                var entry = new LocalFeedbackEntry
                                {
                                    Original = orig,
                                    Suggestion = sugg,
                                    AcceptedCount = accepted,
                                    RejectedCount = rejected
                                };
                                _entries[MakeKey(orig, sugg)] = entry;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Fallback gracefully
            }
        }

        private void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(_filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var sb = new StringBuilder();
                foreach (var pair in _entries.Values)
                {
                    sb.AppendLine($"{pair.Original}\t{pair.Suggestion}\t{pair.AcceptedCount}\t{pair.RejectedCount}");
                }
                File.WriteAllText(_filePath, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // Fallback gracefully
            }
        }
    }
}
