using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Lexicon
{
    public sealed class CandidateItem
    {
        public CandidateItem(string text, double baseScore, string source)
        {
            Text = text ?? string.Empty;
            BaseScore = baseScore;
            Source = source;
        }

        public string Text { get; }
        public double BaseScore { get; set; }
        public string Source { get; }

        public override string ToString() => $"{Text} ({BaseScore:F2}, {Source})";
    }

    /// <summary>
    /// Generates top candidates (up to 8) using Levenshtein distance, confusion sets,
    /// and telex typing patterns. Always retains the original word as candidate #0.
    /// </summary>
    public sealed class FastCandidateGenerator
    {
        private static readonly CultureInfo VietnameseCulture = CultureInfo.GetCultureInfo("vi-VN");
        private readonly VietnameseLexiconSpellChecker? _lexicon;

        public FastCandidateGenerator(VietnameseLexiconSpellChecker? lexicon = null)
        {
            _lexicon = lexicon;
        }

        public IReadOnlyList<CandidateItem> GenerateCandidates(string word, int maxCandidates = 8)
        {
            if (string.IsNullOrWhiteSpace(word))
                return Array.Empty<CandidateItem>();

            var normalizedOriginal = word.Normalize(NormalizationForm.FormC);
            var results = new List<CandidateItem>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Candidate 0 is always the original word (allows the model to choose NO_CHANGE)
            results.Add(new CandidateItem(normalizedOriginal, 0.5, "Original"));
            seen.Add(normalizedOriginal);

            // 1. Check direct deterministic lexicon correction if available
            if (_lexicon != null)
            {
                var deterministic = _lexicon.FindDeterministicCorrection(normalizedOriginal);
                if (deterministic != null && seen.Add(deterministic))
                {
                    results.Add(new CandidateItem(deterministic, 0.95, "LexiconDeterministic"));
                }
            }

            // 2. Check phonetic / confusion sets
            foreach (var alternate in VietnameseConfusionSets.GeneratePhoneticAlternates(normalizedOriginal))
            {
                if (results.Count >= maxCandidates) break;
                if (seen.Add(alternate))
                {
                    var isKnown = _lexicon?.IsKnown(alternate) ?? false;
                    var score = isKnown ? 0.85 : 0.65;
                    results.Add(new CandidateItem(alternate, score, "ConfusionSet"));
                }
            }

            // 3. Check administrative confusion pairs
            if (VietnameseConfusionSets.AdministrativeConfusionPairs.TryGetValue(normalizedOriginal, out var adminCandidate))
            {
                if (results.Count < maxCandidates && seen.Add(adminCandidate))
                {
                    results.Add(new CandidateItem(adminCandidate, 0.90, "AdministrativePair"));
                }
            }

            // 4. Check common telex artifacts (e.g. ngỉ -> nghỉ, xẩy -> xảy)
            var telexCandidate = ResolveSimpleTelexArtifact(normalizedOriginal);
            if (telexCandidate != null && results.Count < maxCandidates && seen.Add(telexCandidate))
            {
                results.Add(new CandidateItem(telexCandidate, 0.88, "TelexResolution"));
            }

            // Order candidates descending by base score, but keep original as candidate if desired
            return results.Take(maxCandidates).ToList();
        }

        private static string? ResolveSimpleTelexArtifact(string word)
        {
            // Frequent single-letter typo patterns in Vietnamese typing
            if (word.Equals("ngỉ", StringComparison.OrdinalIgnoreCase)) return ApplyCase(word, "nghỉ");
            if (word.Equals("ngĩ", StringComparison.OrdinalIgnoreCase)) return ApplyCase(word, "nghĩ");
            if (word.Equals("ngiêng", StringComparison.OrdinalIgnoreCase)) return ApplyCase(word, "nghiêng");
            if (word.Equals("ngi ngờ", StringComparison.OrdinalIgnoreCase)) return ApplyCase(word, "nghi ngờ");
            if (word.Equals("ngành ngề", StringComparison.OrdinalIgnoreCase)) return ApplyCase(word, "ngành nghề");
            return null;
        }

        private static string ApplyCase(string source, string target)
        {
            if (source.All(c => !char.IsLetter(c) || char.IsUpper(c)))
                return target.ToUpper(VietnameseCulture);
            if (source.Length > 0 && char.IsUpper(source[0]))
                return char.ToUpper(target[0], VietnameseCulture) + target.Substring(1);
            return target;
        }
    }
}
