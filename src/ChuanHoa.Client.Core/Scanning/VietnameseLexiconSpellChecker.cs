using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Small, deterministic Vietnamese word checker backed by the signed rule pack.
    /// It deliberately checks individual Vietnamese syllables; contextual real-word
    /// mistakes remain the responsibility of the correction phrase list.
    /// </summary>
    public sealed class VietnameseLexiconSpellChecker
    {
        private static readonly CultureInfo Vietnamese = CultureInfo.GetCultureInfo("vi-VN");
        private readonly HashSet<string> _words;
        private readonly Dictionary<int, string[]> _wordsByLength;

        public VietnameseLexiconSpellChecker(IEnumerable<string> words)
        {
            if (words == null) throw new ArgumentNullException(nameof(words));
            _words = new HashSet<string>(words
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(Normalize), StringComparer.OrdinalIgnoreCase);
            _wordsByLength = _words
                .GroupBy(item => item.Length)
                .ToDictionary(group => group.Key, group => group.ToArray());
        }

        public int Count => _words.Count;

        public bool IsKnown(string word)
        {
            return !string.IsNullOrWhiteSpace(word) && _words.Contains(Normalize(word));
        }

        /// <summary>
        /// Returns a correction only when one candidate is strictly better than every
        /// other candidate. This is safe enough for 1-Click; ambiguous words stay as
        /// comments for the user instead of being changed automatically.
        /// </summary>
        public string? FindDeterministicCorrection(string word)
        {
            if (string.IsNullOrWhiteSpace(word) || IsKnown(word)) return null;
            var normalized = Normalize(word);
            var maximumDistance = normalized.Length >= 7 ? 2 : 1;
            string? best = null;
            var bestDistance = maximumDistance + 1;
            var bestCount = 0;
            for (var length = Math.Max(1, normalized.Length - maximumDistance);
                 length <= normalized.Length + maximumDistance;
                 length++)
            {
                string[] candidates;
                if (!_wordsByLength.TryGetValue(length, out candidates)) continue;
                foreach (var candidate in candidates)
                {
                    var distance = Distance(normalized, candidate, Math.Min(bestDistance, maximumDistance));
                    if (distance > maximumDistance) continue;
                    if (distance < bestDistance)
                    {
                        best = candidate;
                        bestDistance = distance;
                        bestCount = 1;
                    }
                    else if (distance == bestDistance)
                    {
                        bestCount++;
                    }
                }
            }
            return bestCount == 1 ? ApplyCase(word, best!) : null;
        }

        private static string Normalize(string value) =>
            value.Normalize(NormalizationForm.FormC).ToLower(Vietnamese);

        private static string ApplyCase(string source, string target)
        {
            if (source.All(character => !char.IsLetter(character) || char.IsUpper(character)))
                return target.ToUpper(Vietnamese);
            if (source.Length > 0 && char.IsUpper(source[0]))
                return char.ToUpper(target[0], Vietnamese) + target.Substring(1);
            return target;
        }

        private static int Distance(string left, string right, int cutoff)
        {
            if (Math.Abs(left.Length - right.Length) > cutoff) return cutoff + 1;
            var previous = new int[right.Length + 1];
            var current = new int[right.Length + 1];
            for (var column = 0; column <= right.Length; column++) previous[column] = column;
            for (var row = 1; row <= left.Length; row++)
            {
                current[0] = row;
                var minimum = current[0];
                for (var column = 1; column <= right.Length; column++)
                {
                    var substitution = previous[column - 1] +
                        (left[row - 1] == right[column - 1] ? 0 : 1);
                    current[column] = Math.Min(Math.Min(previous[column] + 1, current[column - 1] + 1),
                        substitution);
                    minimum = Math.Min(minimum, current[column]);
                }
                if (minimum > cutoff) return cutoff + 1;
                var swap = previous;
                previous = current;
                current = swap;
            }
            return previous[right.Length];
        }
    }
}
