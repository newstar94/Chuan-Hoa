using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ChuanHoa.Client.Core.Lexicon
{
    /// <summary>
    /// NFC text together with a UTF-16 offset map back to the exact source text.
    /// Word annotations must always use source offsets, even when a decomposed
    /// Vietnamese grapheme becomes shorter after canonical composition.
    /// </summary>
    public sealed class NormalizedTextOffsetMap
    {
        private readonly int[] _sourceStarts;
        private readonly int[] _sourceEnds;

        private NormalizedTextOffsetMap(string source, string normalized,
            int[] sourceStarts, int[] sourceEnds)
        {
            Source = source;
            Normalized = normalized;
            _sourceStarts = sourceStarts;
            _sourceEnds = sourceEnds;
        }

        public string Source { get; }
        public string Normalized { get; }

        public static NormalizedTextOffsetMap Create(string? source)
        {
            var original = source ?? string.Empty;
            if (original.Length == 0)
                return new NormalizedTextOffsetMap(original, string.Empty,
                    Array.Empty<int>(), Array.Empty<int>());

            var normalized = new StringBuilder(original.Length);
            var starts = new List<int>(original.Length);
            var ends = new List<int>(original.Length);
            var enumerator = StringInfo.GetTextElementEnumerator(original);
            while (enumerator.MoveNext())
            {
                var sourceStart = enumerator.ElementIndex;
                var element = enumerator.GetTextElement();
                var sourceEnd = sourceStart + element.Length;
                var composed = element.Normalize(NormalizationForm.FormC);
                normalized.Append(composed);
                for (var index = 0; index < composed.Length; index++)
                {
                    starts.Add(sourceStart);
                    ends.Add(sourceEnd);
                }
            }

            return new NormalizedTextOffsetMap(original, normalized.ToString(),
                starts.ToArray(), ends.ToArray());
        }

        public Tuple<int, int> MapSpan(int normalizedStart, int normalizedLength)
        {
            if (normalizedStart < 0 || normalizedLength <= 0 ||
                normalizedStart + normalizedLength > Normalized.Length)
                throw new ArgumentOutOfRangeException(nameof(normalizedStart));
            var sourceStart = _sourceStarts[normalizedStart];
            var sourceEnd = _sourceEnds[normalizedStart + normalizedLength - 1];
            return Tuple.Create(sourceStart, sourceEnd - sourceStart);
        }
    }
}
