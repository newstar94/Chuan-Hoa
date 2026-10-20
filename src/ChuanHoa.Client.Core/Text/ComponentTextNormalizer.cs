using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Text
{
    /// <summary>
    /// Comparison-only normalization for fixed administrative components
    /// (Quốc hiệu, Tiêu ngữ, tên loại...). The result is never written back to
    /// the document: it only makes recognition independent of Unicode
    /// composition (NFC/NFD), invisible spaces and dash variants.
    /// </summary>
    public static class ComponentTextNormalizer
    {
        public const int MaximumNationalTitleLength = 45;

        private static readonly TimeSpan Timeout = TimeSpan.FromMilliseconds(200);

        private static readonly Regex Whitespace = new Regex(@"\s+",
            RegexOptions.CultureInvariant, Timeout);

        private static readonly Regex NationalTitlePattern = new Regex(
            @"^CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, Timeout);

        // Tolerates spacing, dash variants and letter case, but requires the
        // whole paragraph to be the motto. Missing separators are tolerated so
        // that "Độc lập Tự do Hạnh phúc" is still recognized (and reported).
        private static readonly Regex NationalMottoPattern = new Regex(
            @"^ĐỘC LẬP ?[-,]? ?TỰ DO ?[-,]? ?HẠNH PHÚC$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, Timeout);

        private static readonly Regex CombinedHeaderPattern = new Regex(
            @"^CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM ĐỘC LẬP ?[-,]? ?TỰ DO ?[-,]? ?HẠNH PHÚC$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, Timeout);

        // "UBND XÃ A<TAB>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" typed on one line.
        private static readonly Regex TabSeparatedOrganAndTitle = new Regex(
            @"^[^\t]{2,120}\t+[  ]*CỘNG\s+H(?:ÒA|OÀ)\s+XÃ\s+HỘI\s+CHỦ\s+NGHĨA\s+VIỆT\s+NAM[  ]*$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, Timeout);

        /// <summary>
        /// NFC; NBSP/thin/figure spaces become a space; zero-width characters are
        /// removed; dash variants become "-"; whitespace is collapsed; a trailing
        /// textual rule such as "-----" or "____" is removed. "HÒA" and "HOÀ" are
        /// unified to the traditional placement so both spellings compare equal.
        /// </summary>
        public static string Normalize(string? value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var composed = value!.Normalize(NormalizationForm.FormC);
            var builder = new StringBuilder(composed.Length);
            foreach (var character in composed)
            {
                switch (character)
                {
                    case '​':
                    case '‌':
                    case '‍':
                    case '⁠':
                    case '﻿':
                    case '­':
                        continue;
                    case ' ':
                    case ' ':
                    case ' ':
                    case ' ':
                    case ' ':
                    case '　':
                    case '\v':
                    case '\a':
                        builder.Append(' ');
                        continue;
                    case '‐':
                    case '‑':
                    case '‒':
                    case '–':
                    case '—':
                    case '―':
                    case '−':
                        builder.Append('-');
                        continue;
                    default:
                        builder.Append(character);
                        continue;
                }
            }

            var collapsed = Whitespace.Replace(builder.ToString(), " ").Trim();
            collapsed = StripTrailingRule(collapsed);
            return collapsed
                .Replace("HÒA", "HOÀ")
                .Replace("Hòa", "Hoà")
                .Replace("hòa", "hoà");
        }

        private static string StripTrailingRule(string value)
        {
            // Manual scan instead of a nested-quantifier regex: a long run of
            // dashes followed by another character must not backtrack.
            var end = value.Length;
            var marks = 0;
            while (end > 0)
            {
                var character = value[end - 1];
                if (character == '-' || character == '_' || character == '.') marks++;
                else if (character != ' ') break;
                end--;
            }
            return marks >= 3 ? value.Substring(0, end).TrimEnd() : value;
        }

        /// <summary>The whole paragraph is the national title (any letter case).</summary>
        public static bool IsNationalTitle(string? value)
        {
            var normalized = Normalize(value);
            return normalized.Length > 0 && normalized.Length <= MaximumNationalTitleLength &&
                NationalTitlePattern.IsMatch(normalized);
        }

        /// <summary>The whole paragraph is the national motto (case/dash/space tolerant).</summary>
        public static bool IsNationalMotto(string? value)
        {
            var normalized = Normalize(value);
            return normalized.Length > 0 && normalized.Length <= 40 &&
                NationalMottoPattern.IsMatch(normalized);
        }

        /// <summary>Quốc hiệu and Tiêu ngữ stored in one Word paragraph (\v or \n break).</summary>
        public static bool IsCombinedNationalHeader(string? value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            if (value!.IndexOf('\v') < 0 && value.IndexOf('\n') < 0) return false;
            var normalized = Normalize(value);
            return normalized.Length <= 90 && CombinedHeaderPattern.IsMatch(normalized);
        }

        /// <summary>Title paragraph or combined title + motto paragraph.</summary>
        public static bool IsNationalTitleOrCombinedHeader(string? value) =>
            IsNationalTitle(value) || IsCombinedNationalHeader(value);

        public static bool IsUppercaseNationalTitle(string? value)
        {
            if (!IsNationalTitleOrCombinedHeader(value)) return false;
            var normalized = Normalize(value);
            var titleEnd = Math.Min(normalized.Length, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM".Length);
            var title = normalized.Substring(0, titleEnd);
            foreach (var character in title)
                if (char.IsLetter(character) && !char.IsUpper(character)) return false;
            return true;
        }

        /// <summary>
        /// Detects the one-line "TÊN CƠ QUAN⇥CỘNG HÒA..." layout. This form is only
        /// reported; it is never split automatically.
        /// </summary>
        public static bool IsTabSeparatedOrganAndNationalTitle(string? value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            var text = value!.Normalize(NormalizationForm.FormC).TrimEnd('\r', '\a', ' ');
            return text.Length <= 200 && TabSeparatedOrganAndTitle.IsMatch(text);
        }

        /// <summary>
        /// Similarity guard for whole-paragraph replacements of fixed components.
        /// Rejects replacements whose length differs by more than
        /// <paramref name="maximumLengthDelta"/> characters or that keep less than
        /// <paramref name="minimumRetainedRatio"/> of the original characters.
        /// </summary>
        public static bool IsSafeWholeParagraphReplacement(string? original, string? replacement,
            int maximumLengthDelta = 10, double minimumRetainedRatio = .7d)
        {
            var source = original ?? string.Empty;
            var target = replacement ?? string.Empty;
            if (string.Equals(source, target, StringComparison.Ordinal)) return true;
            if (source.Length == 0 || target.Length == 0) return false;
            if (Math.Abs(source.Length - target.Length) > maximumLengthDelta) return false;
            if (source.Length > 400 || target.Length > 400) return false;
            var retained = LongestCommonSubsequenceLength(
                source.Normalize(NormalizationForm.FormC).ToUpperInvariant(),
                target.Normalize(NormalizationForm.FormC).ToUpperInvariant());
            return retained >= Math.Ceiling(source.Length * minimumRetainedRatio);
        }

        private static int LongestCommonSubsequenceLength(string left, string right)
        {
            var previous = new int[right.Length + 1];
            var current = new int[right.Length + 1];
            for (var i = 1; i <= left.Length; i++)
            {
                for (var j = 1; j <= right.Length; j++)
                {
                    current[j] = left[i - 1] == right[j - 1]
                        ? previous[j - 1] + 1
                        : Math.Max(previous[j], current[j - 1]);
                }
                var swap = previous;
                previous = current;
                current = swap;
                Array.Clear(current, 0, current.Length);
            }
            return previous[right.Length];
        }
    }

    /// <summary>A minimal, offset-based text edit inside one paragraph.</summary>
    public sealed class ComponentTextEdit
    {
        public ComponentTextEdit(int offset, int length, string replacement, string expectedText)
        {
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
            if (length < 0) throw new ArgumentOutOfRangeException(nameof(length));
            Offset = offset;
            Length = length;
            Replacement = replacement ?? throw new ArgumentNullException(nameof(replacement));
            ExpectedText = expectedText ?? throw new ArgumentNullException(nameof(expectedText));
        }

        public int Offset { get; }
        public int Length { get; }
        public string Replacement { get; }
        public string ExpectedText { get; }
    }

    public static class NationalMottoEditPlanner
    {
        public const string Expected = "Độc lập - Tự do - Hạnh phúc";

        private static readonly Regex Motto = new Regex(
            @"^(?<lead>[ \t]*)(?<w1>Độc[ \t ]+lập)(?<s1>[ \t ]*[-‐-―−,]?[ \t ]*)" +
            @"(?<w2>Tự[ \t ]+do)(?<s2>[ \t ]*[-‐-―−,]?[ \t ]*)" +
            @"(?<w3>Hạnh[ \t ]+phúc)(?<trail>[ \t]*)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));

        /// <summary>
        /// Returns the smallest set of edits (separators, inner spaces, letter case)
        /// that turns a paragraph that is entirely the motto into the canonical
        /// form. Returns an empty list when the text is already canonical and
        /// <c>null</c> when the paragraph is not exactly the motto (never edit then).
        /// Text that is not NFC-composed is not edited: offsets would be unsafe.
        /// </summary>
        public static IReadOnlyList<ComponentTextEdit>? Plan(string? printable)
        {
            var text = printable ?? string.Empty;
            if (text.Length == 0 || text.Length > 60) return null;
            if (!text.IsNormalized(NormalizationForm.FormC)) return null;
            var match = Motto.Match(text);
            if (!match.Success) return null;

            var edits = new List<ComponentTextEdit>();
            AddIfDifferent(edits, match.Groups["w1"], "Độc lập");
            AddIfDifferent(edits, match.Groups["s1"], " - ");
            AddIfDifferent(edits, match.Groups["w2"], "Tự do");
            AddIfDifferent(edits, match.Groups["s2"], " - ");
            AddIfDifferent(edits, match.Groups["w3"], "Hạnh phúc");
            return edits;
        }

        public static string Apply(string printable, IReadOnlyList<ComponentTextEdit> edits)
        {
            var builder = new StringBuilder(printable ?? string.Empty);
            for (var index = edits.Count - 1; index >= 0; index--)
            {
                var edit = edits[index];
                builder.Remove(edit.Offset, edit.Length);
                builder.Insert(edit.Offset, edit.Replacement);
            }
            return builder.ToString();
        }

        private static void AddIfDifferent(ICollection<ComponentTextEdit> edits, Group group, string expected)
        {
            if (string.Equals(group.Value, expected, StringComparison.Ordinal)) return;
            edits.Add(new ComponentTextEdit(group.Index, group.Length, expected, group.Value));
        }
    }
}
