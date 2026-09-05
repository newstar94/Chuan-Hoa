using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Scanning
{
    public enum HeadingNumberingKind
    {
        None = 0,
        Decimal = 1,
        Roman = 2,
        Alphabet = 3,
        Article = 4,
        Unnumbered = 5
    }

    public sealed class DetectedHeading
    {
        public DetectedHeading(
            int paragraphIndex,
            int absoluteStart,
            int level,
            string numberText,
            string titleText,
            HeadingNumberingKind kind,
            IReadOnlyList<int> numberParts,
            bool isBold,
            int? outlineLevel,
            bool? keepWithNext,
            string? styleName)
        {
            ParagraphIndex = paragraphIndex;
            AbsoluteStart = absoluteStart;
            Level = level;
            NumberText = numberText ?? string.Empty;
            TitleText = titleText ?? string.Empty;
            Kind = kind;
            NumberParts = numberParts ?? Array.Empty<int>();
            IsBold = isBold;
            OutlineLevel = outlineLevel;
            KeepWithNext = keepWithNext;
            StyleName = styleName;
        }

        public int ParagraphIndex { get; }
        public int AbsoluteStart { get; }
        public int Level { get; }
        public string NumberText { get; }
        public string TitleText { get; }
        public HeadingNumberingKind Kind { get; }
        public IReadOnlyList<int> NumberParts { get; }
        public bool IsBold { get; }
        public int? OutlineLevel { get; }
        public bool? KeepWithNext { get; }
        public string? StyleName { get; }
    }

    public enum HeadingIssueKind
    {
        MissingHeadingStyle,
        SkippedNumber,
        DuplicateNumber,
        MissingKeepWithNext
    }

    public sealed class HeadingIssue
    {
        public HeadingIssue(
            DetectedHeading heading,
            HeadingIssueKind issueKind,
            string currentIssue,
            string expected)
        {
            Heading = heading ?? throw new ArgumentNullException(nameof(heading));
            IssueKind = issueKind;
            CurrentIssue = currentIssue ?? string.Empty;
            Expected = expected ?? string.Empty;
        }

        public DetectedHeading Heading { get; }
        public HeadingIssueKind IssueKind { get; }
        public string CurrentIssue { get; }
        public string Expected { get; }
    }

    /// <summary>
    /// Recognizes hand-typed headings (decimal, Roman, articles, unnumbered)
    /// and analyzes outline continuity and layout properties.
    /// </summary>
    public sealed class HeadingDetector
    {
        private static readonly Regex DecimalPattern = new Regex(
            @"^\s*(?<num>\d+(\.\d+)*)\.?\s*[\.\-\/:]*\s*(?<title>.+)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex RomanPattern = new Regex(
            @"^\s*(?<num>[IVXLCDM]+)\.?\s*[\.\-\/:]*\s*(?<title>.+)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex ArticlePattern = new Regex(
            @"^\s*Điều\s+(?<num>\d+)\.?\s*[\.\-\/:]*\s*(?<title>.*)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex AlphabetPattern = new Regex(
            @"^\s*(?<num>[a-zđ])\)\s*(?<title>.+)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex UnnumberedPattern = new Regex(
            @"^\s*(?<title>MỞ\s+ĐẦU|LỜI\s+NÓI\s+ĐẦU|TỔNG\s+QUAN|KẾT\s+LUẬN|TÀI\s+LIỆU\s+THAM\s+KHẢO|PHỤ\s+LỤC|MỤC\s+LỤC)\s*$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly HashSet<string> ExcludedRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "NationalTitle", "NationalMotto", "NationalMottoSeparator",
            "OrganName", "SuperiorOrganName", "CodeNumber", "PlaceDate",
            "SignerAuthority", "SignerPosition", "SignerFullName",
            "RecipientsLabel", "RecipientsItem", "RecipientsLuuLine"
        };

        public IReadOnlyList<DetectedHeading> Detect(IEnumerable<LocalParagraphSnapshot> paragraphs)
        {
            if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));

            var results = new List<DetectedHeading>();
            foreach (var p in paragraphs)
            {
                if (p.IsInTable) continue;
                if (!string.Equals(p.StoryType, "wdMainTextStory", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrEmpty(p.StoryType)) continue;
                if (ExcludedRoles.Contains(p.Role)) continue;

                var text = (p.Text ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(text) || text.Length > 200) continue;

                // If ends with typical clause/list punctuation, it's not a heading
                var lastChar = text[text.Length - 1];
                if (lastChar == ';' || lastChar == ':' || lastChar == ',') continue;

                var detected = TryDetectHeading(p, text);
                if (detected != null)
                {
                    results.Add(detected);
                }
            }

            return results;
        }

        private static DetectedHeading? TryDetectHeading(LocalParagraphSnapshot p, string text)
        {
            // 1. Unnumbered headings (MỞ ĐẦU, KẾT LUẬN, etc.)
            var unnumberedMatch = UnnumberedPattern.Match(text);
            if (unnumberedMatch.Success)
            {
                return new DetectedHeading(
                    p.Index, p.AbsoluteStart, 1, string.Empty,
                    unnumberedMatch.Groups["title"].Value.Trim(),
                    HeadingNumberingKind.Unnumbered, Array.Empty<int>(),
                    p.Bold == true, p.OutlineLevel, p.KeepWithNext, p.StyleName);
            }

            // 2. Legal Articles: "Điều 1. ..."
            var articleMatch = ArticlePattern.Match(text);
            if (articleMatch.Success && int.TryParse(articleMatch.Groups["num"].Value, out var artNum))
            {
                var title = articleMatch.Groups["title"].Value.Trim();
                return new DetectedHeading(
                    p.Index, p.AbsoluteStart, 1, "Điều " + artNum,
                    title, HeadingNumberingKind.Article, new[] { artNum },
                    p.Bold == true, p.OutlineLevel, p.KeepWithNext, p.StyleName);
            }

            // 3. Decimal: "1. ", "1.1. ", "1.1.1 "
            var decimalMatch = DecimalPattern.Match(text);
            if (decimalMatch.Success)
            {
                var numStr = decimalMatch.Groups["num"].Value;
                var title = decimalMatch.Groups["title"].Value.Trim();
                var parts = ParseDecimalParts(numStr);
                if (parts.Count > 0)
                {
                    // Decimal with 1 part (e.g. "1. ") is only treated as heading if bold or short title
                    // to avoid confusing with simple ordered list items
                    if (parts.Count == 1 && p.Bold != true && title.Length > 100)
                    {
                        // Likely regular body text list item
                    }
                    else
                    {
                        return new DetectedHeading(
                            p.Index, p.AbsoluteStart, parts.Count, numStr,
                            title, HeadingNumberingKind.Decimal, parts,
                            p.Bold == true, p.OutlineLevel, p.KeepWithNext, p.StyleName);
                    }
                }
            }

            // 4. Roman: "I. ", "II. "
            var romanMatch = RomanPattern.Match(text);
            if (romanMatch.Success)
            {
                var romanStr = romanMatch.Groups["num"].Value.ToUpperInvariant();
                var romanVal = ParseRoman(romanStr);
                if (romanVal > 0)
                {
                    var title = romanMatch.Groups["title"].Value.Trim();
                    return new DetectedHeading(
                        p.Index, p.AbsoluteStart, 1, romanStr,
                        title, HeadingNumberingKind.Roman, new[] { romanVal },
                        p.Bold == true, p.OutlineLevel, p.KeepWithNext, p.StyleName);
                }
            }

            return null;
        }

        public IReadOnlyList<HeadingIssue> AnalyzeContinuity(IReadOnlyList<DetectedHeading> headings)
        {
            if (headings == null || headings.Count == 0) return Array.Empty<HeadingIssue>();

            var issues = new List<HeadingIssue>();

            // Group by numbering kind for continuity checking
            var decimalStack = new Dictionary<int, int>(); // level -> last seen index
            int lastArticleNum = 0;
            int lastRomanNum = 0;

            foreach (var h in headings)
            {
                if (h.Kind == HeadingNumberingKind.Article)
                {
                    var currentArt = h.NumberParts[0];
                    if (lastArticleNum > 0)
                    {
                        if (currentArt == lastArticleNum)
                        {
                            issues.Add(new HeadingIssue(h, HeadingIssueKind.DuplicateNumber,
                                "Trùng số thứ tự Điều " + currentArt + ".",
                                "Điều " + (lastArticleNum + 1) + "."));
                        }
                        else if (currentArt > lastArticleNum + 1)
                        {
                            issues.Add(new HeadingIssue(h, HeadingIssueKind.SkippedNumber,
                                "Nhảy cóc số thứ tự từ Điều " + lastArticleNum + " sang Điều " + currentArt + " (thiếu Điều " + (lastArticleNum + 1) + ").",
                                "Điều " + (lastArticleNum + 1) + "."));
                        }
                    }
                    lastArticleNum = currentArt;
                    decimalStack.Clear(); // reset decimal under article
                }
                else if (h.Kind == HeadingNumberingKind.Roman)
                {
                    var currentRoman = h.NumberParts[0];
                    if (lastRomanNum > 0)
                    {
                        if (currentRoman == lastRomanNum)
                        {
                            issues.Add(new HeadingIssue(h, HeadingIssueKind.DuplicateNumber,
                                "Trùng số La Mã " + h.NumberText + ".",
                                "Số La Mã tiếp theo là " + ToRoman(lastRomanNum + 1) + "."));
                        }
                        else if (currentRoman > lastRomanNum + 1)
                        {
                            issues.Add(new HeadingIssue(h, HeadingIssueKind.SkippedNumber,
                                "Nhảy cóc số La Mã từ " + ToRoman(lastRomanNum) + " sang " + h.NumberText + " (thiếu " + ToRoman(lastRomanNum + 1) + ").",
                                ToRoman(lastRomanNum + 1) + "."));
                        }
                    }
                    lastRomanNum = currentRoman;
                    decimalStack.Clear();
                }
                else if (h.Kind == HeadingNumberingKind.Decimal)
                {
                    var level = h.Level;
                    var currentNum = h.NumberParts[h.NumberParts.Count - 1];

                    // Clear deeper levels
                    var keysToRemove = decimalStack.Keys.Where(k => k > level).ToList();
                    foreach (var k in keysToRemove) decimalStack.Remove(k);

                    if (decimalStack.TryGetValue(level, out var lastNum))
                    {
                        if (currentNum == lastNum)
                        {
                            issues.Add(new HeadingIssue(h, HeadingIssueKind.DuplicateNumber,
                                "Trùng số thứ tự đề mục " + h.NumberText + ".",
                                "Đề mục tiếp theo phải tăng số thứ tự."));
                        }
                        else if (currentNum > lastNum + 1)
                        {
                            issues.Add(new HeadingIssue(h, HeadingIssueKind.SkippedNumber,
                                "Nhảy cóc số thứ tự đề mục từ " + FormatDecimal(h.NumberParts, lastNum) +
                                " sang " + h.NumberText + ".",
                                FormatDecimal(h.NumberParts, lastNum + 1)));
                        }
                    }
                    decimalStack[level] = currentNum;
                }
            }

            return issues;
        }

        private static string FormatDecimal(IReadOnlyList<int> prefixParts, int lastNumber)
        {
            if (prefixParts.Count == 1) return lastNumber.ToString(CultureInfo.InvariantCulture);
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < prefixParts.Count - 1; i++)
            {
                sb.Append(prefixParts[i]).Append('.');
            }
            sb.Append(lastNumber);
            return sb.ToString();
        }

        private static IReadOnlyList<int> ParseDecimalParts(string str)
        {
            var tokens = str.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            var list = new List<int>();
            foreach (var t in tokens)
            {
                if (int.TryParse(t, out var val) && val >= 0)
                {
                    list.Add(val);
                }
                else
                {
                    return Array.Empty<int>();
                }
            }
            return list;
        }

        private static int ParseRoman(string roman)
        {
            var romanMap = new Dictionary<char, int>
            {
                {'I', 1}, {'V', 5}, {'X', 10}, {'L', 50}, {'C', 100}, {'D', 500}, {'M', 1000}
            };

            int total = 0;
            int prevValue = 0;
            for (int i = roman.Length - 1; i >= 0; i--)
            {
                if (!romanMap.TryGetValue(roman[i], out var value)) return 0;
                if (value < prevValue)
                    total -= value;
                else
                    total += value;
                prevValue = value;
            }
            return total;
        }

        private static string ToRoman(int number)
        {
            if (number <= 0 || number > 3999) return number.ToString(CultureInfo.InvariantCulture);
            string[] thousands = { "", "M", "MM", "MMM" };
            string[] hundreds = { "", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM" };
            string[] tens = { "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC" };
            string[] ones = { "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX" };

            return thousands[number / 1000] +
                   hundreds[(number % 1000) / 100] +
                   tens[(number % 100) / 10] +
                   ones[number % 10];
        }
    }
}
