using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
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
        Unnumbered = 5,
        /// <summary>"PHẦN I: ..." / "CHƯƠNG 2. ..." with a title.</summary>
        Part = 6
    }

    /// <summary>
    /// Supplies analysis data which is deliberately kept out of the raw Word snapshot.
    /// The dictionaries are keyed by <see cref="LocalParagraphSnapshot.Index"/>.
    /// </summary>
    public sealed class HeadingDetectionContext
    {
        public const string DefaultLogicalBlockId = "document";

        public HeadingDetectionContext(
            IReadOnlyDictionary<int, string>? rolesByParagraphIndex = null,
            IReadOnlyDictionary<int, string>? logicalBlockIdsByParagraphIndex = null)
        {
            RolesByParagraphIndex = rolesByParagraphIndex ?? EmptyMap;
            LogicalBlockIdsByParagraphIndex = logicalBlockIdsByParagraphIndex ?? EmptyMap;
        }

        public IReadOnlyDictionary<int, string> RolesByParagraphIndex { get; }
        public IReadOnlyDictionary<int, string> LogicalBlockIdsByParagraphIndex { get; }

        internal string ResolveRole(LocalParagraphSnapshot paragraph)
        {
            if (RolesByParagraphIndex.TryGetValue(paragraph.Index, out var role) &&
                !string.IsNullOrWhiteSpace(role))
            {
                return role.Trim();
            }

            return paragraph.Role ?? "Unknown";
        }

        internal string ResolveLogicalBlockId(int paragraphIndex)
        {
            if (LogicalBlockIdsByParagraphIndex.TryGetValue(paragraphIndex, out var blockId) &&
                !string.IsNullOrWhiteSpace(blockId))
            {
                return blockId.Trim();
            }

            return DefaultLogicalBlockId;
        }

        private static readonly IReadOnlyDictionary<int, string> EmptyMap =
            new Dictionary<int, string>();
    }

    public sealed class HeadingContinuityOptions
    {
        public HeadingContinuityOptions(
            bool requireFirstNumberAtOne = false,
            IReadOnlyDictionary<int, string>? logicalBlockIdsByParagraphIndex = null)
        {
            RequireFirstNumberAtOne = requireFirstNumberAtOne;
            LogicalBlockIdsByParagraphIndex = logicalBlockIdsByParagraphIndex ?? EmptyMap;
        }

        /// <summary>
        /// Enable only for a complete numbering sequence. It should remain false for excerpts.
        /// </summary>
        public bool RequireFirstNumberAtOne { get; }

        /// <summary>
        /// Optional override for headings produced by the compatibility Detect overload.
        /// </summary>
        public IReadOnlyDictionary<int, string> LogicalBlockIdsByParagraphIndex { get; }

        internal string ResolveLogicalBlockId(DetectedHeading heading)
        {
            if (LogicalBlockIdsByParagraphIndex.TryGetValue(heading.ParagraphIndex, out var blockId) &&
                !string.IsNullOrWhiteSpace(blockId))
            {
                return blockId.Trim();
            }

            return string.IsNullOrWhiteSpace(heading.LogicalBlockId)
                ? HeadingDetectionContext.DefaultLogicalBlockId
                : heading.LogicalBlockId;
        }

        private static readonly IReadOnlyDictionary<int, string> EmptyMap =
            new Dictionary<int, string>();
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
            string? styleName,
            string? logicalBlockId = null)
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
            LogicalBlockId = string.IsNullOrWhiteSpace(logicalBlockId)
                ? HeadingDetectionContext.DefaultLogicalBlockId
                : logicalBlockId!;
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
        public string LogicalBlockId { get; }
    }

    public enum HeadingIssueKind
    {
        MissingHeadingStyle,
        SkippedNumber,
        DuplicateNumber,
        MissingKeepWithNext,
        BackwardNumber,
        MissingParent,
        LevelJump
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
    /// Recognizes conservative academic headings and analyzes numbering within an
    /// explicit logical document block. Legal components are never academic headings.
    /// </summary>
    public sealed class HeadingDetector
    {
        private static readonly Regex DecimalPattern = new Regex(
            @"^\s*(?<num>\d+(?:\.\d+)*)(?:(?:\.|[-–—])\s*|\s+)(?<title>\S.*)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex RomanPattern = new Regex(
            @"^\s*(?<num>[IVXLCDM]+)(?:(?:\.|[-–—])\s*|\s+)(?<title>\S.*)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        // Upper-case keyword only: "Phần 1 của báo cáo..." is a sentence, while
        // "PHẦN I: THUYẾT MINH CHUNG" is the top level of a technical document.
        private static readonly Regex PartPattern = new Regex(
            @"^\s*(?<kw>PHẦN|CHƯƠNG)\s+(?<num>[IVXLCDM]+|\d+)\s*(?:[:.]|[-–—])?\s*(?<title>\S.*)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex UnnumberedPattern = new Regex(
            @"^\s*(?<title>MỞ\s+ĐẦU|LỜI\s+NÓI\s+ĐẦU|TỔNG\s+QUAN|KẾT\s+LUẬN|TÀI\s+LIỆU\s+THAM\s+KHẢO)\s*$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LegalStructurePattern = new Regex(
            @"^\s*(?:Điều\s+\d+|Khoản\s+\d+|Điểm\s+[a-zđ]|[a-zđ]\))",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LegalBasisPattern = new Regex(
            @"^\s*(?:[-–—]\s*)?(?:Căn\s+cứ|Xét(?:\s+đề\s+nghị)?|Theo\s+đề\s+nghị)\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LegalComponentTextPattern = new Regex(
            @"^\s*(?:CỘNG\s+H(?:ÒA|OÀ)\s+XÃ\s+HỘI\s+CHỦ\s+NGHĨA\s+VIỆT\s+NAM|ĐỘC\s+LẬP\s*[-–—]\s*TỰ\s+DO\s*[-–—]\s*HẠNH\s+PHÚC|ĐẢNG\s+CỘNG\s+SẢN\s+VIỆT\s+NAM|Số\s*[:/]|Nơi\s+nhận\b|Kính\s+(?:gửi|trình)\b|Phụ\s+lục(?:\s+[IVXLCDM\d]+)?\s*$|Mục\s+lục\s*$|(?:Phần|Chương)\s+(?:[IVXLCDM]+|thứ\s+\p{L}+)\s*$|(?:Mục|Tiểu\s+mục)\s+\d+\s*$|QUYẾT\s+ĐỊNH\s*$|CÔNG\s+VĂN\s*$|THÔNG\s+BÁO\s*$|BÁO\s+CÁO\s*$|KẾ\s+HOẠCH\s*$)",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly HashSet<string> ExcludedRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "nationalTitle", "nationalMotto", "nationalMottoSeparator", "partyTitle",
            "organName", "superiorOrganName", "codeNumber", "place", "placeDate", "placeAndIssuedDate",
            "typeName", "subject", "subjectContinuation", "officialLetterSubject",
            "legalBasis", "article", "clause", "point",
            "signerAuthority", "signerPosition", "signerFullName",
            "recipientSalutation", "recipientSalutationInline", "recipientSalutationList",
            "recipientLabel", "recipientList", "recipientsLabel", "recipientsItem", "recipientsLuuLine",
            "appendixLabel", "appendixTitle", "appendixReference", "appendixDigitalSignatureInfo",
            "partChapterHeading", "sectionHeading", "structuralTitle"
        };

        public IReadOnlyList<DetectedHeading> Detect(IEnumerable<LocalParagraphSnapshot> paragraphs)
        {
            return Detect(paragraphs, new HeadingDetectionContext());
        }

        public IReadOnlyList<DetectedHeading> Detect(
            IEnumerable<LocalParagraphSnapshot> paragraphs,
            HeadingDetectionContext context)
        {
            if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));
            if (context == null) throw new ArgumentNullException(nameof(context));

            var results = new List<DetectedHeading>();
            // Blocks in which a PHẦN/Roman heading has already been seen: there a
            // bold "1. Thông tin chung" is a sub-heading, not a legal clause.
            var blocksWithUpperLevel = new HashSet<string>(StringComparer.Ordinal);
            foreach (var paragraph in paragraphs)
            {
                if (paragraph == null || !IsScannable(paragraph, context)) continue;

                var text = (paragraph.Text ?? string.Empty).Trim();
                if (!IsPotentialHeadingText(text)) continue;

                var blockId = context.ResolveLogicalBlockId(paragraph.Index);
                var detected = TryDetectHeading(
                    paragraph,
                    text,
                    blockId,
                    blocksWithUpperLevel.Contains(blockId));
                if (detected != null)
                {
                    if (detected.Kind == HeadingNumberingKind.Part ||
                        detected.Kind == HeadingNumberingKind.Roman)
                        blocksWithUpperLevel.Add(blockId);
                    results.Add(detected);
                }
            }

            return AssignHierarchyLevels(results);
        }

        /// <summary>
        /// Vietnamese technical documents nest PHẦN → I. → 1. → 1.1. → 1.1.1.
        /// The level is derived from the order in which the numbering families
        /// first appear in each block, so "I." is Heading 1 and "1." Heading 2
        /// under it, instead of both being level 1. Decimal depth stays relative
        /// to the decimal family's base level.
        /// </summary>
        private static IReadOnlyList<DetectedHeading> AssignHierarchyLevels(List<DetectedHeading> headings)
        {
            if (headings.Count == 0) return headings;
            var result = new List<DetectedHeading>(headings.Count);
            foreach (var block in headings.GroupBy(item => item.LogicalBlockId, StringComparer.Ordinal))
            {
                var families = new List<HeadingNumberingKind>();
                var maximumDecimalDepth = 1;
                foreach (var heading in block)
                {
                    if (heading.Kind == HeadingNumberingKind.Decimal)
                        maximumDecimalDepth = Math.Max(maximumDecimalDepth, heading.NumberParts.Count);
                    if ((heading.Kind == HeadingNumberingKind.Part || heading.Kind == HeadingNumberingKind.Roman ||
                         heading.Kind == HeadingNumberingKind.Decimal) && !families.Contains(heading.Kind))
                        families.Add(heading.Kind);
                }

                var baseLevel = new Dictionary<HeadingNumberingKind, int>();
                var next = 1;
                foreach (var family in families)
                {
                    baseLevel[family] = next;
                    next += family == HeadingNumberingKind.Decimal ? maximumDecimalDepth : 1;
                }

                foreach (var heading in block)
                {
                    int level;
                    switch (heading.Kind)
                    {
                        case HeadingNumberingKind.Decimal:
                            level = baseLevel[HeadingNumberingKind.Decimal] + heading.NumberParts.Count - 1;
                            break;
                        case HeadingNumberingKind.Part:
                        case HeadingNumberingKind.Roman:
                            level = baseLevel[heading.Kind];
                            break;
                        default:
                            level = heading.Level;
                            break;
                    }
                    level = Math.Max(1, Math.Min(9, level));
                    result.Add(level == heading.Level ? heading : new DetectedHeading(
                        heading.ParagraphIndex, heading.AbsoluteStart, level, heading.NumberText,
                        heading.TitleText, heading.Kind, heading.NumberParts, heading.IsBold,
                        heading.OutlineLevel, heading.KeepWithNext, heading.StyleName,
                        heading.LogicalBlockId));
                }
            }
            // Keep document order across blocks.
            return result.OrderBy(item => item.AbsoluteStart).ThenBy(item => item.ParagraphIndex).ToList();
        }

        private static bool IsScannable(LocalParagraphSnapshot paragraph, HeadingDetectionContext context)
        {
            if (paragraph.IsInTable) return false;
            if (!string.IsNullOrEmpty(paragraph.StoryType) &&
                !string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return !ExcludedRoles.Contains(context.ResolveRole(paragraph));
        }

        private static bool IsPotentialHeadingText(string text)
        {
            if (string.IsNullOrWhiteSpace(text) || text.Length > 200) return false;
            if (LegalStructurePattern.IsMatch(text) || LegalBasisPattern.IsMatch(text) ||
                LegalComponentTextPattern.IsMatch(text))
            {
                return false;
            }

            // A trailing ":" is common in technical documents ("1. Thông tin
            // chung:"); the level-1 decimal rule below still rejects clause-like text.
            var lastChar = text[text.Length - 1];
            return lastChar != ';' && lastChar != ',' &&
                   lastChar != '.' && lastChar != '?' && lastChar != '!';
        }

        private static DetectedHeading? TryDetectHeading(
            LocalParagraphSnapshot paragraph,
            string text,
            string logicalBlockId,
            bool underUpperLevelHeading)
        {
            text = text.TrimEnd(':').TrimEnd();
            if (text.Length == 0) return null;

            var partMatch = PartPattern.Match(text);
            if (partMatch.Success)
            {
                var numberText = partMatch.Groups["num"].Value;
                var value = char.IsDigit(numberText[0])
                    ? (int.TryParse(numberText, NumberStyles.None, CultureInfo.InvariantCulture, out var digits) ? digits : 0)
                    : ParseCanonicalRoman(numberText);
                var partTitle = partMatch.Groups["title"].Value.Trim();
                if (value > 0 && IsTitleShaped(partTitle))
                {
                    return CreateHeading(
                        paragraph,
                        1,
                        partMatch.Groups["kw"].Value + " " + numberText,
                        partTitle,
                        HeadingNumberingKind.Part,
                        new[] { value },
                        logicalBlockId);
                }
            }

            var unnumberedMatch = UnnumberedPattern.Match(text);
            if (unnumberedMatch.Success)
            {
                return CreateHeading(
                    paragraph,
                    1,
                    string.Empty,
                    unnumberedMatch.Groups["title"].Value.Trim(),
                    HeadingNumberingKind.Unnumbered,
                    Array.Empty<int>(),
                    logicalBlockId);
            }

            var decimalMatch = DecimalPattern.Match(text);
            if (decimalMatch.Success)
            {
                var numberText = decimalMatch.Groups["num"].Value;
                var title = decimalMatch.Groups["title"].Value.Trim();
                var parts = ParseDecimalParts(numberText);
                if (parts.Count > 0 && IsConfidentDecimalHeading(paragraph, title, parts.Count,
                        underUpperLevelHeading))
                {
                    return CreateHeading(
                        paragraph,
                        parts.Count,
                        numberText,
                        title,
                        HeadingNumberingKind.Decimal,
                        parts,
                        logicalBlockId);
                }
            }

            var romanMatch = RomanPattern.Match(text);
            if (romanMatch.Success)
            {
                var romanText = romanMatch.Groups["num"].Value;
                var romanValue = ParseCanonicalRoman(romanText);
                var title = romanMatch.Groups["title"].Value.Trim();
                if (romanValue > 0 && HasHeadingPresentation(paragraph, title))
                {
                    return CreateHeading(
                        paragraph,
                        1,
                        romanText,
                        title,
                        HeadingNumberingKind.Roman,
                        new[] { romanValue },
                        logicalBlockId);
                }
            }

            // Alphabet markers such as "a)" remain legal/list points by default.
            return null;
        }

        private static DetectedHeading CreateHeading(
            LocalParagraphSnapshot paragraph,
            int level,
            string numberText,
            string title,
            HeadingNumberingKind kind,
            IReadOnlyList<int> numberParts,
            string logicalBlockId)
        {
            return new DetectedHeading(
                paragraph.Index,
                paragraph.AbsoluteStart,
                level,
                numberText,
                title,
                kind,
                numberParts,
                paragraph.Bold == true,
                paragraph.OutlineLevel,
                paragraph.KeepWithNext,
                paragraph.StyleName,
                logicalBlockId);
        }

        private static bool IsConfidentDecimalHeading(
            LocalParagraphSnapshot paragraph,
            string title,
            int level,
            bool underUpperLevelHeading)
        {
            if (!IsTitleShaped(title)) return false;

            var hasExplicitHeadingMetadata = HasExplicitHeadingMetadata(paragraph);
            if (level == 1)
            {
                // A bare "1. ..." is also the canonical shape of a legal clause.
                // Bold alone is therefore insufficient at level one, unless the
                // block already has a PHẦN/Roman heading above it.
                return hasExplicitHeadingMetadata || IsUppercaseTitle(title) ||
                    (underUpperLevelHeading && paragraph.Bold == true);
            }

            return hasExplicitHeadingMetadata || paragraph.Bold == true || IsUppercaseTitle(title);
        }

        private static bool HasHeadingPresentation(LocalParagraphSnapshot paragraph, string title)
        {
            return IsTitleShaped(title) &&
                   (HasExplicitHeadingMetadata(paragraph) || paragraph.Bold == true || IsUppercaseTitle(title));
        }

        private static bool HasExplicitHeadingMetadata(LocalParagraphSnapshot paragraph)
        {
            if (paragraph.OutlineLevel.HasValue &&
                paragraph.OutlineLevel.Value >= 1 &&
                paragraph.OutlineLevel.Value <= 9)
            {
                return true;
            }

            var styleName = paragraph.StyleName ?? string.Empty;
            return styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) ||
                   styleName.StartsWith("Tiêu đề", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTitleShaped(string title)
        {
            if (string.IsNullOrWhiteSpace(title) || title.Length > 160) return false;

            var wordCount = Regex.Matches(title, @"\S+").Count;
            if (wordCount == 0 || wordCount > 20) return false;

            var firstLetter = title.FirstOrDefault(char.IsLetter);
            return firstLetter == default(char) || char.IsUpper(firstLetter);
        }

        private static bool IsUppercaseTitle(string title)
        {
            var hasLetter = false;
            foreach (var character in title)
            {
                if (!char.IsLetter(character)) continue;
                hasLetter = true;
                if (char.IsLower(character)) return false;
            }

            return hasLetter;
        }

        public IReadOnlyList<HeadingIssue> AnalyzeContinuity(IReadOnlyList<DetectedHeading> headings)
        {
            return AnalyzeContinuity(headings, new HeadingContinuityOptions());
        }

        public IReadOnlyList<HeadingIssue> AnalyzeContinuity(
            IReadOnlyList<DetectedHeading> headings,
            HeadingContinuityOptions options)
        {
            if (headings == null || headings.Count == 0) return Array.Empty<HeadingIssue>();
            if (options == null) throw new ArgumentNullException(nameof(options));

            var issues = new List<HeadingIssue>();
            var states = new Dictionary<string, ContinuityState>(StringComparer.Ordinal);
            // Numbering restarts under every new parent: Roman numbers restart in
            // each PHẦN and "1." restarts under each new Roman heading.
            var partScopes = new Dictionary<string, string>(StringComparer.Ordinal);
            var romanScopes = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (var heading in headings)
            {
                if (heading == null) continue;

                if (heading.Kind != HeadingNumberingKind.Decimal &&
                    heading.Kind != HeadingNumberingKind.Roman &&
                    heading.Kind != HeadingNumberingKind.Article &&
                    heading.Kind != HeadingNumberingKind.Part)
                {
                    continue;
                }

                var blockId = options.ResolveLogicalBlockId(heading);
                string partScope;
                partScopes.TryGetValue(blockId, out partScope);
                string romanScope;
                romanScopes.TryGetValue(blockId, out romanScope);
                var scope = heading.Kind == HeadingNumberingKind.Roman ? (partScope ?? string.Empty)
                    : heading.Kind == HeadingNumberingKind.Decimal ? (romanScope ?? partScope ?? string.Empty)
                    : string.Empty;
                if (heading.Kind == HeadingNumberingKind.Part)
                {
                    partScopes[blockId] = "P" + heading.NumberText;
                    romanScopes.Remove(blockId);
                }
                else if (heading.Kind == HeadingNumberingKind.Roman)
                {
                    romanScopes[blockId] = (partScope ?? string.Empty) + "/R" + heading.NumberText;
                }
                var stateKey = blockId + "\u001f" + ((int)heading.Kind).ToString(CultureInfo.InvariantCulture) +
                    "\u001f" + scope;
                if (!states.TryGetValue(stateKey, out var state))
                {
                    state = new ContinuityState();
                    states.Add(stateKey, state);
                }

                if (heading.Kind == HeadingNumberingKind.Decimal)
                {
                    AnalyzeDecimalHeading(heading, state, options, issues);
                }
                else if (heading.Kind == HeadingNumberingKind.Roman ||
                         heading.Kind == HeadingNumberingKind.Article ||
                         heading.Kind == HeadingNumberingKind.Part)
                {
                    AnalyzeScalarHeading(heading, state, options, issues);
                }
            }

            return issues;
        }

        private static void AnalyzeDecimalHeading(
            DetectedHeading heading,
            ContinuityState state,
            HeadingContinuityOptions options,
            ICollection<HeadingIssue> issues)
        {
            if (heading.NumberParts.Count == 0 || heading.NumberParts.Any(part => part <= 0)) return;

            var path = JoinParts(heading.NumberParts, heading.NumberParts.Count);
            var parentPath = JoinParts(heading.NumberParts, heading.NumberParts.Count - 1);
            var currentNumber = heading.NumberParts[heading.NumberParts.Count - 1];
            HeadingIssue? structuralIssue = null;

            if (state.SeenDecimalPaths.Contains(path))
            {
                var expectedNumber = state.LastDecimalByParent.TryGetValue(parentPath, out var duplicateLast)
                    ? duplicateLast + 1
                    : currentNumber + 1;
                structuralIssue = Issue(
                    heading,
                    HeadingIssueKind.DuplicateNumber,
                    "Trùng số thứ tự đề mục " + heading.NumberText + ".",
                    FormatSiblingNumber(heading.NumberParts, expectedNumber));
            }
            else if (state.LastDecimalByParent.TryGetValue(parentPath, out var lastSibling) &&
                     currentNumber < lastSibling)
            {
                structuralIssue = Issue(
                    heading,
                    HeadingIssueKind.BackwardNumber,
                    "Số thứ tự đề mục đi lùi từ " +
                    FormatSiblingNumber(heading.NumberParts, lastSibling) + " về " + heading.NumberText + ".",
                    FormatSiblingNumber(heading.NumberParts, lastSibling + 1));
            }
            else if (state.PreviousDecimalLevel > 0 &&
                     heading.Level > state.PreviousDecimalLevel + 1)
            {
                structuralIssue = Issue(
                    heading,
                    HeadingIssueKind.LevelJump,
                    "Đề mục " + heading.NumberText + " nhảy từ cấp " +
                    state.PreviousDecimalLevel + " lên cấp " + heading.Level + ".",
                    "Chỉ chuyển tối đa đến cấp " + (state.PreviousDecimalLevel + 1) + ".");
            }
            else if (heading.NumberParts.Count > 1 && !state.SeenDecimalPaths.Contains(parentPath) &&
                     !state.LastDecimalByParent.ContainsKey(parentPath))
            {
                structuralIssue = Issue(
                    heading,
                    HeadingIssueKind.MissingParent,
                    "Đề mục " + heading.NumberText + " không có đề mục cha " + parentPath + ".",
                    "Bổ sung đề mục cha " + parentPath + " trước đề mục này.");
            }
            else if (state.LastDecimalByParent.TryGetValue(parentPath, out lastSibling) &&
                     currentNumber > lastSibling + 1)
            {
                structuralIssue = Issue(
                    heading,
                    HeadingIssueKind.SkippedNumber,
                    "Nhảy cóc số thứ tự đề mục từ " +
                    FormatSiblingNumber(heading.NumberParts, lastSibling) + " sang " + heading.NumberText + ".",
                    FormatSiblingNumber(heading.NumberParts, lastSibling + 1));
            }
            else if (!state.LastDecimalByParent.ContainsKey(parentPath) &&
                     options.RequireFirstNumberAtOne && currentNumber != 1)
            {
                structuralIssue = Issue(
                    heading,
                    HeadingIssueKind.SkippedNumber,
                    "Dãy đề mục bắt đầu từ " + heading.NumberText + " thay vì số 1.",
                    FormatSiblingNumber(heading.NumberParts, 1));
            }

            if (structuralIssue != null) issues.Add(structuralIssue);

            state.SeenDecimalPaths.Add(path);
            if (!state.LastDecimalByParent.TryGetValue(parentPath, out var previousLast) ||
                currentNumber > previousLast)
            {
                state.LastDecimalByParent[parentPath] = currentNumber;
            }
            state.PreviousDecimalLevel = heading.Level;
        }

        private static void AnalyzeScalarHeading(
            DetectedHeading heading,
            ContinuityState state,
            HeadingContinuityOptions options,
            ICollection<HeadingIssue> issues)
        {
            if (heading.NumberParts.Count == 0 || heading.NumberParts[0] <= 0) return;

            var current = heading.NumberParts[0];
            if (!state.LastScalarNumber.HasValue)
            {
                if (options.RequireFirstNumberAtOne && current != 1)
                {
                    issues.Add(Issue(
                        heading,
                        HeadingIssueKind.SkippedNumber,
                        "Dãy đề mục bắt đầu từ " + heading.NumberText + " thay vì số đầu tiên.",
                        FormatScalar(heading.Kind, 1)));
                }

                state.LastScalarNumber = current;
                return;
            }

            var last = state.LastScalarNumber.Value;
            if (current == last)
            {
                issues.Add(Issue(
                    heading,
                    HeadingIssueKind.DuplicateNumber,
                    "Trùng số thứ tự đề mục " + heading.NumberText + ".",
                    FormatScalar(heading.Kind, last + 1)));
            }
            else if (current < last)
            {
                issues.Add(Issue(
                    heading,
                    HeadingIssueKind.BackwardNumber,
                    "Số thứ tự đề mục đi lùi từ " + FormatScalar(heading.Kind, last) +
                    " về " + heading.NumberText + ".",
                    FormatScalar(heading.Kind, last + 1)));
            }
            else if (current > last + 1)
            {
                issues.Add(Issue(
                    heading,
                    HeadingIssueKind.SkippedNumber,
                    "Nhảy cóc số thứ tự đề mục từ " + FormatScalar(heading.Kind, last) +
                    " sang " + heading.NumberText + ".",
                    FormatScalar(heading.Kind, last + 1)));
            }

            if (current > last) state.LastScalarNumber = current;
        }

        private static HeadingIssue Issue(
            DetectedHeading heading,
            HeadingIssueKind kind,
            string current,
            string expected)
        {
            return new HeadingIssue(heading, kind, current, expected);
        }

        private static string JoinParts(IReadOnlyList<int> parts, int count)
        {
            if (count <= 0) return string.Empty;
            return string.Join(".", parts.Take(count).Select(
                part => part.ToString(CultureInfo.InvariantCulture)));
        }

        private static string FormatSiblingNumber(IReadOnlyList<int> parts, int number)
        {
            if (parts.Count <= 1) return number.ToString(CultureInfo.InvariantCulture);

            var builder = new StringBuilder();
            for (var index = 0; index < parts.Count - 1; index++)
            {
                if (index > 0) builder.Append('.');
                builder.Append(parts[index].ToString(CultureInfo.InvariantCulture));
            }
            builder.Append('.').Append(number.ToString(CultureInfo.InvariantCulture));
            return builder.ToString();
        }

        private static string FormatScalar(HeadingNumberingKind kind, int number)
        {
            if (kind == HeadingNumberingKind.Roman) return ToRoman(number);
            if (kind == HeadingNumberingKind.Part) return "PHẦN " + ToRoman(number);
            if (kind == HeadingNumberingKind.Article) return "Điều " + number.ToString(CultureInfo.InvariantCulture);
            return number.ToString(CultureInfo.InvariantCulture);
        }

        private static IReadOnlyList<int> ParseDecimalParts(string value)
        {
            var tokens = value.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            var parts = new List<int>();
            foreach (var token in tokens)
            {
                if ((token.Length > 1 && token[0] == '0') ||
                    !int.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out var number) ||
                    number <= 0)
                {
                    return Array.Empty<int>();
                }

                parts.Add(number);
            }

            return parts;
        }

        private static int ParseCanonicalRoman(string roman)
        {
            if (string.IsNullOrEmpty(roman) || roman.Length > 15) return 0;

            var total = 0;
            var previousValue = 0;
            for (var index = roman.Length - 1; index >= 0; index--)
            {
                var value = RomanValue(roman[index]);
                if (value == 0) return 0;
                total += value < previousValue ? -value : value;
                previousValue = value;
            }

            return total > 0 && total <= 3999 &&
                   string.Equals(ToRoman(total), roman, StringComparison.Ordinal)
                ? total
                : 0;
        }

        private static int RomanValue(char value)
        {
            switch (value)
            {
                case 'I': return 1;
                case 'V': return 5;
                case 'X': return 10;
                case 'L': return 50;
                case 'C': return 100;
                case 'D': return 500;
                case 'M': return 1000;
                default: return 0;
            }
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

        private sealed class ContinuityState
        {
            public HashSet<string> SeenDecimalPaths { get; } = new HashSet<string>(StringComparer.Ordinal);
            public Dictionary<string, int> LastDecimalByParent { get; } =
                new Dictionary<string, int>(StringComparer.Ordinal);
            public int PreviousDecimalLevel { get; set; }
            public int? LastScalarNumber { get; set; }
        }
    }
}
