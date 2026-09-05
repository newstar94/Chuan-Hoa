using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.Client.Core.Scanning
{
    public sealed class AcademicTypographyScanOutcome
    {
        public AcademicTypographyScanOutcome(IReadOnlyList<AnnotationFinding> findings,
            IReadOnlyList<DetectedHeading> headings, IReadOnlyList<string> notEvaluatedRuleCodes,
            AcademicTypographyAdvisoryProfile profile)
        {
            Findings = findings ?? Array.Empty<AnnotationFinding>();
            Headings = headings ?? Array.Empty<DetectedHeading>();
            NotEvaluatedRuleCodes = notEvaluatedRuleCodes ?? Array.Empty<string>();
            Profile = profile ?? throw new ArgumentNullException(nameof(profile));
        }

        public IReadOnlyList<AnnotationFinding> Findings { get; }
        public IReadOnlyList<DetectedHeading> Headings { get; }
        public IReadOnlyList<string> NotEvaluatedRuleCodes { get; }
        public AcademicTypographyAdvisoryProfile Profile { get; }
    }

    /// <summary>
    /// Pure-.NET advisory scanner. Only a verified, signed AcademicTypography
    /// profile may enable these suggestions; none of them is a legal violation.
    /// </summary>
    public sealed class LatexTypographicScanner
    {
        private static readonly Regex TableCaptionRegex = Rx(
            @"^\s*Bảng\s+(?:\d+(?:\.\d+)*|[IVXLCDM]+)(?:[.\-–—/:]\s*|\s+)", true);
        private static readonly Regex FigureCaptionRegex = Rx(
            @"^\s*Hình\s+(?:\d+(?:\.\d+)*|[IVXLCDM]+)(?:[.\-–—/:]\s*|\s+)", true);
        private static readonly Regex LeadingListRegex = Rx(
            @"^\s*(?:[-–—+•·▪◦‣⁃]|\(?\d+(?:\.\d+)*[.)]|\(?[a-zđ][.)])\s+", true);
        private static readonly Regex LegalTextRegex = Rx(
            @"^\s*(?:Điều\s+\d+|Khoản\s+\d+|Điểm\s+[a-zđ]|Căn\s+cứ|Xét\s+đề\s+nghị|Theo\s+đề\s+nghị)\b", true);
        private static readonly Regex WordRegex = Rx(@"\p{L}+(?:[-']\p{L}+)*");
        private static readonly Regex MathCommandRegex = Rx(@"\\[A-Za-z]{2,}");
        private static readonly Regex MathOperatorRegex = Rx(
            @"(?:[=<>+*/^_]|\\(?:frac|sqrt|sum|int|prod|lim|alpha|beta|gamma|delta|theta|lambda|mu|pi|sigma|phi|omega)\b)");

        private static readonly HashSet<string> BodyExcludedRoles = new HashSet<string>(
            new[]
            {
                "nationalTitle", "nationalMotto", "nationalMottoSeparator", "partyTitle",
                "organName", "superiorOrganName", "codeNumber", "place", "placeDate",
                "placeAndIssuedDate", "typeName", "subject", "subjectContinuation",
                "officialLetterSubject", "legalBasis", "article", "clause", "point",
                "signerAuthority", "signerPosition", "signerFullName", "recipientSalutation",
                "recipientSalutationInline", "recipientSalutationList", "recipientLabel",
                "recipientList", "appendixLabel", "appendixTitle", "appendixReference",
                "appendixDigitalSignatureInfo", "caption", "tableCaption", "figureCaption"
            }, StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> BooktabsExcludedRoles = new HashSet<string>(
            BodyExcludedRoles.Concat(new[] { "structuralTitle", "partChapterHeading", "sectionHeading" }),
            StringComparer.OrdinalIgnoreCase);

        private readonly HeadingDetector _headingDetector = new HeadingDetector();

        public IReadOnlyList<AnnotationFinding> Scan(LocalScanSnapshot snapshot, LocalRulePack rules,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var roleDetector = new DocumentRoleDetector();
            var blocks = roleDetector.DetectBlocks(snapshot);
            var roles = MergeRoles(blocks);
            var analysis = DerivedAnalysisContext.Create(snapshot, blocks, roles, _headingDetector);
            return ScanDetailed(snapshot, rules, analysis, cancellationToken).Findings;
        }

        public AcademicTypographyScanOutcome ScanDetailed(LocalScanSnapshot snapshot, LocalRulePack rules,
            DerivedAnalysisContext analysis,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            if (analysis == null) throw new ArgumentNullException(nameof(analysis));

            var profile = rules.AcademicTypography;
            if (!profile.Enabled)
                return new AcademicTypographyScanOutcome(Array.Empty<AnnotationFinding>(),
                    Array.Empty<DetectedHeading>(), Array.Empty<string>(), profile);

            var findings = new List<AnnotationFinding>();
            var notEvaluated = new HashSet<string>(StringComparer.Ordinal);
            var paragraphs = snapshot.Paragraphs.OrderBy(p => p.Index).ToArray();
            var paragraphMap = paragraphs.ToDictionary(p => p.Index);
            var headings = analysis.Headings.Where(h => HeadingConfidence(h, paragraphMap) >=
                profile.Thresholds.HeadingConfidenceMinimum).ToArray();

            cancellationToken.ThrowIfCancellationRequested();
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.SectionStyle))
                CheckHeadingStyles(findings, headings, paragraphMap);
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.SectionContinuity))
                CheckHeadingContinuity(findings, headings, analysis, paragraphMap);
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.PaginationKeep))
                CheckKeepWithNext(findings, headings, paragraphMap, paragraphs, analysis, snapshot);

            cancellationToken.ThrowIfCancellationRequested();
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.PaginationWidow))
                CheckWidowControl(findings, notEvaluated, paragraphs, headings, analysis, snapshot,
                    profile.Thresholds.BodyConfidenceMinimum);
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.MathSyntax))
                CheckMathSyntax(findings, notEvaluated, paragraphs, snapshot,
                    profile.Thresholds.MathMinimumSignalCount);

            cancellationToken.ThrowIfCancellationRequested();
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.TableBooktabs))
                CheckBooktabs(findings, notEvaluated, snapshot, analysis);
            if (profile.IsRuleEnabled(AcademicTypographyRuleCodes.CaptionPosition))
                CheckCaptions(findings, notEvaluated, snapshot, paragraphs,
                    profile.Thresholds.CaptionMaxBlankParagraphs);

            return new AcademicTypographyScanOutcome(findings, headings,
                notEvaluated.OrderBy(code => code, StringComparer.Ordinal).ToArray(), profile);
        }

        private static void CheckHeadingStyles(ICollection<AnnotationFinding> findings,
            IEnumerable<DetectedHeading> headings,
            IReadOnlyDictionary<int, LocalParagraphSnapshot> paragraphs)
        {
            foreach (var heading in headings)
            {
                LocalParagraphSnapshot paragraph;
                if (!paragraphs.TryGetValue(heading.ParagraphIndex, out paragraph)) continue;
                var expectedStyleId = -(heading.Level + 1);
                if (paragraph.BuiltInStyleId == expectedStyleId || paragraph.OutlineLevel == heading.Level)
                    continue;
                AddParagraph(findings, AcademicTypographyRuleCodes.SectionStyle, paragraph,
                    "Đề mục “" + DisplayHeading(heading) + "” chưa dùng đúng Heading " + heading.Level + ".",
                    "Gán Heading " + heading.Level + " sau khi kiểm tra Style không làm đổi thể thức pháp lý; " +
                    "khuyến nghị này không phải lỗi NĐ30/HD05.");
            }
        }

        private void CheckHeadingContinuity(ICollection<AnnotationFinding> findings,
            IReadOnlyList<DetectedHeading> headings, DerivedAnalysisContext analysis,
            IReadOnlyDictionary<int, LocalParagraphSnapshot> paragraphMap)
        {
            foreach (var issue in _headingDetector.AnalyzeContinuity(headings,
                new HeadingContinuityOptions(false, analysis.LogicalBlockIdsByParagraphIndex)))
            {
                LocalParagraphSnapshot paragraph;
                if (!paragraphMap.TryGetValue(issue.Heading.ParagraphIndex, out paragraph)) continue;
                AddParagraph(findings, AcademicTypographyRuleCodes.SectionContinuity, paragraph,
                    issue.CurrentIssue, issue.Expected + " Không tự đánh lại số; đây là khuyến nghị biên tập.");
            }
        }

        private static void CheckKeepWithNext(ICollection<AnnotationFinding> findings,
            IEnumerable<DetectedHeading> headings,
            IReadOnlyDictionary<int, LocalParagraphSnapshot> paragraphMap,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs,
            DerivedAnalysisContext analysis, LocalScanSnapshot snapshot)
        {
            foreach (var heading in headings)
            {
                LocalParagraphSnapshot paragraph;
                if (!paragraphMap.TryGetValue(heading.ParagraphIndex, out paragraph) ||
                    paragraph.KeepWithNext != false || snapshot.IntersectsProtectedSpan(paragraph)) continue;
                var next = paragraphs.FirstOrDefault(p => p.Index > paragraph.Index &&
                    !string.IsNullOrWhiteSpace(p.Text) &&
                    string.Equals(analysis.BlockIdOf(p), heading.LogicalBlockId, StringComparison.Ordinal));
                if (next == null || next.IsInTable || snapshot.IntersectsProtectedSpan(next)) continue;
                AddParagraph(findings, AcademicTypographyRuleCodes.PaginationKeep, paragraph,
                    "Đề mục “" + DisplayHeading(heading) + "” chưa bật Keep with next.",
                    "Bật Keep with next để giữ đề mục cùng đoạn nội dung ngay sau; đây là khuyến nghị trình bày, " +
                    "không phải lỗi NĐ30/HD05.");
            }
        }

        private static void CheckWidowControl(ICollection<AnnotationFinding> findings,
            ISet<string> notEvaluated, IEnumerable<LocalParagraphSnapshot> paragraphs,
            IEnumerable<DetectedHeading> headings, DerivedAnalysisContext analysis,
            LocalScanSnapshot snapshot, double confidenceMinimum)
        {
            var headingIndexes = new HashSet<int>(headings.Select(h => h.ParagraphIndex));
            foreach (var paragraph in paragraphs)
            {
                if (!IsBodyCandidate(paragraph, analysis, snapshot, headingIndexes)) continue;
                if (BodyConfidence(paragraph.Text) < confidenceMinimum) continue;
                if (!paragraph.WidowControl.HasValue)
                {
                    notEvaluated.Add(AcademicTypographyRuleCodes.PaginationWidow);
                    continue;
                }
                if (paragraph.WidowControl.Value) continue;
                AddParagraph(findings, AcademicTypographyRuleCodes.PaginationWidow, paragraph,
                    "Đoạn thân bài chưa bật Widow/Orphan control.",
                    "Bật Widow/Orphan control để tránh một dòng đơn lẻ ở đầu hoặc cuối trang; " +
                    "đây là khuyến nghị trình bày.");
            }
        }

        private static bool IsBodyCandidate(LocalParagraphSnapshot paragraph,
            DerivedAnalysisContext analysis, LocalScanSnapshot snapshot, ISet<int> headingIndexes)
        {
            if (!string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.OrdinalIgnoreCase) ||
                paragraph.IsInTable || headingIndexes.Contains(paragraph.Index) ||
                snapshot.IntersectsProtectedSpan(paragraph) || string.IsNullOrWhiteSpace(paragraph.Text) ||
                LeadingListRegex.IsMatch(paragraph.Text) || LegalTextRegex.IsMatch(paragraph.Text) ||
                IsCaption(paragraph)) return false;
            return !BodyExcludedRoles.Contains(analysis.RoleOf(paragraph));
        }

        private static double BodyConfidence(string text)
        {
            var value = (text ?? string.Empty).Trim();
            var words = WordRegex.Matches(value).Cast<Match>().ToArray();
            if (words.Length < 8) return 0d;
            var letterCount = value.Count(char.IsLetter);
            var nonSpaceCount = value.Count(c => !char.IsWhiteSpace(c));
            if (nonSpaceCount == 0 || letterCount / (double)nonSpaceCount < .65d) return 0d;
            var punctuation = value.EndsWith(".", StringComparison.Ordinal) ||
                value.EndsWith("?", StringComparison.Ordinal) || value.EndsWith("!", StringComparison.Ordinal);
            if (!punctuation) return .90d;
            return words.Length >= 12 ? .99d : .96d;
        }

        private static void CheckBooktabs(ICollection<AnnotationFinding> findings,
            ISet<string> notEvaluated, LocalScanSnapshot snapshot, DerivedAnalysisContext analysis)
        {
            foreach (var table in snapshot.Tables)
            {
                var tableParagraphs = snapshot.Paragraphs.Where(p => p.TableIndex == table.Index &&
                    string.Equals(p.StoryType, table.StoryType, StringComparison.Ordinal)).ToArray();
                if (table.IsNested || table.NestingDepth > 1 || table.HasMergedCellsState != false ||
                    table.RowCount < 2 || table.ColumnCount < 2 || table.HeaderRowIndexes.Count == 0 ||
                    tableParagraphs.Length == 0 || tableParagraphs.Any(p =>
                        BooktabsExcludedRoles.Contains(analysis.RoleOf(p)))) continue;

                var required = new[] { table.TopBorder, table.BottomBorder, table.LeftBorder,
                    table.RightBorder, table.InsideVerticalBorder, table.HeaderSeparatorBorder };
                if (required.Any(border => border.State == LocalSnapshotValueState.Unknown))
                {
                    notEvaluated.Add(AcademicTypographyRuleCodes.TableBooktabs);
                    continue;
                }

                var problems = new List<string>();
                if (table.LeftBorder.State == LocalSnapshotValueState.Present ||
                    table.RightBorder.State == LocalSnapshotValueState.Present ||
                    table.InsideVerticalBorder.State == LocalSnapshotValueState.Present)
                    problems.Add("còn viền dọc");
                if (!IsBorderWithin(table.TopBorder, 1d, 1.5d)) problems.Add("viền trên chưa ở mức 1–1,5 pt");
                if (!IsBorderWithin(table.BottomBorder, 1d, 1.5d)) problems.Add("viền dưới chưa ở mức 1–1,5 pt");
                if (!IsBorderWithin(table.HeaderSeparatorBorder, .5d, .75d))
                    problems.Add("đường ngăn hàng tiêu đề chưa ở mức 0,5–0,75 pt");
                if (problems.Count == 0) continue;

                AddParagraph(findings, AcademicTypographyRuleCodes.TableBooktabs, tableParagraphs[0],
                    "Bảng dữ liệu " + table.Index + " chưa theo gợi ý booktabs: " +
                    string.Join(", ", problems) + ".",
                    "Có thể bỏ viền dọc, dùng viền trên/dưới 1–1,5 pt và đường ngăn hàng tiêu đề " +
                    "0,5–0,75 pt; không áp dụng cho bảng bố cục hoặc biểu mẫu pháp lý.");
            }
        }

        private static bool IsBorderWithin(LocalBorderSnapshot border, double minimum, double maximum) =>
            border.State == LocalSnapshotValueState.Present && border.WeightPoints.HasValue &&
            border.WeightPoints.Value >= minimum - .05d && border.WeightPoints.Value <= maximum + .05d;

        private static void CheckCaptions(ICollection<AnnotationFinding> findings,
            ISet<string> notEvaluated, LocalScanSnapshot snapshot,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs, int maximumBlankParagraphs)
        {
            var ordered = paragraphs.Where(p => string.Equals(p.StoryType, "wdMainTextStory",
                StringComparison.OrdinalIgnoreCase)).OrderBy(p => p.Index).ToArray();
            foreach (var table in snapshot.Tables.Where(t => string.Equals(t.StoryType,
                "wdMainTextStory", StringComparison.OrdinalIgnoreCase)))
            {
                var tableParagraphs = ordered.Where(p => p.TableIndex == table.Index).ToArray();
                if (tableParagraphs.Length == 0)
                {
                    notEvaluated.Add(AcademicTypographyRuleCodes.CaptionPosition);
                    continue;
                }
                var below = AdjacentParagraph(ordered, tableParagraphs.Max(p => p.Index), true,
                    maximumBlankParagraphs);
                if (below != null && IsTableCaption(below))
                    AddParagraph(findings, AcademicTypographyRuleCodes.CaptionPosition, below,
                        "Chú thích Bảng đang đặt phía dưới bảng liên kết.",
                        "Đặt chú thích Bảng phía trên bảng; đây là khuyến nghị biên tập, không tự di chuyển nội dung.");
            }

            foreach (var graphic in snapshot.GraphicObjects.Where(g => !g.IsProtected &&
                string.Equals(g.StoryType, "wdMainTextStory", StringComparison.OrdinalIgnoreCase)))
            {
                if (!graphic.AnchorParagraphIndex.HasValue)
                {
                    notEvaluated.Add(AcademicTypographyRuleCodes.CaptionPosition);
                    continue;
                }
                var above = AdjacentParagraph(ordered, graphic.AnchorParagraphIndex.Value, false,
                    maximumBlankParagraphs);
                if (above != null && IsFigureCaption(above))
                    AddParagraph(findings, AcademicTypographyRuleCodes.CaptionPosition, above,
                        "Chú thích Hình đang đặt phía trên hình liên kết.",
                        "Đặt chú thích Hình phía dưới hình; đây là khuyến nghị biên tập, không tự di chuyển nội dung.");
            }
        }

        private static LocalParagraphSnapshot? AdjacentParagraph(
            IReadOnlyList<LocalParagraphSnapshot> paragraphs, int anchorIndex, bool after,
            int maximumBlankParagraphs)
        {
            var candidates = after
                ? paragraphs.Where(p => p.Index > anchorIndex).OrderBy(p => p.Index)
                : paragraphs.Where(p => p.Index < anchorIndex).OrderByDescending(p => p.Index);
            var blanks = 0;
            foreach (var paragraph in candidates)
            {
                if (paragraph.IsInTable) continue;
                if (string.IsNullOrWhiteSpace(paragraph.Text))
                {
                    if (++blanks > maximumBlankParagraphs) return null;
                    continue;
                }
                return paragraph;
            }
            return null;
        }

        private static void CheckMathSyntax(ICollection<AnnotationFinding> findings,
            ISet<string> notEvaluated, IEnumerable<LocalParagraphSnapshot> paragraphs,
            LocalScanSnapshot snapshot, int minimumSignalCount)
        {
            foreach (var paragraph in paragraphs)
            {
                if (string.IsNullOrEmpty(paragraph.Text) || paragraph.Text.IndexOf('$') < 0) continue;
                if (snapshot.IntersectsProtectedSpan(paragraph) || paragraph.IsInTable) continue;
                if (!paragraph.HasField.HasValue || !paragraph.HasMathObject.HasValue ||
                    !paragraph.HasHyperlink.HasValue || !paragraph.HasContentControl.HasValue)
                {
                    notEvaluated.Add(AcademicTypographyRuleCodes.MathSyntax);
                    continue;
                }
                if (paragraph.HasField.Value || paragraph.HasMathObject.Value ||
                    paragraph.HasHyperlink.Value || paragraph.HasContentControl.Value) continue;

                bool ambiguous;
                var expressions = ParseMathExpressions(paragraph.Text, out ambiguous);
                if (ambiguous) continue;
                foreach (var expression in expressions)
                {
                    if (MathSignalCount(expression.Content) < minimumSignalCount ||
                        LooksLikeCurrency(expression.Content)) continue;
                    AddSpan(findings, AcademicTypographyRuleCodes.MathSyntax, paragraph,
                        expression.Start, expression.Length,
                        "Còn cú pháp toán LaTeX thô “" + expression.Value + "”.",
                        "Chuyển biểu thức sang Word Equation (OMath) sau khi kiểm tra nội dung; " +
                        "không tự chuyển đổi cú pháp mơ hồ.");
                }
            }
        }

        private static IReadOnlyList<MathExpression> ParseMathExpressions(string text, out bool ambiguous)
        {
            ambiguous = false;
            var result = new List<MathExpression>();
            var index = 0;
            while (index < text.Length)
            {
                if (text[index] != '$' || IsEscaped(text, index)) { index++; continue; }
                var delimiterLength = index + 1 < text.Length && text[index + 1] == '$' ? 2 : 1;
                var end = FindClosingDelimiter(text, index + delimiterLength, delimiterLength);
                if (end < 0) { ambiguous = true; return Array.Empty<MathExpression>(); }
                var content = text.Substring(index + delimiterLength, end - index - delimiterLength);
                if (content.IndexOf('$') >= 0 || string.IsNullOrWhiteSpace(content))
                {
                    ambiguous = true;
                    return Array.Empty<MathExpression>();
                }
                var length = end + delimiterLength - index;
                result.Add(new MathExpression(index, length, content, text.Substring(index, length)));
                index += length;
            }
            return result;
        }

        private static int FindClosingDelimiter(string text, int start, int delimiterLength)
        {
            for (var index = start; index < text.Length; index++)
            {
                if (text[index] != '$' || IsEscaped(text, index)) continue;
                var currentLength = index + 1 < text.Length && text[index + 1] == '$' ? 2 : 1;
                return currentLength == delimiterLength ? index : -1;
            }
            return -1;
        }

        private static bool IsEscaped(string text, int index)
        {
            var slashes = 0;
            for (var current = index - 1; current >= 0 && text[current] == '\\'; current--) slashes++;
            return slashes % 2 == 1;
        }

        private static int MathSignalCount(string content)
        {
            var count = MathOperatorRegex.Matches(content).Count;
            if (MathCommandRegex.IsMatch(content)) count++;
                if (content.Any(char.IsLetter) && content.Any(char.IsDigit) &&
                    (content.IndexOf('=') >= 0 || content.IndexOf('^') >= 0 ||
                     content.IndexOf('_') >= 0 || content.IndexOf('\\') >= 0)) count++;
            return count;
        }

        private static bool LooksLikeCurrency(string content)
        {
            var value = content.Trim();
            decimal amount;
            return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out amount) ||
                Regex.IsMatch(value, @"^\d+(?:[.,]\d+)?\s*(?:USD|VND|đ|₫)?$",
                    RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        }

        private static double HeadingConfidence(DetectedHeading heading,
            IReadOnlyDictionary<int, LocalParagraphSnapshot> paragraphs)
        {
            LocalParagraphSnapshot paragraph;
            if (!paragraphs.TryGetValue(heading.ParagraphIndex, out paragraph)) return 0d;
            if (paragraph.BuiltInStyleId.HasValue ||
                (paragraph.OutlineLevel.HasValue && paragraph.OutlineLevel.Value >= 1 &&
                 paragraph.OutlineLevel.Value <= 9)) return .99d;
            if (heading.IsBold && heading.Kind != HeadingNumberingKind.Decimal) return .96d;
            if (heading.IsBold && heading.NumberParts.Count > 1) return .95d;
            if (heading.TitleText.Any(char.IsLetter) &&
                heading.TitleText.Where(char.IsLetter).All(char.IsUpper)) return .93d;
            return 0d;
        }

        private static bool IsCaption(LocalParagraphSnapshot paragraph) =>
            IsTableCaption(paragraph) || IsFigureCaption(paragraph);

        private static bool IsTableCaption(LocalParagraphSnapshot paragraph) =>
            string.Equals(paragraph.CaptionKind, "Table", StringComparison.OrdinalIgnoreCase) ||
            TableCaptionRegex.IsMatch(paragraph.Text ?? string.Empty);

        private static bool IsFigureCaption(LocalParagraphSnapshot paragraph) =>
            string.Equals(paragraph.CaptionKind, "Figure", StringComparison.OrdinalIgnoreCase) ||
            FigureCaptionRegex.IsMatch(paragraph.Text ?? string.Empty);

        private static Dictionary<int, string> MergeRoles(IEnumerable<LogicalDocumentBlock> blocks)
        {
            var result = new Dictionary<int, string>();
            foreach (var block in blocks)
                foreach (var role in block.Roles) result[role.Key] = role.Value;
            return result;
        }

        private static string DisplayHeading(DetectedHeading heading) =>
            string.IsNullOrWhiteSpace(heading.NumberText)
                ? heading.TitleText
                : heading.NumberText + " " + heading.TitleText;

        private static void AddParagraph(ICollection<AnnotationFinding> findings, string ruleCode,
            LocalParagraphSnapshot paragraph, string current, string expected)
        {
            findings.Add(new AnnotationFinding(ruleCode + "-P" + paragraph.Index, ruleCode, "Suggestion",
                current, expected, "Khuyến nghị LaTeX/Typst",
                Anchor(AnnotationAnchorKind.Paragraph, paragraph, null, null, string.Empty)));
        }

        private static void AddSpan(ICollection<AnnotationFinding> findings, string ruleCode,
            LocalParagraphSnapshot paragraph, int offset, int length, string current, string expected)
        {
            findings.Add(new AnnotationFinding(ruleCode + "-P" + paragraph.Index + "-C" + offset,
                ruleCode, "Suggestion", current, expected, "Khuyến nghị LaTeX/Typst",
                Anchor(AnnotationAnchorKind.TextSpan, paragraph, offset, length,
                    paragraph.Text.Substring(offset, length))));
        }

        private static AnnotationAnchor Anchor(AnnotationAnchorKind kind,
            LocalParagraphSnapshot paragraph, int? offset, int? length, string expectedText) =>
            new AnnotationAnchor(kind, paragraph.StoryType, paragraph.Index, offset, length,
                expectedText, paragraph.SectionIndex, paragraph.TableIndex,
                paragraph.RowIndex, paragraph.CellIndex);

        private static Regex Rx(string pattern, bool ignoreCase = false) => new Regex(pattern,
            RegexOptions.CultureInvariant | RegexOptions.Compiled |
            (ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None),
            TimeSpan.FromMilliseconds(200));

        private sealed class MathExpression
        {
            public MathExpression(int start, int length, string content, string value)
            { Start = start; Length = length; Content = content; Value = value; }
            public int Start { get; }
            public int Length { get; }
            public string Content { get; }
            public string Value { get; }
        }
    }
}
