using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ChuanHoa.Client.Core.Annotations
{
    public sealed class AnnotationPlanner
    {
        public AnnotationPlan CreatePlan(
            string lane,
            string scanId,
            string expectedDocumentFingerprint,
            long expectedRevision,
            AnnotationDocumentSnapshot document,
            IReadOnlyList<AnnotationFinding> findings) =>
            CreatePlan(lane, scanId, expectedDocumentFingerprint, expectedRevision, document,
                findings, AnnotationPlanOptions.Unlimited);

        public AnnotationPlan CreatePlan(
            string lane,
            string scanId,
            string expectedDocumentFingerprint,
            long expectedRevision,
            AnnotationDocumentSnapshot document,
            IReadOnlyList<AnnotationFinding> findings,
            AnnotationPlanOptions options)
        {
            if (string.IsNullOrWhiteSpace(lane))
            {
                throw new ArgumentException("An annotation lane is required.", nameof(lane));
            }
            if (string.IsNullOrWhiteSpace(scanId))
            {
                throw new ArgumentException("A scan id is required.", nameof(scanId));
            }
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }
            if (findings == null)
            {
                throw new ArgumentNullException(nameof(findings));
            }
            options = options ?? AnnotationPlanOptions.Unlimited;

            var unresolved = new List<UnresolvedAnnotation>();

            if (!string.Equals(expectedDocumentFingerprint, document.DocumentFingerprint, StringComparison.Ordinal))
            {
                AddAllUnresolved(findings, AnnotationResolutionCode.DocumentMismatch, unresolved);
                return new AnnotationPlan(lane, scanId, Array.Empty<AnnotationCommentInstruction>(),
                    Array.Empty<AnnotationVisualInstruction>(), unresolved);
            }
            if (expectedRevision != document.Revision)
            {
                AddAllUnresolved(findings, AnnotationResolutionCode.RevisionMismatch, unresolved);
                return new AnnotationPlan(lane, scanId, Array.Empty<AnnotationCommentInstruction>(),
                    Array.Empty<AnnotationVisualInstruction>(), unresolved);
            }

            // Build every lookup once. The previous implementation filtered and
            // sorted all paragraphs again for each finding (O(findings × paragraphs
            // × log paragraphs)), which took tens of seconds on long documents.
            var index = new DocumentIndex(document);
            var resolvedFindings = new List<ResolvedFinding>();
            foreach (var group in findings.GroupBy(item => item.FindingId, StringComparer.Ordinal))
            {
                var finding = group.First();
                if (group.Skip(1).Any(item => !AreEquivalent(finding, item)))
                {
                    unresolved.Add(new UnresolvedAnnotation(
                        finding.FindingId,
                        AnnotationResolutionCode.ConflictingDuplicateFinding));
                    continue;
                }

                if (!AnnotationPresentationPolicy.ShouldAnnotate(finding))
                {
                    continue;
                }

                ResolvedAnchor? resolved;
                AnnotationResolutionCode code;
                if (!TryResolve(index, finding.Anchor, out resolved, out code))
                {
                    unresolved.Add(new UnresolvedAnnotation(finding.FindingId, code));
                    continue;
                }

                resolvedFindings.Add(new ResolvedFinding(finding, resolved!));
            }

            return options.GroupByParagraph || options.MaximumComments > 0
                ? CreateBudgetedPlan(lane, scanId, resolvedFindings, unresolved, options)
                : CreateUnlimitedPlan(lane, scanId, resolvedFindings, unresolved);
        }

        private static AnnotationPlan CreateUnlimitedPlan(string lane, string scanId,
            IReadOnlyList<ResolvedFinding> resolvedFindings, List<UnresolvedAnnotation> unresolved)
        {
            var comments = new List<AnnotationCommentInstruction>(resolvedFindings.Count);
            var visualRanges = new List<AnnotationVisualInstruction>();
            foreach (var item in resolvedFindings)
            {
                comments.Add(new AnnotationCommentInstruction(
                    item.Finding.FindingId,
                    item.Anchor.StoryType,
                    item.Anchor.SectionIndex,
                    item.Anchor.Start,
                    Math.Max(1, item.Anchor.Length),
                    Marker(lane, item.Finding.FindingId),
                    BuildComment(item.Finding)));
                AddVisual(visualRanges, item, item.Finding.FindingId);
            }
            return new AnnotationPlan(lane, scanId, comments, MergeVisualRanges(visualRanges),
                unresolved, Array.Empty<AnnotationFinding>(), comments.Count);
        }

        private static AnnotationPlan CreateBudgetedPlan(string lane, string scanId,
            IReadOnlyList<ResolvedFinding> resolvedFindings, List<UnresolvedAnnotation> unresolved,
            AnnotationPlanOptions options)
        {
            var omitted = new List<AnnotationFinding>();
            var ordered = resolvedFindings
                .OrderBy(item => item.Anchor.StoryOrder)
                .ThenBy(item => item.Anchor.Start)
                .ThenBy(item => item.Finding.FindingId, StringComparer.Ordinal)
                .ToList();

            // A rule code that occurs in many paragraphs (for example the body font
            // of a whole technical document) becomes one summary comment at its
            // first occurrence instead of hundreds of identical comments.
            var summaryRules = new Dictionary<string, int>(StringComparer.Ordinal);
            if (options.RepeatedRuleSummaryThreshold > 0)
            {
                foreach (var rule in ordered.GroupBy(item => item.Finding.RuleCode, StringComparer.Ordinal))
                {
                    var paragraphCount = rule.Select(item => item.Anchor.ParagraphKey)
                        .Distinct(StringComparer.Ordinal).Count();
                    if (paragraphCount > options.RepeatedRuleSummaryThreshold)
                        summaryRules[rule.Key] = paragraphCount;
                }
            }

            var units = new List<CommentUnit>();
            var unitsByParagraph = new Dictionary<string, CommentUnit>(StringComparer.Ordinal);
            var summarizedRules = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in ordered)
            {
                int repeatedParagraphs;
                if (summaryRules.TryGetValue(item.Finding.RuleCode, out repeatedParagraphs))
                {
                    if (!summarizedRules.Add(item.Finding.RuleCode))
                    {
                        omitted.Add(item.Finding);
                        continue;
                    }
                    units.Add(new CommentUnit(item, repeatedParagraphs));
                    continue;
                }

                if (!options.GroupByParagraph)
                {
                    units.Add(new CommentUnit(item, 0));
                    continue;
                }

                CommentUnit unit;
                if (unitsByParagraph.TryGetValue(item.Anchor.ParagraphKey, out unit))
                {
                    unit.Add(item);
                    continue;
                }
                unit = new CommentUnit(item, 0);
                unitsByParagraph[item.Anchor.ParagraphKey] = unit;
                units.Add(unit);
            }

            var comments = new List<AnnotationCommentInstruction>();
            var visualRanges = new List<AnnotationVisualInstruction>();
            var annotated = 0;
            foreach (var unit in units
                .OrderBy(item => item.Primary.Anchor.StoryOrder)
                .ThenBy(item => item.Start))
            {
                if (options.MaximumComments > 0 && comments.Count >= options.MaximumComments)
                {
                    omitted.AddRange(unit.Items.Select(item => item.Finding));
                    continue;
                }

                var primary = unit.Primary;
                comments.Add(new AnnotationCommentInstruction(
                    primary.Finding.FindingId,
                    primary.Anchor.StoryType,
                    primary.Anchor.SectionIndex,
                    unit.Start,
                    Math.Max(1, unit.End - unit.Start),
                    Marker(lane, primary.Finding.FindingId),
                    unit.BuildText()));
                annotated += unit.Items.Count;
                foreach (var item in unit.Items)
                    AddVisual(visualRanges, item, primary.Finding.FindingId);
            }

            return new AnnotationPlan(lane, scanId, comments, MergeVisualRanges(visualRanges),
                unresolved, omitted, annotated);
        }

        private static void AddVisual(ICollection<AnnotationVisualInstruction> visualRanges,
            ResolvedFinding item, string ownerFindingId)
        {
            if (item.Anchor.ShouldMarkRed && item.Anchor.Length > 0 &&
                AnnotationPresentationPolicy.ShouldMarkRed(item.Finding))
            {
                visualRanges.Add(new AnnotationVisualInstruction(
                    ownerFindingId,
                    item.Anchor.StoryType,
                    item.Anchor.SectionIndex,
                    item.Anchor.Start,
                    item.Anchor.Length));
            }
        }

        private static string Marker(string lane, string findingId) =>
            "[CHUẨN HÓA:" + lane.ToUpperInvariant() + ":" + findingId + "]";

        private static bool TryResolve(
            DocumentIndex index,
            AnnotationAnchor anchor,
            out ResolvedAnchor? resolved,
            out AnnotationResolutionCode code)
        {
            resolved = null;
            code = AnnotationResolutionCode.UnsupportedAnchor;
            var story = index.Story(anchor.StoryType);
            if (story == null)
            {
                code = AnnotationResolutionCode.StoryNotFound;
                return false;
            }

            AnnotationParagraphSnapshot? paragraph;
            var shouldMarkRed = true;
            switch (anchor.Kind)
            {
                case AnnotationAnchorKind.TextSpan:
                case AnnotationAnchorKind.Paragraph:
                    paragraph = story.FindParagraph(anchor);
                    if (paragraph == null)
                    {
                        code = AnnotationResolutionCode.ParagraphNotFound;
                        return false;
                    }
                    break;
                case AnnotationAnchorKind.Section:
                    paragraph = story.FirstInSection(anchor.SectionIndex);
                    shouldMarkRed = false;
                    if (paragraph == null)
                    {
                        code = AnnotationResolutionCode.SectionNotFound;
                        return false;
                    }
                    break;
                case AnnotationAnchorKind.Document:
                    paragraph = story.First;
                    shouldMarkRed = false;
                    break;
                default:
                    return false;
            }

            var resolvedParagraph = paragraph!;

            var printableLength = TrimParagraphTerminator(resolvedParagraph.Text).Length;
            var offset = 0;
            var length = Math.Min(1, printableLength);
            if (anchor.Kind == AnnotationAnchorKind.TextSpan)
            {
                if (!anchor.StartOffset.HasValue || !anchor.Length.HasValue ||
                    anchor.StartOffset.Value < 0 || anchor.Length.Value <= 0 ||
                    anchor.StartOffset.Value > printableLength - anchor.Length.Value)
                {
                    code = AnnotationResolutionCode.InvalidRange;
                    return false;
                }
                offset = anchor.StartOffset.Value;
                length = anchor.Length.Value;
                if (string.IsNullOrEmpty(anchor.ExpectedText) || anchor.ExpectedText.Length != length ||
                    string.CompareOrdinal(resolvedParagraph.Text, offset, anchor.ExpectedText, 0, length) != 0)
                {
                    code = AnnotationResolutionCode.ExpectedTextMismatch;
                    return false;
                }
            }
            else if (anchor.Kind == AnnotationAnchorKind.Paragraph)
            {
                length = printableLength;
                if (length <= 0)
                {
                    code = AnnotationResolutionCode.InvalidRange;
                    return false;
                }
            }

            var absoluteStart = resolvedParagraph.AbsoluteStart + offset;
            if (story.IntersectsProtectedSpan(absoluteStart, Math.Max(1, length)))
            {
                code = AnnotationResolutionCode.ProtectedSpan;
                return false;
            }

            resolved = new ResolvedAnchor(
                anchor.StoryType,
                story.Order,
                resolvedParagraph.SectionIndex,
                absoluteStart,
                length,
                shouldMarkRed,
                story.Order.ToString(CultureInfo.InvariantCulture) + ":" +
                    resolvedParagraph.AbsoluteStart.ToString(CultureInfo.InvariantCulture));
            code = AnnotationResolutionCode.Resolved;
            return true;
        }

        private static IReadOnlyList<AnnotationVisualInstruction> MergeVisualRanges(
            IEnumerable<AnnotationVisualInstruction> ranges)
        {
            // Linear merge per (finding, story, section) after one sort. The old
            // implementation searched the whole merged list with LastOrDefault for
            // every range.
            var merged = new List<AnnotationVisualInstruction>();
            foreach (var group in ranges.GroupBy(
                item => item.FindingId + "\u001f" + item.StoryType + "\u001f" +
                    item.SectionIndex.ToString(CultureInfo.InvariantCulture),
                StringComparer.Ordinal))
            {
                AnnotationVisualInstruction? previous = null;
                foreach (var current in group.OrderBy(item => item.Start).ThenBy(item => item.Length))
                {
                    if (previous == null || current.Start > previous.Start + previous.Length)
                    {
                        if (previous != null) merged.Add(previous);
                        previous = current;
                        continue;
                    }

                    var end = Math.Max(previous.Start + previous.Length, current.Start + current.Length);
                    previous = new AnnotationVisualInstruction(
                        previous.FindingId,
                        previous.StoryType,
                        previous.SectionIndex,
                        previous.Start,
                        end - previous.Start);
                }
                if (previous != null) merged.Add(previous);
            }
            return merged;
        }

        private static string BuildComment(AnnotationFinding finding)
        {
            var builder = new StringBuilder();
            builder.Append("Hiện tại: ").AppendLine(AnnotationPresentationPolicy.CurrentIssue(finding));
            builder.Append("Yêu cầu đúng: ").Append(finding.Expected.Trim());
            return builder.ToString();
        }

        private static bool AreEquivalent(AnnotationFinding left, AnnotationFinding right)
        {
            return string.Equals(left.RuleCode, right.RuleCode, StringComparison.Ordinal) &&
                string.Equals(left.Severity, right.Severity, StringComparison.Ordinal) &&
                left.SeverityLevel == right.SeverityLevel &&
                string.Equals(left.CurrentIssue, right.CurrentIssue, StringComparison.Ordinal) &&
                string.Equals(left.Expected, right.Expected, StringComparison.Ordinal) &&
                string.Equals(left.Citation, right.Citation, StringComparison.Ordinal) &&
                left.SourceFamily == right.SourceFamily &&
                left.Anchor.Kind == right.Anchor.Kind &&
                string.Equals(left.Anchor.StoryType, right.Anchor.StoryType, StringComparison.Ordinal) &&
                left.Anchor.ParagraphIndex == right.Anchor.ParagraphIndex &&
                left.Anchor.StartOffset == right.Anchor.StartOffset &&
                left.Anchor.Length == right.Anchor.Length &&
                string.Equals(left.Anchor.ExpectedText, right.Anchor.ExpectedText, StringComparison.Ordinal) &&
                left.Anchor.SectionIndex == right.Anchor.SectionIndex &&
                left.Anchor.TableIndex == right.Anchor.TableIndex &&
                left.Anchor.RowIndex == right.Anchor.RowIndex &&
                left.Anchor.CellIndex == right.Anchor.CellIndex;
        }

        private static string TrimParagraphTerminator(string text)
        {
            return (text ?? string.Empty).TrimEnd('\r', '\a');
        }

        private static void AddAllUnresolved(
            IEnumerable<AnnotationFinding> findings,
            AnnotationResolutionCode code,
            ICollection<UnresolvedAnnotation> unresolved)
        {
            foreach (var finding in findings)
            {
                unresolved.Add(new UnresolvedAnnotation(finding.FindingId, code));
            }
        }

        /// <summary>All lookups the planner needs, built once per plan.</summary>
        private sealed class DocumentIndex
        {
            private readonly Dictionary<string, StoryIndex> _stories =
                new Dictionary<string, StoryIndex>(StringComparer.Ordinal);

            public DocumentIndex(AnnotationDocumentSnapshot document)
            {
                var order = 0;
                foreach (var story in document.Paragraphs.GroupBy(item => item.StoryType, StringComparer.Ordinal))
                {
                    var spans = document.ProtectedSpans
                        .Where(span => string.Equals(span.StoryType, story.Key, StringComparison.Ordinal));
                    _stories[story.Key] = new StoryIndex(order++, story, spans);
                }
            }

            public StoryIndex? Story(string storyType)
            {
                StoryIndex story;
                return _stories.TryGetValue(storyType ?? string.Empty, out story) ? story : null;
            }
        }

        private sealed class StoryIndex
        {
            private readonly Dictionary<int, List<AnnotationParagraphSnapshot>> _byIndex =
                new Dictionary<int, List<AnnotationParagraphSnapshot>>();
            private readonly Dictionary<int, AnnotationParagraphSnapshot> _firstInSection =
                new Dictionary<int, AnnotationParagraphSnapshot>();
            private readonly int[] _spanStarts;
            private readonly int[] _spanMaxEnds;

            public StoryIndex(int order, IEnumerable<AnnotationParagraphSnapshot> paragraphs,
                IEnumerable<AnnotationProtectedSpan> protectedSpans)
            {
                Order = order;
                var sorted = paragraphs.OrderBy(item => item.AbsoluteStart).ToList();
                First = sorted[0];
                foreach (var paragraph in sorted)
                {
                    List<AnnotationParagraphSnapshot> list;
                    if (!_byIndex.TryGetValue(paragraph.ParagraphIndex, out list))
                    {
                        list = new List<AnnotationParagraphSnapshot>(1);
                        _byIndex[paragraph.ParagraphIndex] = list;
                    }
                    list.Add(paragraph);
                    if (!_firstInSection.ContainsKey(paragraph.SectionIndex))
                        _firstInSection[paragraph.SectionIndex] = paragraph;
                }

                var spans = protectedSpans.OrderBy(span => span.Start).ToArray();
                _spanStarts = new int[spans.Length];
                _spanMaxEnds = new int[spans.Length];
                var maximumEnd = int.MinValue;
                for (var i = 0; i < spans.Length; i++)
                {
                    _spanStarts[i] = spans[i].Start;
                    maximumEnd = Math.Max(maximumEnd, spans[i].Start + spans[i].Length);
                    _spanMaxEnds[i] = maximumEnd;
                }
            }

            public int Order { get; }
            public AnnotationParagraphSnapshot First { get; }

            public AnnotationParagraphSnapshot? FirstInSection(int? sectionIndex)
            {
                if (!sectionIndex.HasValue) return null;
                AnnotationParagraphSnapshot paragraph;
                return _firstInSection.TryGetValue(sectionIndex.Value, out paragraph) ? paragraph : null;
            }

            public AnnotationParagraphSnapshot? FindParagraph(AnnotationAnchor anchor)
            {
                if (!anchor.ParagraphIndex.HasValue) return null;
                List<AnnotationParagraphSnapshot> candidates;
                if (!_byIndex.TryGetValue(anchor.ParagraphIndex.Value, out candidates)) return null;

                AnnotationParagraphSnapshot? match = null;
                var count = 0;
                foreach (var item in candidates)
                {
                    if ((!anchor.SectionIndex.HasValue || item.SectionIndex == anchor.SectionIndex.Value) &&
                        (!anchor.TableIndex.HasValue || item.TableIndex == anchor.TableIndex) &&
                        (!anchor.RowIndex.HasValue || item.RowIndex == anchor.RowIndex) &&
                        (!anchor.CellIndex.HasValue || item.CellIndex == anchor.CellIndex))
                    {
                        match = item;
                        if (++count > 1) break;
                    }
                }
                if (count == 1) return match;
                return candidates.Count == 1 ? candidates[0] : null;
            }

            public bool IntersectsProtectedSpan(int start, int length)
            {
                if (_spanStarts.Length == 0) return false;
                var end = start + length;
                // Last span whose start is before the queried end.
                var low = 0;
                var high = _spanStarts.Length - 1;
                var last = -1;
                while (low <= high)
                {
                    var middle = low + (high - low) / 2;
                    if (_spanStarts[middle] < end)
                    {
                        last = middle;
                        low = middle + 1;
                    }
                    else high = middle - 1;
                }
                return last >= 0 && _spanMaxEnds[last] > start;
            }
        }

        private sealed class ResolvedFinding
        {
            public ResolvedFinding(AnnotationFinding finding, ResolvedAnchor anchor)
            {
                Finding = finding;
                Anchor = anchor;
            }

            public AnnotationFinding Finding { get; }
            public ResolvedAnchor Anchor { get; }
        }

        private sealed class CommentUnit
        {
            private readonly int _repeatedParagraphs;

            public CommentUnit(ResolvedFinding primary, int repeatedParagraphs)
            {
                Primary = primary;
                _repeatedParagraphs = repeatedParagraphs;
                Items.Add(primary);
                Start = primary.Anchor.Start;
                End = primary.Anchor.Start + Math.Max(1, primary.Anchor.Length);
            }

            public ResolvedFinding Primary { get; }
            public List<ResolvedFinding> Items { get; } = new List<ResolvedFinding>();
            public int Start { get; private set; }
            public int End { get; private set; }

            public void Add(ResolvedFinding item)
            {
                Items.Add(item);
                Start = Math.Min(Start, item.Anchor.Start);
                End = Math.Max(End, item.Anchor.Start + Math.Max(1, item.Anchor.Length));
            }

            public string BuildText()
            {
                if (_repeatedParagraphs > 0)
                {
                    var builder = new StringBuilder();
                    builder.Append("Lỗi lặp lại ở ")
                        .Append(_repeatedParagraphs.ToString(CultureInfo.InvariantCulture))
                        .Append(" đoạn (chỉ đánh dấu lần đầu; danh sách đầy đủ trong báo cáo).")
                        .AppendLine();
                    builder.Append(BuildComment(Primary.Finding));
                    return builder.ToString();
                }
                if (Items.Count == 1) return BuildComment(Primary.Finding);

                var grouped = new StringBuilder();
                grouped.Append("Đoạn này có ")
                    .Append(Items.Count.ToString(CultureInfo.InvariantCulture))
                    .Append(" lỗi:");
                foreach (var item in Items)
                {
                    grouped.AppendLine();
                    grouped.Append("• ").Append(AnnotationPresentationPolicy.CurrentIssue(item.Finding))
                        .Append(" → ").Append(item.Finding.Expected.Trim());
                }
                return grouped.ToString();
            }
        }

        private sealed class ResolvedAnchor
        {
            public ResolvedAnchor(
                string storyType,
                int storyOrder,
                int sectionIndex,
                int start,
                int length,
                bool shouldMarkRed,
                string paragraphKey)
            {
                StoryType = storyType;
                StoryOrder = storyOrder;
                SectionIndex = sectionIndex;
                Start = start;
                Length = length;
                ShouldMarkRed = shouldMarkRed;
                ParagraphKey = paragraphKey;
            }

            public string StoryType { get; }
            public int StoryOrder { get; }
            public int SectionIndex { get; }
            public int Start { get; }
            public int Length { get; }
            public bool ShouldMarkRed { get; }
            public string ParagraphKey { get; }
        }
    }
}
