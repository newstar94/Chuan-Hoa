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
            IReadOnlyList<AnnotationFinding> findings)
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

            var comments = new List<AnnotationCommentInstruction>();
            var visualRanges = new List<AnnotationVisualInstruction>();
            var unresolved = new List<UnresolvedAnnotation>();

            if (!string.Equals(expectedDocumentFingerprint, document.DocumentFingerprint, StringComparison.Ordinal))
            {
                AddAllUnresolved(findings, AnnotationResolutionCode.DocumentMismatch, unresolved);
                return new AnnotationPlan(lane, scanId, comments, visualRanges, unresolved);
            }
            if (expectedRevision != document.Revision)
            {
                AddAllUnresolved(findings, AnnotationResolutionCode.RevisionMismatch, unresolved);
                return new AnnotationPlan(lane, scanId, comments, visualRanges, unresolved);
            }

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
                if (!TryResolve(document, finding.Anchor, out resolved, out code))
                {
                    unresolved.Add(new UnresolvedAnnotation(finding.FindingId, code));
                    continue;
                }

                var resolvedAnchor = resolved!;

                var marker = "[CHUẨN HÓA:" + lane.ToUpperInvariant() + ":" + finding.FindingId + "]";
                comments.Add(new AnnotationCommentInstruction(
                    finding.FindingId,
                    resolvedAnchor.StoryType,
                    resolvedAnchor.SectionIndex,
                    resolvedAnchor.Start,
                    Math.Max(1, resolvedAnchor.Length),
                    marker,
                    BuildComment(finding)));
                if (resolvedAnchor.ShouldMarkRed && resolvedAnchor.Length > 0 &&
                    AnnotationPresentationPolicy.ShouldMarkRed(finding))
                {
                    visualRanges.Add(new AnnotationVisualInstruction(
                        finding.FindingId,
                        resolvedAnchor.StoryType,
                        resolvedAnchor.SectionIndex,
                        resolvedAnchor.Start,
                        resolvedAnchor.Length));
                }
            }

            return new AnnotationPlan(
                lane,
                scanId,
                comments,
                MergeVisualRanges(visualRanges),
                unresolved);
        }

        private static bool TryResolve(
            AnnotationDocumentSnapshot document,
            AnnotationAnchor anchor,
            out ResolvedAnchor? resolved,
            out AnnotationResolutionCode code)
        {
            resolved = null;
            code = AnnotationResolutionCode.UnsupportedAnchor;
            var storyParagraphs = document.Paragraphs
                .Where(item => string.Equals(item.StoryType, anchor.StoryType, StringComparison.Ordinal))
                .OrderBy(item => item.AbsoluteStart)
                .ToList();
            if (storyParagraphs.Count == 0)
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
                    paragraph = FindParagraph(storyParagraphs, anchor);
                    if (paragraph == null)
                    {
                        code = AnnotationResolutionCode.ParagraphNotFound;
                        return false;
                    }
                    break;
                case AnnotationAnchorKind.Section:
                    paragraph = storyParagraphs.FirstOrDefault(item => item.SectionIndex == anchor.SectionIndex);
                    shouldMarkRed = false;
                    if (paragraph == null)
                    {
                        code = AnnotationResolutionCode.SectionNotFound;
                        return false;
                    }
                    break;
                case AnnotationAnchorKind.Document:
                    paragraph = storyParagraphs[0];
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
                    !string.Equals(resolvedParagraph.Text.Substring(offset, length), anchor.ExpectedText, StringComparison.Ordinal))
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
            if (document.ProtectedSpans.Any(span =>
                string.Equals(span.StoryType, anchor.StoryType, StringComparison.Ordinal) &&
                Intersects(absoluteStart, Math.Max(1, length), span.Start, span.Length)))
            {
                code = AnnotationResolutionCode.ProtectedSpan;
                return false;
            }

            resolved = new ResolvedAnchor(
                anchor.StoryType,
                resolvedParagraph.SectionIndex,
                absoluteStart,
                length,
                shouldMarkRed);
            code = AnnotationResolutionCode.Resolved;
            return true;
        }

        private static AnnotationParagraphSnapshot? FindParagraph(
            IEnumerable<AnnotationParagraphSnapshot> paragraphs,
            AnnotationAnchor anchor)
        {
            if (!anchor.ParagraphIndex.HasValue)
            {
                return null;
            }

            var matches = paragraphs.Where(item =>
                item.ParagraphIndex == anchor.ParagraphIndex.Value &&
                (!anchor.SectionIndex.HasValue || item.SectionIndex == anchor.SectionIndex.Value) &&
                (!anchor.TableIndex.HasValue || item.TableIndex == anchor.TableIndex) &&
                (!anchor.RowIndex.HasValue || item.RowIndex == anchor.RowIndex) &&
                (!anchor.CellIndex.HasValue || item.CellIndex == anchor.CellIndex)).Take(2).ToList();
            if (matches.Count == 1) return matches[0];

            var fallback = paragraphs.Where(item => item.ParagraphIndex == anchor.ParagraphIndex.Value).Take(2).ToList();
            return fallback.Count == 1 ? fallback[0] : null;
        }

        private static IReadOnlyList<AnnotationVisualInstruction> MergeVisualRanges(
            IEnumerable<AnnotationVisualInstruction> ranges)
        {
            var merged = new List<AnnotationVisualInstruction>();
            foreach (var group in ranges.GroupBy(
                item => item.FindingId + "\u001f" + item.StoryType + "\u001f" +
                    item.SectionIndex.ToString(CultureInfo.InvariantCulture),
                StringComparer.Ordinal))
            {
                foreach (var current in group.OrderBy(item => item.Start).ThenBy(item => item.Length))
                {
                    var previous = merged.LastOrDefault(item =>
                        string.Equals(item.FindingId, current.FindingId, StringComparison.Ordinal) &&
                        string.Equals(item.StoryType, current.StoryType, StringComparison.Ordinal) &&
                        item.SectionIndex == current.SectionIndex);
                    if (previous == null || current.Start > previous.Start + previous.Length)
                    {
                        merged.Add(current);
                        continue;
                    }

                    var end = Math.Max(previous.Start + previous.Length, current.Start + current.Length);
                    merged[merged.Count - 1] = new AnnotationVisualInstruction(
                        previous.FindingId,
                        previous.StoryType,
                        previous.SectionIndex,
                        previous.Start,
                        end - previous.Start);
                }
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

        private static bool Intersects(int firstStart, int firstLength, int secondStart, int secondLength)
        {
            return firstStart < secondStart + secondLength && secondStart < firstStart + firstLength;
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

        private sealed class ResolvedAnchor
        {
            public ResolvedAnchor(
                string storyType,
                int sectionIndex,
                int start,
                int length,
                bool shouldMarkRed)
            {
                StoryType = storyType;
                SectionIndex = sectionIndex;
                Start = start;
                Length = length;
                ShouldMarkRed = shouldMarkRed;
            }

            public string StoryType { get; }
            public int SectionIndex { get; }
            public int Start { get; }
            public int Length { get; }
            public bool ShouldMarkRed { get; }
        }
    }
}
