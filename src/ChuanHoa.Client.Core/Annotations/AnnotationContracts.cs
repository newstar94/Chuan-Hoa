using System;
using System.Collections.Generic;

namespace ChuanHoa.Client.Core.Annotations
{
    public enum AnnotationSourceFamily
    {
        NotEvaluated = 0,
        Unknown = 1,
        Nd30 = 2,
        Hd05 = 3,
        LatexTypst = 4,
        LocalLanguage = 5
    }

    public enum AnnotationSeverityLevel
    {
        NotEvaluated = 0,
        Unknown = 1,
        Suggestion = 2,
        Warning = 3,
        Error = 4
    }

    public enum AnnotationAnchorKind
    {
        TextSpan,
        Paragraph,
        Section,
        Document
    }

    public enum AnnotationResolutionCode
    {
        Resolved,
        DocumentMismatch,
        RevisionMismatch,
        StoryNotFound,
        ParagraphNotFound,
        SectionNotFound,
        InvalidRange,
        ExpectedTextMismatch,
        ProtectedSpan,
        UnsupportedAnchor,
        ConflictingDuplicateFinding
    }

    public sealed class AnnotationFinding
    {
        public AnnotationFinding(
            string findingId,
            string ruleCode,
            string severity,
            string currentIssue,
            string expected,
            string citation,
            AnnotationAnchor anchor)
            : this(
                findingId,
                ruleCode,
                severity,
                currentIssue,
                expected,
                citation,
                anchor,
                AnnotationSourceFamilyResolver.FromRuleCode(ruleCode))
        {
        }

        public AnnotationFinding(
            string findingId,
            string ruleCode,
            string severity,
            string currentIssue,
            string expected,
            string citation,
            AnnotationAnchor anchor,
            AnnotationSourceFamily sourceFamily)
        {
            FindingId = Require(findingId, nameof(findingId));
            RuleCode = Require(ruleCode, nameof(ruleCode));
            Severity = Require(severity, nameof(severity));
            SeverityLevel = AnnotationSeverityResolver.FromValue(severity);
            CurrentIssue = Require(currentIssue, nameof(currentIssue));
            Expected = expected ?? string.Empty;
            Citation = citation ?? string.Empty;
            Anchor = anchor ?? throw new ArgumentNullException(nameof(anchor));
            SourceFamily = Enum.IsDefined(typeof(AnnotationSourceFamily), sourceFamily)
                ? sourceFamily
                : AnnotationSourceFamily.NotEvaluated;
        }

        public string FindingId { get; }
        public string RuleCode { get; }
        public string Severity { get; }
        public AnnotationSeverityLevel SeverityLevel { get; }
        public string CurrentIssue { get; }
        public string Expected { get; }
        public string Citation { get; }
        public AnnotationAnchor Anchor { get; }
        public AnnotationSourceFamily SourceFamily { get; }

        private static string Require(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("A non-empty value is required.", parameterName);
            }

            return value;
        }
    }

    public sealed class AnnotationAnchor
    {
        public AnnotationAnchor(
            AnnotationAnchorKind kind,
            string storyType,
            int? paragraphIndex,
            int? startOffset,
            int? length,
            string expectedText,
            int? sectionIndex = null,
            int? tableIndex = null,
            int? rowIndex = null,
            int? cellIndex = null)
        {
            Kind = kind;
            StoryType = string.IsNullOrWhiteSpace(storyType) ? "MainTextStory" : storyType;
            ParagraphIndex = paragraphIndex;
            StartOffset = startOffset;
            Length = length;
            ExpectedText = expectedText ?? string.Empty;
            SectionIndex = sectionIndex;
            TableIndex = tableIndex;
            RowIndex = rowIndex;
            CellIndex = cellIndex;
        }

        public AnnotationAnchorKind Kind { get; }
        public string StoryType { get; }
        public int? ParagraphIndex { get; }
        public int? StartOffset { get; }
        public int? Length { get; }
        public string ExpectedText { get; }
        public int? SectionIndex { get; }
        public int? TableIndex { get; }
        public int? RowIndex { get; }
        public int? CellIndex { get; }
    }

    public sealed class AnnotationDocumentSnapshot
    {
        public AnnotationDocumentSnapshot(
            string documentFingerprint,
            long revision,
            IReadOnlyList<AnnotationParagraphSnapshot> paragraphs,
            IReadOnlyList<AnnotationProtectedSpan> protectedSpans)
        {
            DocumentFingerprint = documentFingerprint ?? string.Empty;
            Revision = revision;
            Paragraphs = paragraphs ?? throw new ArgumentNullException(nameof(paragraphs));
            ProtectedSpans = protectedSpans ?? throw new ArgumentNullException(nameof(protectedSpans));
        }

        public string DocumentFingerprint { get; }
        public long Revision { get; }
        public IReadOnlyList<AnnotationParagraphSnapshot> Paragraphs { get; }
        public IReadOnlyList<AnnotationProtectedSpan> ProtectedSpans { get; }
    }

    public sealed class AnnotationParagraphSnapshot
    {
        public AnnotationParagraphSnapshot(
            string storyType,
            int paragraphIndex,
            int sectionIndex,
            int absoluteStart,
            string text,
            int? tableIndex = null,
            int? rowIndex = null,
            int? cellIndex = null)
        {
            StoryType = storyType ?? string.Empty;
            ParagraphIndex = paragraphIndex;
            SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart;
            Text = text ?? string.Empty;
            TableIndex = tableIndex;
            RowIndex = rowIndex;
            CellIndex = cellIndex;
        }

        public string StoryType { get; }
        public int ParagraphIndex { get; }
        public int SectionIndex { get; }
        public int AbsoluteStart { get; }
        public string Text { get; }
        public int? TableIndex { get; }
        public int? RowIndex { get; }
        public int? CellIndex { get; }
    }

    public sealed class AnnotationProtectedSpan
    {
        public AnnotationProtectedSpan(string storyType, int start, int length)
        {
            StoryType = storyType ?? string.Empty;
            Start = start;
            Length = length;
        }

        public string StoryType { get; }
        public int Start { get; }
        public int Length { get; }
    }

    public sealed class AnnotationCommentInstruction
    {
        public AnnotationCommentInstruction(
            string findingId,
            string storyType,
            int sectionIndex,
            int start,
            int length,
            string marker,
            string commentText)
        {
            FindingId = findingId;
            StoryType = storyType;
            SectionIndex = sectionIndex;
            Start = start;
            Length = length;
            Marker = marker;
            CommentText = commentText;
        }

        public string FindingId { get; }
        public string StoryType { get; }
        public int SectionIndex { get; }
        public int Start { get; }
        public int Length { get; }
        public string Marker { get; }
        public string CommentText { get; }
    }

    public sealed class AnnotationVisualInstruction
    {
        public AnnotationVisualInstruction(string storyType, int sectionIndex, int start, int length)
            : this(string.Empty, storyType, sectionIndex, start, length)
        {
        }

        public AnnotationVisualInstruction(string findingId, string storyType, int sectionIndex,
            int start, int length)
        {
            FindingId = findingId ?? string.Empty;
            StoryType = storyType;
            SectionIndex = sectionIndex;
            Start = start;
            Length = length;
        }

        public string FindingId { get; }
        public string StoryType { get; }
        public int SectionIndex { get; }
        public int Start { get; }
        public int Length { get; }
    }

    public sealed class UnresolvedAnnotation
    {
        public UnresolvedAnnotation(string findingId, AnnotationResolutionCode code)
        {
            FindingId = findingId;
            Code = code;
        }

        public string FindingId { get; }
        public AnnotationResolutionCode Code { get; }
    }

    public sealed class AnnotationPlan
    {
        public AnnotationPlan(
            string lane,
            string scanId,
            IReadOnlyList<AnnotationCommentInstruction> comments,
            IReadOnlyList<AnnotationVisualInstruction> visualRanges,
            IReadOnlyList<UnresolvedAnnotation> unresolved)
        {
            Lane = lane;
            ScanId = scanId;
            Comments = comments;
            VisualRanges = visualRanges;
            Unresolved = unresolved;
        }

        public string Lane { get; }
        public string ScanId { get; }
        public IReadOnlyList<AnnotationCommentInstruction> Comments { get; }
        public IReadOnlyList<AnnotationVisualInstruction> VisualRanges { get; }
        public IReadOnlyList<UnresolvedAnnotation> Unresolved { get; }
    }
}
