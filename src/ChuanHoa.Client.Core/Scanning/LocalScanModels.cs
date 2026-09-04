using System;
using System.Collections.Generic;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Scanning
{
    public sealed class LocalSectionSnapshot
    {
        public LocalSectionSnapshot(int index, double pageWidthPoints, double pageHeightPoints, double topMarginPoints,
            double bottomMarginPoints, double leftMarginPoints, double rightMarginPoints, bool isLandscape,
            bool hasPageNumbers = false, bool restartPageNumbering = false, int? startingPageNumber = null,
            int? pageNumberAlignment = null)
        {
            Index = index; PageWidthPoints = pageWidthPoints; PageHeightPoints = pageHeightPoints;
            TopMarginPoints = topMarginPoints; BottomMarginPoints = bottomMarginPoints;
            LeftMarginPoints = leftMarginPoints; RightMarginPoints = rightMarginPoints; IsLandscape = isLandscape;
            HasPageNumbers = hasPageNumbers; RestartPageNumbering = restartPageNumbering;
            StartingPageNumber = startingPageNumber; PageNumberAlignment = pageNumberAlignment;
        }
        public int Index { get; }
        public double PageWidthPoints { get; }
        public double PageHeightPoints { get; }
        public double TopMarginPoints { get; }
        public double BottomMarginPoints { get; }
        public double LeftMarginPoints { get; }
        public double RightMarginPoints { get; }
        public bool IsLandscape { get; }
        public bool HasPageNumbers { get; }
        public bool RestartPageNumbering { get; }
        public int? StartingPageNumber { get; }
        public int? PageNumberAlignment { get; }
    }

    public sealed class LocalParagraphSnapshot
    {
        public LocalParagraphSnapshot(int index, string text, string storyType, int sectionIndex, int absoluteStart,
            string fontName, int? tableIndex = null, int? rowIndex = null, int? cellIndex = null,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null, int? alignment = null,
            double? firstLineIndentPoints = null, double? spaceBeforePoints = null, double? spaceAfterPoints = null,
            bool isInTable = false, string role = "Unknown", int? fontColor = null, int? underline = null,
            bool hasBottomBorder = false, double? lineSpacingPoints = null, int? lineSpacingRule = null,
            int? outlineLevel = null, int pageNumber = 0, double? pageLeftPoints = null,
            double? pageTopPoints = null, double? textWidthPoints = null)
        {
            Index = index; Text = text ?? string.Empty; StoryType = storyType ?? string.Empty;
            SectionIndex = sectionIndex; AbsoluteStart = absoluteStart; FontName = fontName;
            TableIndex = tableIndex; RowIndex = rowIndex; CellIndex = cellIndex;
            FontSizePoints = fontSizePoints; Bold = bold; Italic = italic; Alignment = alignment;
            FirstLineIndentPoints = firstLineIndentPoints; SpaceBeforePoints = spaceBeforePoints;
            SpaceAfterPoints = spaceAfterPoints; IsInTable = isInTable;
            Role = role ?? "Unknown"; FontColor = fontColor; Underline = underline;
            HasBottomBorder = hasBottomBorder; LineSpacingPoints = lineSpacingPoints;
            LineSpacingRule = lineSpacingRule; OutlineLevel = outlineLevel; PageNumber = pageNumber;
            PageLeftPoints = pageLeftPoints; PageTopPoints = pageTopPoints; TextWidthPoints = textWidthPoints;
        }
        public int Index { get; }
        public string Text { get; }
        public string StoryType { get; }
        public int SectionIndex { get; }
        public int AbsoluteStart { get; }
        public string FontName { get; }
        public int? TableIndex { get; }
        public int? RowIndex { get; }
        public int? CellIndex { get; }
        public double? FontSizePoints { get; }
        public bool? Bold { get; }
        public bool? Italic { get; }
        public int? Alignment { get; }
        public double? FirstLineIndentPoints { get; }
        public double? SpaceBeforePoints { get; }
        public double? SpaceAfterPoints { get; }
        public bool IsInTable { get; }
        public string Role { get; }
        public int? FontColor { get; }
        public int? Underline { get; }
        public bool HasBottomBorder { get; }
        public double? LineSpacingPoints { get; }
        public int? LineSpacingRule { get; }
        public int? OutlineLevel { get; }
        public int PageNumber { get; }
        public double? PageLeftPoints { get; }
        public double? PageTopPoints { get; }
        public double? TextWidthPoints { get; }
    }

    public sealed class LocalLineShapeSnapshot
    {
        public LocalLineShapeSnapshot(int index, string name, int shapeType, string anchorStoryType,
            int anchorSectionIndex, int anchorAbsoluteStart, int? anchorParagraphIndex, int anchorPageNumber,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            double? pageLeftPoints, double? pageTopPoints, int relativeHorizontalPosition,
            int relativeVerticalPosition, bool lineVisible, int? dashStyle, double? weightPoints,
            int? color, int? beginArrowheadStyle, int? endArrowheadStyle)
        {
            Index = index; Name = name ?? string.Empty; ShapeType = shapeType;
            AnchorStoryType = anchorStoryType ?? string.Empty; AnchorSectionIndex = anchorSectionIndex;
            AnchorAbsoluteStart = anchorAbsoluteStart; AnchorParagraphIndex = anchorParagraphIndex;
            AnchorPageNumber = anchorPageNumber; LeftPoints = leftPoints; TopPoints = topPoints;
            WidthPoints = widthPoints; HeightPoints = heightPoints; PageLeftPoints = pageLeftPoints;
            PageTopPoints = pageTopPoints; RelativeHorizontalPosition = relativeHorizontalPosition;
            RelativeVerticalPosition = relativeVerticalPosition; LineVisible = lineVisible;
            DashStyle = dashStyle; WeightPoints = weightPoints; Color = color;
            BeginArrowheadStyle = beginArrowheadStyle; EndArrowheadStyle = endArrowheadStyle;
        }

        public int Index { get; }
        public string Name { get; }
        public int ShapeType { get; }
        public string AnchorStoryType { get; }
        public int AnchorSectionIndex { get; }
        public int AnchorAbsoluteStart { get; }
        public int? AnchorParagraphIndex { get; }
        public int AnchorPageNumber { get; }
        public double LeftPoints { get; }
        public double TopPoints { get; }
        public double WidthPoints { get; }
        public double HeightPoints { get; }
        public double? PageLeftPoints { get; }
        public double? PageTopPoints { get; }
        public int RelativeHorizontalPosition { get; }
        public int RelativeVerticalPosition { get; }
        public bool LineVisible { get; }
        public int? DashStyle { get; }
        public double? WeightPoints { get; }
        public int? Color { get; }
        public int? BeginArrowheadStyle { get; }
        public int? EndArrowheadStyle { get; }
    }

    public sealed class LocalScanSnapshot
    {
        public LocalScanSnapshot(string documentFingerprint, long revision, IReadOnlyList<LocalSectionSnapshot> sections,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs, IReadOnlyList<AnnotationProtectedSpan> protectedSpans,
            IReadOnlyList<LocalLineShapeSnapshot>? lineShapes = null, string regimeCode = "UNKNOWN",
            string documentTypeCode = LocalDocumentTypeCodes.Unknown,
            bool regimeWasSelectedManually = false, bool documentTypeWasSelectedManually = false,
            string? dictionaryScopeId = null)
        {
            DocumentFingerprint = documentFingerprint; Revision = revision; Sections = sections;
            Paragraphs = paragraphs; ProtectedSpans = protectedSpans;
            LineShapes = lineShapes ?? Array.Empty<LocalLineShapeSnapshot>(); RegimeCode = regimeCode ?? "UNKNOWN";
            DocumentTypeCode = documentTypeCode ?? LocalDocumentTypeCodes.Unknown;
            RegimeWasSelectedManually = regimeWasSelectedManually;
            DocumentTypeWasSelectedManually = documentTypeWasSelectedManually;
            DictionaryScopeId = string.IsNullOrWhiteSpace(dictionaryScopeId)
                ? documentFingerprint
                : dictionaryScopeId!;
        }
        public string DocumentFingerprint { get; }
        public long Revision { get; }
        public IReadOnlyList<LocalSectionSnapshot> Sections { get; }
        public IReadOnlyList<LocalParagraphSnapshot> Paragraphs { get; }
        public IReadOnlyList<AnnotationProtectedSpan> ProtectedSpans { get; }
        public IReadOnlyList<LocalLineShapeSnapshot> LineShapes { get; }
        public string RegimeCode { get; }
        public string DocumentTypeCode { get; }
        public bool RegimeWasSelectedManually { get; }
        public bool DocumentTypeWasSelectedManually { get; }
        /// <summary>
        /// Stable key for session-only personal-dictionary ignores. Unlike the content
        /// fingerprint, this value must not change after an ordinary document edit.
        /// Non-Word callers default to the fingerprint for backward compatibility.
        /// </summary>
        public string DictionaryScopeId { get; }
    }

    public sealed class LocalScanResult
    {
        public LocalScanResult(string scanId, string lane, string rulePackId, string documentFingerprint, long revision,
            IReadOnlyList<AnnotationFinding> findings)
        {
            ScanId = scanId; Lane = lane; RulePackId = rulePackId; DocumentFingerprint = documentFingerprint;
            Revision = revision; Findings = findings;
        }
        public string ScanId { get; }
        public string Lane { get; }
        public string RulePackId { get; }
        public string DocumentFingerprint { get; }
        public long Revision { get; }
        public IReadOnlyList<AnnotationFinding> Findings { get; }
    }
}
