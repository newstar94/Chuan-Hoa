using System;
using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Scanning
{
    public enum LocalSnapshotValueState
    {
        Unknown = 0,
        None = 1,
        Present = 2
    }

    public sealed class LocalBorderSnapshot
    {
        public LocalBorderSnapshot(LocalSnapshotValueState state, int? lineStyle = null,
            double? weightPoints = null)
        {
            State = state;
            LineStyle = lineStyle;
            WeightPoints = weightPoints;
        }

        public LocalSnapshotValueState State { get; }
        public int? LineStyle { get; }
        public double? WeightPoints { get; }

        public static LocalBorderSnapshot Unknown { get; } =
            new LocalBorderSnapshot(LocalSnapshotValueState.Unknown);

        public static LocalBorderSnapshot None { get; } =
            new LocalBorderSnapshot(LocalSnapshotValueState.None);
    }

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
            double? pageTopPoints = null, double? textWidthPoints = null,
            bool? keepWithNext = null, bool? widowControl = null, string? styleName = null,
            int? absoluteEnd = null, int? builtInStyleId = null,
            bool? hasField = null, bool? hasMathObject = null, bool? hasHyperlink = null,
            bool? hasContentControl = null, string? captionKind = null,
            int tableNestingDepth = 0)
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
            KeepWithNext = keepWithNext; WidowControl = widowControl; StyleName = styleName;
            AbsoluteEnd = absoluteEnd ?? (absoluteStart + (text ?? string.Empty).Length);
            BuiltInStyleId = builtInStyleId; HasField = hasField;
            HasMathObject = hasMathObject; HasHyperlink = hasHyperlink;
            HasContentControl = hasContentControl; CaptionKind = captionKind ?? string.Empty;
            TableNestingDepth = Math.Max(0, tableNestingDepth);
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
        public bool? KeepWithNext { get; }
        public bool? WidowControl { get; }
        public string? StyleName { get; }
        public int AbsoluteEnd { get; }
        public int? BuiltInStyleId { get; }
        public bool? HasField { get; }
        public bool? HasMathObject { get; }
        public bool? HasHyperlink { get; }
        public bool? HasContentControl { get; }
        public string CaptionKind { get; }
        public int TableNestingDepth { get; }
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

    public sealed class LocalTableSnapshot
    {
        public LocalTableSnapshot(
            int index,
            int rowCount,
            int columnCount,
            bool hasMergedCells = false,
            bool isNested = false,
            IReadOnlyList<int>? headerRowIndexes = null,
            bool hasVerticalBorders = false,
            bool hasHeaderSeparatorBorder = true,
            int? associatedCaptionParagraphIndex = null,
            bool isCaptionAbove = true,
            string storyType = "wdMainTextStory", int sectionIndex = 1,
            int absoluteStart = 0, int absoluteEnd = 0, int nestingDepth = 0,
            LocalBorderSnapshot? topBorder = null, LocalBorderSnapshot? bottomBorder = null,
            LocalBorderSnapshot? leftBorder = null, LocalBorderSnapshot? rightBorder = null,
            LocalBorderSnapshot? insideHorizontalBorder = null,
            LocalBorderSnapshot? insideVerticalBorder = null,
            LocalBorderSnapshot? headerSeparatorBorder = null,
            bool? hasMergedCellsState = null)
        {
            Index = index;
            RowCount = rowCount;
            ColumnCount = columnCount;
            HasMergedCells = hasMergedCells;
            IsNested = isNested;
            HeaderRowIndexes = headerRowIndexes ?? Array.Empty<int>();
            HasVerticalBorders = hasVerticalBorders;
            HasHeaderSeparatorBorder = hasHeaderSeparatorBorder;
            AssociatedCaptionParagraphIndex = associatedCaptionParagraphIndex;
            IsCaptionAbove = isCaptionAbove;
            StoryType = storyType ?? string.Empty;
            SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart;
            AbsoluteEnd = Math.Max(absoluteStart, absoluteEnd);
            NestingDepth = Math.Max(0, nestingDepth);
            TopBorder = topBorder ?? LocalBorderSnapshot.Unknown;
            BottomBorder = bottomBorder ?? LocalBorderSnapshot.Unknown;
            LeftBorder = leftBorder ?? (hasVerticalBorders
                ? new LocalBorderSnapshot(LocalSnapshotValueState.Present)
                : LocalBorderSnapshot.None);
            RightBorder = rightBorder ?? (hasVerticalBorders
                ? new LocalBorderSnapshot(LocalSnapshotValueState.Present)
                : LocalBorderSnapshot.None);
            InsideHorizontalBorder = insideHorizontalBorder ?? LocalBorderSnapshot.Unknown;
            InsideVerticalBorder = insideVerticalBorder ?? (hasVerticalBorders
                ? new LocalBorderSnapshot(LocalSnapshotValueState.Present)
                : LocalBorderSnapshot.None);
            HeaderSeparatorBorder = headerSeparatorBorder ?? (hasHeaderSeparatorBorder
                ? new LocalBorderSnapshot(LocalSnapshotValueState.Present)
                : LocalBorderSnapshot.None);
            HasMergedCellsState = hasMergedCellsState ?? hasMergedCells;
        }

        public int Index { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        public bool HasMergedCells { get; }
        public bool IsNested { get; }
        public IReadOnlyList<int> HeaderRowIndexes { get; }
        public bool HasVerticalBorders { get; }
        public bool HasHeaderSeparatorBorder { get; }
        public int? AssociatedCaptionParagraphIndex { get; }
        public bool IsCaptionAbove { get; }
        public string StoryType { get; }
        public int SectionIndex { get; }
        public int AbsoluteStart { get; }
        public int AbsoluteEnd { get; }
        public int NestingDepth { get; }
        public LocalBorderSnapshot TopBorder { get; }
        public LocalBorderSnapshot BottomBorder { get; }
        public LocalBorderSnapshot LeftBorder { get; }
        public LocalBorderSnapshot RightBorder { get; }
        public LocalBorderSnapshot InsideHorizontalBorder { get; }
        public LocalBorderSnapshot InsideVerticalBorder { get; }
        public LocalBorderSnapshot HeaderSeparatorBorder { get; }
        public bool? HasMergedCellsState { get; }
    }

    public sealed class LocalGraphicObjectSnapshot
    {
        public LocalGraphicObjectSnapshot(int index, string objectKind, bool isInline,
            string storyType, int sectionIndex, int absoluteStart, int absoluteEnd,
            int? anchorParagraphIndex = null, bool isProtected = false)
        {
            Index = index;
            ObjectKind = objectKind ?? string.Empty;
            IsInline = isInline;
            StoryType = storyType ?? string.Empty;
            SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart;
            AbsoluteEnd = Math.Max(absoluteStart, absoluteEnd);
            AnchorParagraphIndex = anchorParagraphIndex;
            IsProtected = isProtected;
        }

        public int Index { get; }
        public string ObjectKind { get; }
        public bool IsInline { get; }
        public string StoryType { get; }
        public int SectionIndex { get; }
        public int AbsoluteStart { get; }
        public int AbsoluteEnd { get; }
        public int? AnchorParagraphIndex { get; }
        public bool IsProtected { get; }
    }

    public sealed class LocalScanSnapshot
    {
        public LocalScanSnapshot(string documentFingerprint, long revision, IReadOnlyList<LocalSectionSnapshot> sections,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs, IReadOnlyList<AnnotationProtectedSpan> protectedSpans,
            IReadOnlyList<LocalLineShapeSnapshot>? lineShapes = null, string regimeCode = "UNKNOWN",
            string documentTypeCode = LocalDocumentTypeCodes.Unknown,
            bool regimeWasSelectedManually = false, bool documentTypeWasSelectedManually = false,
            string? dictionaryScopeId = null,
            IReadOnlyList<LocalTableSnapshot>? tables = null,
            IReadOnlyList<LocalGraphicObjectSnapshot>? graphicObjects = null,
            int schemaVersion = 3)
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
            Tables = tables ?? Array.Empty<LocalTableSnapshot>();
            GraphicObjects = graphicObjects ?? Array.Empty<LocalGraphicObjectSnapshot>();
            SchemaVersion = schemaVersion;
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
        public IReadOnlyList<LocalTableSnapshot> Tables { get; }
        public IReadOnlyList<LocalGraphicObjectSnapshot> GraphicObjects { get; }
        public int SchemaVersion { get; }

        public bool IntersectsProtectedSpan(LocalParagraphSnapshot paragraph)
        {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            return ProtectedSpans.Any(span =>
                string.Equals(span.StoryType, paragraph.StoryType, StringComparison.Ordinal) &&
                span.Start < paragraph.AbsoluteEnd &&
                span.Start + span.Length > paragraph.AbsoluteStart);
        }
    }

    public sealed class LocalScanResult
    {
        public LocalScanResult(string scanId, string lane, string rulePackId, string documentFingerprint, long revision,
            IReadOnlyList<AnnotationFinding> findings, string rulePackVersion = "",
            int detectorPolicyVersion = 0,
            IReadOnlyList<string>? notEvaluatedRuleCodes = null,
            bool academicTypographyEnabled = false,
            int academicHeadingCount = 0, int headingLevel1Count = 0,
            int headingLevel2Count = 0, int headingLevel3Count = 0)
        {
            ScanId = scanId; Lane = lane; RulePackId = rulePackId; DocumentFingerprint = documentFingerprint;
            Revision = revision; Findings = findings;
            RulePackVersion = rulePackVersion ?? string.Empty;
            DetectorPolicyVersion = detectorPolicyVersion;
            NotEvaluatedRuleCodes = notEvaluatedRuleCodes ?? Array.Empty<string>();
            AcademicTypographyEnabled = academicTypographyEnabled;
            AcademicHeadingCount = Math.Max(0, academicHeadingCount);
            HeadingLevel1Count = Math.Max(0, headingLevel1Count);
            HeadingLevel2Count = Math.Max(0, headingLevel2Count);
            HeadingLevel3Count = Math.Max(0, headingLevel3Count);
        }
        public string ScanId { get; }
        public string Lane { get; }
        public string RulePackId { get; }
        public string DocumentFingerprint { get; }
        public long Revision { get; }
        public IReadOnlyList<AnnotationFinding> Findings { get; }
        public string RulePackVersion { get; }
        public int DetectorPolicyVersion { get; }
        public IReadOnlyList<string> NotEvaluatedRuleCodes { get; }
        public bool AcademicTypographyEnabled { get; }
        public int AcademicHeadingCount { get; }
        public int HeadingLevel1Count { get; }
        public int HeadingLevel2Count { get; }
        public int HeadingLevel3Count { get; }
    }
}
