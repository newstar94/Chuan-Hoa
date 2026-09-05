using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class WordDocumentSnapshot
    {
        public WordDocumentSnapshot(
            int schemaVersion,
            string documentFingerprint,
            long revision,
            string regimeCode,
            string documentTypeCode,
            bool regimeWasSelectedManually,
            bool documentTypeWasSelectedManually,
            WordDocumentPreflight preflight,
            IReadOnlyList<WordSectionSnapshot> sections,
            IReadOnlyList<WordParagraphSnapshot> paragraphs,
            IReadOnlyList<WordLineShapeSnapshot> lineShapes,
            IReadOnlyList<WordTableSnapshot> tables,
            IReadOnlyList<WordGraphicObjectSnapshot> graphicObjects,
            IReadOnlyList<WordProtectedSpanSnapshot> protectedSpans)
        {
            SchemaVersion = schemaVersion;
            DocumentFingerprint = documentFingerprint;
            Revision = revision;
            RegimeCode = regimeCode;
            DocumentTypeCode = documentTypeCode;
            RegimeWasSelectedManually = regimeWasSelectedManually;
            DocumentTypeWasSelectedManually = documentTypeWasSelectedManually;
            Preflight = preflight;
            Sections = sections;
            Paragraphs = paragraphs;
            LineShapes = lineShapes;
            Tables = tables;
            GraphicObjects = graphicObjects;
            ProtectedSpans = protectedSpans;
        }

        public int SchemaVersion { get; }
        public string DocumentFingerprint { get; }
        public long Revision { get; }
        public string RegimeCode { get; }
        public string DocumentTypeCode { get; }
        public bool RegimeWasSelectedManually { get; }
        public bool DocumentTypeWasSelectedManually { get; }
        public WordDocumentPreflight Preflight { get; }
        public IReadOnlyList<WordSectionSnapshot> Sections { get; }
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs { get; }
        public IReadOnlyList<WordLineShapeSnapshot> LineShapes { get; }
        public IReadOnlyList<WordTableSnapshot> Tables { get; }
        public IReadOnlyList<WordGraphicObjectSnapshot> GraphicObjects { get; }
        public IReadOnlyList<WordProtectedSpanSnapshot> ProtectedSpans { get; }
    }

    public sealed class WordDocumentPreflight
    {
        public WordDocumentPreflight(
            bool isReadOnly,
            bool isProtected,
            bool isCompatibilityMode,
            bool trackChangesEnabled,
            string fileFormat,
            bool hasActiveWindow)
        {
            IsReadOnly = isReadOnly;
            IsProtected = isProtected;
            IsCompatibilityMode = isCompatibilityMode;
            TrackChangesEnabled = trackChangesEnabled;
            FileFormat = fileFormat;
            HasActiveWindow = hasActiveWindow;
        }

        public bool IsReadOnly { get; }
        public bool IsProtected { get; }
        public bool IsCompatibilityMode { get; }
        public bool TrackChangesEnabled { get; }
        public string FileFormat { get; }
        public bool HasActiveWindow { get; }
    }

    public sealed class WordSectionSnapshot
    {
        public WordSectionSnapshot(
            int index,
            double pageWidthPoints,
            double pageHeightPoints,
            double topMarginPoints,
            double bottomMarginPoints,
            double leftMarginPoints,
            double rightMarginPoints,
            bool isLandscape,
            bool hasPageNumbers,
            bool restartPageNumbering,
            int? startingPageNumber,
            int? pageNumberAlignment)
        {
            Index = index;
            PageWidthPoints = pageWidthPoints;
            PageHeightPoints = pageHeightPoints;
            TopMarginPoints = topMarginPoints;
            BottomMarginPoints = bottomMarginPoints;
            LeftMarginPoints = leftMarginPoints;
            RightMarginPoints = rightMarginPoints;
            IsLandscape = isLandscape;
            HasPageNumbers = hasPageNumbers;
            RestartPageNumbering = restartPageNumbering;
            StartingPageNumber = startingPageNumber;
            PageNumberAlignment = pageNumberAlignment;
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

    public sealed class WordParagraphSnapshot
    {
        public WordParagraphSnapshot(
            int index,
            string text,
            string role,
            double roleConfidence,
            string? fontName,
            double? fontSizePoints,
            bool? bold,
            bool? italic,
            int? alignment,
            double? firstLineIndentPoints,
            double? spaceBeforePoints,
            double? spaceAfterPoints,
            bool isInTable,
            string storyType,
            int sectionIndex,
            int absoluteStart,
            int? tableIndex,
            int? rowIndex,
            int? cellIndex,
            int? fontColor,
            int? underline,
            bool hasBottomBorder,
            double? lineSpacingPoints,
            int? lineSpacingRule,
            int? outlineLevel,
            int pageNumber,
            double? pageLeftPoints,
            double? pageTopPoints,
            double? textWidthPoints,
            bool? keepWithNext = null,
            bool? widowControl = null,
            string? styleName = null, int? absoluteEnd = null,
            int? builtInStyleId = null, bool? hasField = null,
            bool? hasMathObject = null, bool? hasHyperlink = null,
            bool? hasContentControl = null, string? captionKind = null,
            int tableNestingDepth = 0)
        {
            Index = index;
            Text = text;
            Role = role;
            RoleConfidence = roleConfidence;
            FontName = fontName;
            FontSizePoints = fontSizePoints;
            Bold = bold;
            Italic = italic;
            Alignment = alignment;
            FirstLineIndentPoints = firstLineIndentPoints;
            SpaceBeforePoints = spaceBeforePoints;
            SpaceAfterPoints = spaceAfterPoints;
            IsInTable = isInTable;
            StoryType = storyType;
            SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart;
            TableIndex = tableIndex;
            RowIndex = rowIndex;
            CellIndex = cellIndex;
            FontColor = fontColor;
            Underline = underline;
            HasBottomBorder = hasBottomBorder;
            LineSpacingPoints = lineSpacingPoints;
            LineSpacingRule = lineSpacingRule;
            OutlineLevel = outlineLevel;
            PageNumber = pageNumber;
            PageLeftPoints = pageLeftPoints;
            PageTopPoints = pageTopPoints;
            TextWidthPoints = textWidthPoints;
            KeepWithNext = keepWithNext;
            WidowControl = widowControl;
            StyleName = styleName;
            AbsoluteEnd = absoluteEnd ?? absoluteStart + (text ?? string.Empty).Length;
            BuiltInStyleId = builtInStyleId;
            HasField = hasField;
            HasMathObject = hasMathObject;
            HasHyperlink = hasHyperlink;
            HasContentControl = hasContentControl;
            CaptionKind = captionKind ?? string.Empty;
            TableNestingDepth = Math.Max(0, tableNestingDepth);
        }

        public int Index { get; }
        public string Text { get; }
        public string Role { get; }
        public double RoleConfidence { get; }
        public string? FontName { get; }
        public double? FontSizePoints { get; }
        public bool? Bold { get; }
        public bool? Italic { get; }
        public int? Alignment { get; }
        public double? FirstLineIndentPoints { get; }
        public double? SpaceBeforePoints { get; }
        public double? SpaceAfterPoints { get; }
        public bool IsInTable { get; }
        public string StoryType { get; }
        public int SectionIndex { get; }
        public int AbsoluteStart { get; }
        public int? TableIndex { get; }
        public int? RowIndex { get; }
        public int? CellIndex { get; }
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

    public sealed class WordLineShapeSnapshot
    {
        public WordLineShapeSnapshot(int index, string name, int shapeType, string anchorStoryType,
            int anchorSectionIndex, int anchorAbsoluteStart, int? anchorParagraphIndex, int anchorPageNumber,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            double? pageLeftPoints, double? pageTopPoints, int relativeHorizontalPosition,
            int relativeVerticalPosition, bool lineVisible, int? dashStyle, double? weightPoints,
            int? color, int? beginArrowheadStyle, int? endArrowheadStyle)
        {
            Index = index; Name = name; ShapeType = shapeType; AnchorStoryType = anchorStoryType;
            AnchorSectionIndex = anchorSectionIndex; AnchorAbsoluteStart = anchorAbsoluteStart;
            AnchorParagraphIndex = anchorParagraphIndex; AnchorPageNumber = anchorPageNumber;
            LeftPoints = leftPoints; TopPoints = topPoints; WidthPoints = widthPoints; HeightPoints = heightPoints;
            PageLeftPoints = pageLeftPoints; PageTopPoints = pageTopPoints;
            RelativeHorizontalPosition = relativeHorizontalPosition; RelativeVerticalPosition = relativeVerticalPosition;
            LineVisible = lineVisible; DashStyle = dashStyle; WeightPoints = weightPoints; Color = color;
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

    public sealed class WordTableSnapshot
    {
        public WordTableSnapshot(
            int index,
            int rowCount,
            int columnCount,
            bool hasMergedCells,
            bool isNested,
            IReadOnlyList<int> headerRowIndexes,
            bool hasVerticalBorders = false,
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
            HeaderRowIndexes = headerRowIndexes;
            HasVerticalBorders = hasVerticalBorders;
            StoryType = storyType ?? string.Empty;
            SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart;
            AbsoluteEnd = Math.Max(absoluteStart, absoluteEnd);
            NestingDepth = Math.Max(0, nestingDepth);
            TopBorder = topBorder ?? LocalBorderSnapshot.Unknown;
            BottomBorder = bottomBorder ?? LocalBorderSnapshot.Unknown;
            LeftBorder = leftBorder ?? LocalBorderSnapshot.Unknown;
            RightBorder = rightBorder ?? LocalBorderSnapshot.Unknown;
            InsideHorizontalBorder = insideHorizontalBorder ?? LocalBorderSnapshot.Unknown;
            InsideVerticalBorder = insideVerticalBorder ?? LocalBorderSnapshot.Unknown;
            HeaderSeparatorBorder = headerSeparatorBorder ?? LocalBorderSnapshot.Unknown;
            HasMergedCellsState = hasMergedCellsState;
        }

        public int Index { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        public bool HasMergedCells { get; }
        public bool IsNested { get; }
        public IReadOnlyList<int> HeaderRowIndexes { get; }
        public bool HasVerticalBorders { get; }
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

    public sealed class WordGraphicObjectSnapshot
    {
        public WordGraphicObjectSnapshot(int index, string objectKind, bool isInline,
            string storyType, int sectionIndex, int absoluteStart, int absoluteEnd,
            int? anchorParagraphIndex, bool isProtected)
        {
            Index = index; ObjectKind = objectKind ?? string.Empty; IsInline = isInline;
            StoryType = storyType ?? string.Empty; SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart; AbsoluteEnd = Math.Max(absoluteStart, absoluteEnd);
            AnchorParagraphIndex = anchorParagraphIndex; IsProtected = isProtected;
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

    public sealed class WordProtectedSpanSnapshot
    {
        public WordProtectedSpanSnapshot(
            string storyType,
            int sectionIndex,
            int absoluteStart,
            int length,
            string kind)
        {
            StoryType = storyType;
            SectionIndex = sectionIndex;
            AbsoluteStart = absoluteStart;
            Length = length;
            Kind = kind;
        }

        public string StoryType { get; }
        public int SectionIndex { get; }
        public int AbsoluteStart { get; }
        public int Length { get; }
        public string Kind { get; }
    }

    public sealed class WordDocumentSnapshotBuilder
    {
        private const int SnapshotSchemaVersion = 3;
        private readonly Action<string, long>? _stageCompleted;

        public WordDocumentSnapshotBuilder()
            : this(null)
        {
        }

        internal WordDocumentSnapshotBuilder(Action<string, long>? stageCompleted)
        {
            _stageCompleted = stageCompleted;
        }

        private sealed class SectionBoundary
        {
            public SectionBoundary(int index, int start, int end)
            {
                Index = index;
                Start = start;
                End = end;
            }

            public int Index { get; }
            public int Start { get; }
            public int End { get; }
        }

        public WordDocumentSnapshot Build(
            Word.Document document,
            DocumentContext context,
            WordDocumentCapability capability,
            DocumentOperationSession? operation = null)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            if (capability == null || !capability.CanReadDocument)
            {
                throw new InvalidOperationException(capability == null
                    ? "Document capability is required."
                    : capability.Reason);
            }

            operation?.Transition(DocumentOperationState.Capturing, "đọc section");
            var sections = Measure("sections", () => CaptureSections(document, operation));
            var sectionBoundaries = Measure("section-boundaries", () => CaptureSectionBoundaries(document));
            operation?.Checkpoint();
            operation?.ReportProgress(0, 1, "đọc bảng");
            var tables = Measure("tables", () => CaptureTables(document, operation));
            var allowPageLayout = Measure("page-layout-safety", () => IsPageLayoutCaptureSafe(document));
            var builtInStyles = Measure("built-in-styles", () => CaptureBuiltInStyleNames(document));
            var paragraphs = Measure("paragraphs", () =>
                CaptureParagraphs(document, tables, sectionBoundaries, builtInStyles, allowPageLayout, operation));
            operation?.ReportProgress(paragraphs.Count, paragraphs.Count, "đọc đường kẻ");
            var lineShapes = Measure("line-shapes", () =>
                CaptureLineShapes(document, sections, paragraphs, operation));
            operation?.ReportProgress(paragraphs.Count, paragraphs.Count, "đọc hình và đối tượng");
            var protectedSpans = Measure("protected-spans", () => CaptureProtectedSpans(document, operation));
            var graphicObjects = Measure("graphic-objects", () =>
                CaptureGraphicObjects(document, paragraphs, protectedSpans, operation));
            operation?.ReportProgress(paragraphs.Count, paragraphs.Count, "tạo dấu nhận dạng tài liệu");
            var fingerprint = Measure("fingerprint", () =>
                CaptureFingerprint(document, sections, paragraphs, lineShapes, tables,
                    graphicObjects, protectedSpans));
            operation?.Checkpoint();
            var preflight = new WordDocumentPreflight(
                capability.IsReadOnly,
                capability.IsProtected,
                document.CompatibilityMode != 0,
                capability.TrackChangesEnabled,
                ReadFileFormat(document),
                capability.HasActiveWindow);

            return new WordDocumentSnapshot(
                SnapshotSchemaVersion,
                fingerprint,
                context.NextSnapshotRevision(),
                context.RegimeCode,
                context.DocumentTypeCode,
                context.RegimeWasSelectedManually,
                context.DocumentTypeWasSelectedManually,
                preflight,
                sections,
                paragraphs,
                lineShapes,
                tables,
                graphicObjects,
                protectedSpans);
        }

        private T Measure<T>(string stage, Func<T> action)
        {
            if (_stageCompleted == null) return action();
            var timer = Stopwatch.StartNew();
            try { return action(); }
            finally
            {
                timer.Stop();
                _stageCompleted(stage, timer.ElapsedMilliseconds);
            }
        }

        private long BeginTiming() => _stageCompleted == null ? 0L : Stopwatch.GetTimestamp();

        private static void AddElapsed(ref long totalTicks, long started)
        {
            if (started != 0L) totalTicks += Stopwatch.GetTimestamp() - started;
        }

        private void ReportParagraphTiming(Word.WdStoryType storyType, string stage, long ticks)
        {
            if (_stageCompleted == null || ticks == 0L) return;
            var milliseconds = (long)Math.Round(ticks * 1000d / Stopwatch.Frequency,
                MidpointRounding.AwayFromZero);
            _stageCompleted("paragraphs/" + storyType + "/" + stage, milliseconds);
        }

        private static IReadOnlyList<WordSectionSnapshot> CaptureSections(Word.Document document,
            DocumentOperationSession? operation)
        {
            var snapshots = new List<WordSectionSnapshot>();
            Word.Sections? documentSections = null;
            try
            {
                documentSections = document.Sections;
                var sectionCount = documentSections.Count;
                for (var index = 1; index <= sectionCount; index++)
                {
                    operation?.ReportProgress(index - 1, sectionCount, "đọc section");
                    Word.Section? section = null;
                    Word.PageSetup? setup = null;
                    try
                    {
                        section = documentSections[index];
                        setup = section.PageSetup;
                        var numbering = ReadPageNumbering(section);
                        snapshots.Add(new WordSectionSnapshot(
                            index,
                            setup.PageWidth,
                            setup.PageHeight,
                            setup.TopMargin,
                            setup.BottomMargin,
                            setup.LeftMargin,
                            setup.RightMargin,
                            setup.Orientation == Word.WdOrientation.wdOrientLandscape,
                            numbering.Item1, numbering.Item2, numbering.Item3, numbering.Item4));
                    }
                    finally
                    {
                        Release(setup);
                        Release(section);
                    }
                }
            }
            finally { Release(documentSections); }

            return snapshots;
        }

        private static IReadOnlyList<SectionBoundary> CaptureSectionBoundaries(Word.Document document)
        {
            var result = new List<SectionBoundary>();
            Word.Sections? sections = null;
            try
            {
                sections = document.Sections;
                var count = sections.Count;
                for (var index = 1; index <= count; index++)
                {
                    Word.Section? section = null;
                    Word.Range? range = null;
                    try
                    {
                        section = sections[index];
                        range = section.Range.Duplicate;
                        result.Add(new SectionBoundary(index, range.Start, range.End));
                    }
                    finally
                    {
                        Release(range);
                        Release(section);
                    }
                }
            }
            finally { Release(sections); }
            return result;
        }

        private IReadOnlyList<WordParagraphSnapshot> CaptureParagraphs(Word.Document document,
            IReadOnlyList<WordTableSnapshot> tables,
            IReadOnlyList<SectionBoundary> sectionBoundaries,
            IReadOnlyDictionary<string, int> builtInStyles,
            bool allowPageLayout, DocumentOperationSession? operation)
        {
            var snapshots = new List<WordParagraphSnapshot>();
            var paragraphIndex = 0;
            var application = document.Application;
            try
            {
                foreach (var storyType in SupportedStoryTypes())
                {
                    Word.Range? story = null;
                    try
                    {
                        story = document.StoryRanges[storyType];
                    }
                    catch (COMException)
                    {
                        continue;
                    }

                    while (story != null)
                    {
                        Word.Range? nextStory = null;
                        try
                        {
                            CaptureStoryParagraphs(story, storyType, snapshots, ref paragraphIndex,
                                tables, sectionBoundaries, builtInStyles, allowPageLayout, application, operation);
                            nextStory = story.NextStoryRange;
                        }
                        finally
                        {
                            Release(story);
                        }

                        story = nextStory;
                    }
                }
            }
            finally
            {
                try { application.StatusBar = string.Empty; }
                catch (COMException) { }
                Release(application);
            }

            return snapshots;
        }

        private void CaptureStoryParagraphs(
            Word.Range story,
            Word.WdStoryType storyType,
            ICollection<WordParagraphSnapshot> snapshots,
            ref int paragraphIndex,
            IReadOnlyList<WordTableSnapshot> tables,
            IReadOnlyList<SectionBoundary> sectionBoundaries,
            IReadOnlyDictionary<string, int> builtInStyles,
            bool allowPageLayout,
            Word.Application application,
            DocumentOperationSession? operation)
        {
            var paragraphs = story.Paragraphs;
            long acquireTicks = 0;
            long tableTicks = 0;
            long basicFormatTicks = 0;
            long borderTicks = 0;
            long styleTicks = 0;
            long semanticTicks = 0;
            long sectionTicks = 0;
            long layoutTicks = 0;
            long snapshotTicks = 0;
            try
            {
                // Cache Count once. Re-reading this COM property for every item is
                // disproportionately expensive for stories containing large tables.
                var paragraphCount = paragraphs.Count;
                var largeStory = paragraphCount > 400;
                if (largeStory)
                {
                    CaptureLargeStoryParagraphs(story, storyType, snapshots, ref paragraphIndex,
                        tables, sectionBoundaries, builtInStyles, application, paragraphCount, operation);
                    return;
                }
                var currentSectionIndex = ReadSectionIndex(story);
                var storyParagraphIndex = 0;
                for (var itemIndex = 1; itemIndex <= paragraphCount; itemIndex++)
                {
                    storyParagraphIndex++;
                    if (storyParagraphIndex == 1 || storyParagraphIndex % 25 == 0)
                        PulseWordUi(application, storyParagraphIndex, paragraphCount, storyType, operation);
                    Word.Paragraph? paragraph = null;
                    Word.Range? range = null;
                    Word.Font? font = null;
                    Word.ParagraphFormat? format = null;
                    try
                    {
                        var phaseStarted = BeginTiming();
                        paragraph = paragraphs[itemIndex];
                        range = paragraph.Range.Duplicate;
                        paragraphIndex++;
                        var absoluteStart = range.Start;
                        var absoluteEnd = range.End;
                        var text = NormalizeParagraphText(range.Text);
                        AddElapsed(ref acquireTicks, phaseStarted);
                        phaseStarted = BeginTiming();
                        var tableAtRange = FindTableAt(tables, storyType, absoluteStart);
                        var withInTable = tableAtRange != null;
                        var tableCoordinates = ReadTableCoordinates(tableAtRange, range);
                        AddElapsed(ref tableTicks, phaseStarted);
                        var captureFormatting = !string.IsNullOrWhiteSpace(text) &&
                            (!largeStory || !withInTable || paragraphIndex <= 150);
                        string? fontName = null;
                        double? fontSize = null;
                        bool? bold = null;
                        bool? italic = null;
                        int? alignment = null;
                        double? firstLineIndent = null;
                        double? spaceBefore = null;
                        double? spaceAfter = null;
                        int? fontColor = null;
                        int? underline = null;
                        var hasBottomBorder = false;
                        double? lineSpacing = null;
                        int? lineSpacingRule = null;
                        int? outlineLevel = null;
                        bool? keepWithNext = null;
                        bool? widowControl = null;
                        string? styleName = null;
                        int? builtInStyleId = null;
                        bool? hasField = null;
                        bool? hasMathObject = null;
                        bool? hasHyperlink = null;
                        bool? hasContentControl = null;
                        string? captionKind = null;
                        if (captureFormatting)
                        {
                            phaseStarted = BeginTiming();
                            font = range.Font;
                            format = range.ParagraphFormat;
                            fontName = ReadFontName(font);
                            fontSize = ReadNullableFloat(font.Size);
                            bold = ReadNullableBoolean(font.Bold);
                            italic = ReadNullableBoolean(font.Italic);
                            alignment = (int)format.Alignment;
                            firstLineIndent = ReadNullableFloat(format.FirstLineIndent);
                            spaceBefore = ReadNullableFloat(format.SpaceBefore);
                            spaceAfter = ReadNullableFloat(format.SpaceAfter);
                            fontColor = ReadNullableInteger((int)font.Color);
                            underline = ReadNullableInteger((int)font.Underline);
                            lineSpacing = ReadNullableFloat(format.LineSpacing);
                            lineSpacingRule = ReadNullableInteger((int)format.LineSpacingRule);
                            outlineLevel = ReadNullableInteger((int)format.OutlineLevel);
                            keepWithNext = ReadNullableBoolean(format.KeepWithNext);
                            widowControl = ReadNullableBoolean(format.WidowControl);
                            AddElapsed(ref basicFormatTicks, phaseStarted);
                            if (ShouldCaptureStyleIdentity(text, withInTable, bold, outlineLevel))
                            {
                                phaseStarted = BeginTiming();
                                ReadStyleIdentity(range, builtInStyles, out styleName, out builtInStyleId);
                                AddElapsed(ref styleTicks, phaseStarted);
                            }
                            if (ShouldCaptureSemanticFacts(text))
                            {
                                phaseStarted = BeginTiming();
                                ReadRangeSemanticFacts(range, out hasField, out hasMathObject,
                                    out hasHyperlink, out hasContentControl, out captionKind);
                                AddElapsed(ref semanticTicks, phaseStarted);
                            }
                        }
                        phaseStarted = BeginTiming();
                        if (storyType == Word.WdStoryType.wdMainTextStory)
                            currentSectionIndex = ResolveSectionIndex(sectionBoundaries, range.Start,
                                currentSectionIndex);
                        AddElapsed(ref sectionTicks, phaseStarted);
                        double? textWidth = null;
                        double? pageLeft = null;
                        double? pageTop = null;
                        var pageNumber = 0;
                        if (allowPageLayout && paragraphCount <= 400 &&
                            ShouldCapturePageLayout(storyType, paragraphIndex, text, alignment.GetValueOrDefault(-1)))
                        {
                            phaseStarted = BeginTiming();
                            // Word's page-relative Information properties trigger full
                            // repagination. Calling them for every cell paragraph can keep
                            // Word's UI thread busy indefinitely on table-heavy documents.
                            // Only first-page, centered heading candidates need these
                            // coordinates for Line Shape validation.
                            pageNumber = SafeInformation(range, Word.WdInformation.wdActiveEndAdjustedPageNumber);
                            pageTop = WordTextMeasurement.ReadFirstTextLineTop(range);
                            textWidth = WordTextMeasurement.MeasureParagraphWidth(range, fontName, fontSize,
                                bold, italic);
                            var center = WordTextMeasurement.ReadHorizontalCenter(range);
                            if (center.HasValue && textWidth.GetValueOrDefault() > 0d)
                                pageLeft = center.Value - textWidth.Value / 2d;
                            AddElapsed(ref layoutTicks, phaseStarted);
                        }
                        phaseStarted = BeginTiming();
                        snapshots.Add(new WordParagraphSnapshot(
                            paragraphIndex,
                            text,
                            "Unknown",
                            0.0d,
                            fontName,
                            fontSize,
                            bold,
                            italic,
                            alignment,
                            firstLineIndent,
                            spaceBefore,
                            spaceAfter,
                            withInTable,
                            storyType.ToString(),
                            currentSectionIndex,
                            absoluteStart,
                            tableCoordinates.Item1,
                            tableCoordinates.Item2,
                            tableCoordinates.Item3,
                            fontColor,
                            underline,
                            hasBottomBorder,
                            lineSpacing,
                            lineSpacingRule,
                            outlineLevel,
                            pageNumber,
                            pageLeft,
                            pageTop,
                            textWidth,
                            keepWithNext,
                            widowControl,
                            styleName,
                            absoluteEnd,
                            builtInStyleId,
                            hasField,
                            hasMathObject,
                            hasHyperlink,
                            hasContentControl,
                            captionKind,
                            tableCoordinates.Item4));
                        AddElapsed(ref snapshotTicks, phaseStarted);
                    }
                    finally
                    {
                        Release(format);
                        Release(font);
                        Release(range);
                        Release(paragraph);
                    }
                }
            }
            finally
            {
                Release(paragraphs);
                ReportParagraphTiming(storyType, "acquire", acquireTicks);
                ReportParagraphTiming(storyType, "table", tableTicks);
                ReportParagraphTiming(storyType, "basic-format", basicFormatTicks);
                ReportParagraphTiming(storyType, "border", borderTicks);
                ReportParagraphTiming(storyType, "style", styleTicks);
                ReportParagraphTiming(storyType, "semantic", semanticTicks);
                ReportParagraphTiming(storyType, "section", sectionTicks);
                ReportParagraphTiming(storyType, "layout", layoutTicks);
                ReportParagraphTiming(storyType, "snapshot", snapshotTicks);
            }
        }

        private static void CaptureLargeStoryParagraphs(Word.Range story, Word.WdStoryType storyType,
            ICollection<WordParagraphSnapshot> snapshots, ref int paragraphIndex,
            IReadOnlyList<WordTableSnapshot> tables,
            IReadOnlyList<SectionBoundary> sectionBoundaries,
            IReadOnlyDictionary<string, int> builtInStyles,
            Word.Application application, int expectedParagraphCount,
            DocumentOperationSession? operation)
        {
            var storyParagraphIndex = 0;
            var currentSectionIndex = ReadSectionIndex(story);
            Word.Paragraphs? paragraphs = null;
            try
            {
                paragraphs = story.Paragraphs;
                var paragraphCount = paragraphs.Count;
                for (var itemIndex = 1; itemIndex <= paragraphCount; itemIndex++)
                {
                    Word.Paragraph? paragraph = null;
                    Word.Range? range = null;
                    Word.Font? font = null;
                    Word.ParagraphFormat? format = null;
                    try
                    {
                        paragraph = paragraphs[itemIndex];
                        storyParagraphIndex++;
                        paragraphIndex++;
                        if (storyParagraphIndex == 1 || storyParagraphIndex % 25 == 0)
                            PulseWordUi(application, storyParagraphIndex, expectedParagraphCount, storyType, operation);

                        // Range.Text cannot be used as a coordinate map for tables.
                        // A cell terminator is returned as CR + BEL (two chars), while
                        // Word's Start/End coordinate advances by either one or two
                        // positions depending on the table structure. Read the COM
                        // paragraph range itself so every later annotation has the
                        // authoritative Word coordinate.
                        range = paragraph.Range.Duplicate;
                        var absoluteStart = range.Start;
                        var absoluteEnd = range.End;
                        var text = NormalizeParagraphText(range.Text);
                        var tableAtRange = FindTableAt(tables, storyType, absoluteStart);
                        var tableCoordinates = ReadTableCoordinates(tableAtRange, range);
                        var isInTable = tableCoordinates.Item1.HasValue;
                        var captureFormatting = !string.IsNullOrWhiteSpace(text) &&
                            (!isInTable || storyParagraphIndex <= 150);
                        string? fontName = null;
                        double? fontSize = null;
                        bool? bold = null;
                        bool? italic = null;
                        int? alignment = null;
                        double? firstLineIndent = null;
                        double? spaceBefore = null;
                        double? spaceAfter = null;
                        int? fontColor = null;
                        int? underline = null;
                        var hasBottomBorder = false;
                        double? lineSpacing = null;
                        int? lineSpacingRule = null;
                        int? outlineLevel = null;
                        bool? keepWithNext = null;
                        bool? widowControl = null;
                        string? styleName = null;
                        int? builtInStyleId = null;
                        bool? hasField = null;
                        bool? hasMathObject = null;
                        bool? hasHyperlink = null;
                        bool? hasContentControl = null;
                        string? captionKind = null;
                        if (captureFormatting)
                        {
                            font = range.Font;
                            format = range.ParagraphFormat;
                            fontName = ReadFontName(font);
                            fontSize = ReadNullableFloat(font.Size);
                            bold = ReadNullableBoolean(font.Bold);
                            italic = ReadNullableBoolean(font.Italic);
                            alignment = (int)format.Alignment;
                            firstLineIndent = ReadNullableFloat(format.FirstLineIndent);
                            spaceBefore = ReadNullableFloat(format.SpaceBefore);
                            spaceAfter = ReadNullableFloat(format.SpaceAfter);
                            fontColor = ReadNullableInteger((int)font.Color);
                            underline = ReadNullableInteger((int)font.Underline);
                            lineSpacing = ReadNullableFloat(format.LineSpacing);
                            lineSpacingRule = ReadNullableInteger((int)format.LineSpacingRule);
                            outlineLevel = ReadNullableInteger((int)format.OutlineLevel);
                            keepWithNext = ReadNullableBoolean(format.KeepWithNext);
                            widowControl = ReadNullableBoolean(format.WidowControl);
                            if (ShouldCaptureStyleIdentity(text, isInTable, bold, outlineLevel))
                                ReadStyleIdentity(range, builtInStyles, out styleName, out builtInStyleId);
                            if (ShouldCaptureSemanticFacts(text))
                                ReadRangeSemanticFacts(range, out hasField, out hasMathObject,
                                    out hasHyperlink, out hasContentControl, out captionKind);
                        }
                        if (storyType == Word.WdStoryType.wdMainTextStory)
                            currentSectionIndex = ResolveSectionIndex(sectionBoundaries, absoluteStart,
                                currentSectionIndex);

                        snapshots.Add(new WordParagraphSnapshot(paragraphIndex, text, "Unknown", 0.0d,
                            fontName, fontSize, bold, italic, alignment, firstLineIndent, spaceBefore,
                            spaceAfter, isInTable, storyType.ToString(), currentSectionIndex, absoluteStart,
                            tableCoordinates.Item1, tableCoordinates.Item2, tableCoordinates.Item3,
                            fontColor, underline, hasBottomBorder, lineSpacing,
                            lineSpacingRule, outlineLevel, 0, null, null, null,
                            keepWithNext, widowControl, styleName, absoluteEnd, builtInStyleId,
                            hasField, hasMathObject, hasHyperlink, hasContentControl, captionKind,
                            tableCoordinates.Item4));
                    }
                    finally
                    {
                        Release(format);
                        Release(font);
                        Release(range);
                        Release(paragraph);
                    }
                }
            }
            finally
            {
                Release(paragraphs);
            }

            if (storyParagraphIndex != expectedParagraphCount)
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture,
                    "Word paragraph stream mismatch: captured {0}, expected {1}.",
                    storyParagraphIndex, expectedParagraphCount));
        }

        private static IReadOnlyList<Tuple<int, int>> CaptureTableSpans(Word.Range story)
        {
            var result = new List<Tuple<int, int>>();
            Word.Tables? tables = null;
            try
            {
                tables = story.Tables;
                var tableCount = tables.Count;
                for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    Word.Table? table = null;
                    Word.Range? range = null;
                    try
                    {
                        table = tables[tableIndex];
                        range = table.Range;
                        result.Add(Tuple.Create(range.Start, range.End));
                    }
                    finally
                    {
                        Release(range);
                        Release(table);
                    }
                }
            }
            finally
            {
                Release(tables);
            }
            return result;
        }

        private static void PulseWordUi(Word.Application application, int completedParagraphs,
            int storyParagraphs, Word.WdStoryType storyType,
            DocumentOperationSession? operation)
        {
            if (operation != null)
            {
                operation.ReportProgress(completedParagraphs, storyParagraphs,
                    "vùng " + storyType);
                return;
            }
            try
            {
                application.StatusBar = string.Format(CultureInfo.CurrentCulture,
                    "Chuẩn hóa đang đọc dữ liệu: {0:N0} đoạn (vùng {1}, {2:N0} đoạn)...",
                    completedParagraphs, storyType, storyParagraphs);
            }
            catch (COMException)
            {
            }

            // Snapshot capture stays on Word's STA thread. Do not pump the Windows
            // message queue here: doing so can re-enter another Ribbon callback while
            // Word still owns transient Range/Table RCWs from this capture batch.
        }

        private static IReadOnlyList<WordLineShapeSnapshot> CaptureLineShapes(
            Word.Document document,
            IReadOnlyList<WordSectionSnapshot> sections,
            IReadOnlyList<WordParagraphSnapshot> paragraphs,
            DocumentOperationSession? operation)
        {
            var snapshots = new List<WordLineShapeSnapshot>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            Word.Shapes? documentShapes = null;
            try
            {
                documentShapes = document.Shapes;
                CaptureShapeCollection(documentShapes, sections, paragraphs, snapshots, seen, operation);
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(documentShapes);
            }

            Word.Sections? documentSections = null;
            try
            {
                documentSections = document.Sections;
                var sectionCount = documentSections.Count;
                for (var sectionIndex = 1; sectionIndex <= sectionCount; sectionIndex++)
                {
                    if (sectionIndex == 1 || sectionIndex % 25 == 0)
                        operation?.ReportProgress(sectionIndex - 1, sectionCount,
                            "đọc đường kẻ trong section");
                    Word.Section? section = null;
                    Word.HeadersFooters? headers = null;
                    Word.HeadersFooters? footers = null;
                    try
                    {
                        section = documentSections[sectionIndex];
                        headers = section.Headers;
                        footers = section.Footers;
                        CaptureHeaderFooterShapes(headers, sections, paragraphs, snapshots, seen, operation);
                        headers = null;
                        CaptureHeaderFooterShapes(footers, sections, paragraphs, snapshots, seen, operation);
                        footers = null;
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        Release(footers);
                        Release(headers);
                        Release(section);
                    }
                }
            }
            finally { Release(documentSections); }

            return snapshots;
        }

        private static void CaptureHeaderFooterShapes(Word.HeadersFooters collection,
            IReadOnlyList<WordSectionSnapshot> sections, IReadOnlyList<WordParagraphSnapshot> paragraphs,
            ICollection<WordLineShapeSnapshot> snapshots, ISet<string> seen,
            DocumentOperationSession? operation)
        {
            try
            {
                foreach (Word.WdHeaderFooterIndex index in new[]
                {
                    Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                    Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                    Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages
                })
                {
                    operation?.Checkpoint();
                    Word.HeaderFooter? headerFooter = null;
                    Word.Shapes? shapes = null;
                    try
                    {
                        headerFooter = collection[index];
                        if (!headerFooter.Exists) continue;
                        shapes = headerFooter.Shapes;
                        CaptureShapeCollection(shapes, sections, paragraphs, snapshots, seen, operation);
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        Release(shapes);
                        Release(headerFooter);
                    }
                }
            }
            finally
            {
                Release(collection);
            }
        }

        private static void CaptureShapeCollection(Word.Shapes shapes,
            IReadOnlyList<WordSectionSnapshot> sections, IReadOnlyList<WordParagraphSnapshot> paragraphs,
            ICollection<WordLineShapeSnapshot> snapshots, ISet<string> seen,
            DocumentOperationSession? operation)
        {
            var shapeCount = shapes.Count;
            for (var index = 1; index <= shapeCount; index++)
            {
                if (index == 1 || index % 25 == 0)
                    operation?.ReportProgress(index - 1, shapeCount, "đọc đường kẻ");
                Word.Shape? shape = null;
                Word.Range? anchor = null;
                Word.LineFormat? line = null;
                Word.ColorFormat? color = null;
                try
                {
                    shape = shapes[index];
                    if ((int)shape.Type != 9) continue;
                    anchor = shape.Anchor.Duplicate;
                    var key = anchor.StoryType + ":" + anchor.Start + ":" + shape.Name;
                    if (!seen.Add(key)) continue;
                    line = shape.Line;
                    var lineVisible = false;
                    int? dashStyle = null;
                    double? lineWeight = null;
                    int? lineColor = null;
                    int? beginArrowheadStyle = null;
                    int? endArrowheadStyle = null;
                    try { lineVisible = (int)line.Visible != 0; } catch (COMException) { }
                    try { dashStyle = ReadNullableInteger((int)line.DashStyle); } catch (COMException) { }
                    try { lineWeight = ReadNullableFloat(line.Weight); } catch (COMException) { }
                    try
                    {
                        color = line.ForeColor;
                        lineColor = ReadNullableInteger(color.RGB);
                    }
                    catch (COMException) { }
                    try { beginArrowheadStyle = ReadNullableInteger((int)line.BeginArrowheadStyle); }
                    catch (COMException) { }
                    try { endArrowheadStyle = ReadNullableInteger((int)line.EndArrowheadStyle); }
                    catch (COMException) { }
                    var sectionIndex = ReadSectionIndex(anchor);
                    var section = FindSection(sections, sectionIndex);
                    var anchorLeft = SafeInformationPoints(anchor, Word.WdInformation.wdHorizontalPositionRelativeToPage);
                    var anchorTop = SafeInformationPoints(anchor, Word.WdInformation.wdVerticalPositionRelativeToPage);
                    var relativeHorizontal = (int)shape.RelativeHorizontalPosition;
                    var relativeVertical = (int)shape.RelativeVerticalPosition;
                    snapshots.Add(new WordLineShapeSnapshot(
                        snapshots.Count + 1,
                        shape.Name ?? string.Empty,
                        (int)shape.Type,
                        anchor.StoryType.ToString(),
                        sectionIndex,
                        anchor.Start,
                        FindAnchorParagraph(paragraphs, anchor.StoryType.ToString(), anchor.Start),
                        SafeInformation(anchor, Word.WdInformation.wdActiveEndAdjustedPageNumber),
                        shape.Left,
                        shape.Top,
                        shape.Width,
                        shape.Height,
                        ResolvePageLeft(shape.Left, relativeHorizontal, anchorLeft, section),
                        ResolvePageTop(shape.Top, relativeVertical, anchorTop, section),
                        relativeHorizontal,
                        relativeVertical,
                        lineVisible,
                        dashStyle,
                        lineWeight,
                        lineColor,
                        beginArrowheadStyle,
                        endArrowheadStyle));
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(color);
                    Release(line);
                    Release(anchor);
                    Release(shape);
                }
            }
        }

        private static IReadOnlyList<WordTableSnapshot> CaptureTables(Word.Document document,
            DocumentOperationSession? operation)
        {
            var snapshots = new List<WordTableSnapshot>();
            Word.Tables? documentTables = null;
            try
            {
                documentTables = document.Tables;
                var tableCount = documentTables.Count;
                for (var index = 1; index <= tableCount; index++)
                {
                    operation?.ReportProgress(index - 1, tableCount, "đọc bảng");
                    Word.Table? table = null;
                    Word.Range? tableRange = null;
                    Word.Rows? rows = null;
                    Word.Columns? columns = null;
                    Word.Borders? tableBorders = null;
                    try
                    {
                        table = documentTables[index];
                        tableRange = table.Range.Duplicate;
                        var headerRows = new List<int>();
                        var rowCount = 0;
                        var columnCount = 0;
                        LocalBorderSnapshot headerSeparator = LocalBorderSnapshot.Unknown;
                        try
                        {
                            rows = table.Rows;
                            rowCount = rows.Count;
                            for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                            {
                                if (rowIndex == 1 || rowIndex % 25 == 0) operation?.Checkpoint();
                                Word.Row? row = null;
                                try
                                {
                                    row = rows[rowIndex];
                                    if (row.HeadingFormat != 0)
                                    {
                                        headerRows.Add(rowIndex);
                                        headerSeparator = ReadRowBottomBorder(row);
                                    }
                                }
                                catch (COMException) { break; }
                                finally { Release(row); }
                            }
                        }
                        catch (COMException) { }
                        try { columns = table.Columns; columnCount = columns.Count; }
                        catch (COMException) { }

                        bool? hasMergedCells = null;
                        try { hasMergedCells = !table.Uniform; }
                        catch (COMException) { }

                        tableBorders = table.Borders;
                        var top = ReadBorder(tableBorders, Word.WdBorderType.wdBorderTop);
                        var bottom = ReadBorder(tableBorders, Word.WdBorderType.wdBorderBottom);
                        var left = ReadBorder(tableBorders, Word.WdBorderType.wdBorderLeft);
                        var right = ReadBorder(tableBorders, Word.WdBorderType.wdBorderRight);
                        var insideHorizontal = ReadBorder(tableBorders, Word.WdBorderType.wdBorderHorizontal);
                        var insideVertical = ReadBorder(tableBorders, Word.WdBorderType.wdBorderVertical);
                        var hasVerticalBorders = left.State == LocalSnapshotValueState.Present ||
                            right.State == LocalSnapshotValueState.Present ||
                            insideVertical.State == LocalSnapshotValueState.Present;

                        snapshots.Add(new WordTableSnapshot(
                            index,
                            rowCount,
                            columnCount,
                            hasMergedCells.GetValueOrDefault(true),
                            table.NestingLevel > 1,
                            headerRows,
                            hasVerticalBorders,
                            tableRange.StoryType.ToString(),
                            ReadSectionIndex(tableRange),
                            tableRange.Start,
                            tableRange.End,
                            Math.Max(1, table.NestingLevel),
                            top, bottom, left, right, insideHorizontal, insideVertical,
                            headerSeparator, hasMergedCells));
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        Release(tableBorders);
                        Release(columns);
                        Release(rows);
                        Release(tableRange);
                        Release(table);
                    }
                }
            }
            finally { Release(documentTables); }

            return snapshots;
        }

        private static LocalBorderSnapshot ReadRowBottomBorder(Word.Row row)
        {
            Word.Borders? borders = null;
            try
            {
                borders = row.Borders;
                return ReadBorder(borders, Word.WdBorderType.wdBorderBottom);
            }
            catch (COMException) { return LocalBorderSnapshot.Unknown; }
            finally { Release(borders); }
        }

        private static LocalBorderSnapshot ReadBorder(Word.Borders borders, Word.WdBorderType borderType)
        {
            Word.Border? border = null;
            try
            {
                border = borders[borderType];
                var style = (int)border.LineStyle;
                if (style == (int)Word.WdLineStyle.wdLineStyleNone)
                    return LocalBorderSnapshot.None;
                return new LocalBorderSnapshot(LocalSnapshotValueState.Present, style,
                    ((int)border.LineWidth) / 8d);
            }
            catch (COMException) { return LocalBorderSnapshot.Unknown; }
            finally { Release(border); }
        }

        private static IReadOnlyList<WordProtectedSpanSnapshot> CaptureProtectedSpans(Word.Document document,
            DocumentOperationSession? operation)
        {
            var snapshots = new List<WordProtectedSpanSnapshot>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            Word.ContentControls? documentControls = null;
            try
            {
                documentControls = document.ContentControls;
                var controlCount = documentControls.Count;
                for (var index = 1; index <= controlCount; index++)
                {
                    if (index == 1 || index % 25 == 0)
                        operation?.ReportProgress(index - 1, controlCount, "đọc vùng được bảo vệ");
                    Word.ContentControl? control = null;
                    Word.Range? range = null;
                    try
                    {
                        control = documentControls[index];
                        range = control.Range.Duplicate;
                        AddProtectedSpan(snapshots, seen, range,
                            "ContentControl:" + control.Type + (control.LockContents ? ":Locked" : string.Empty));
                    }
                    finally
                    {
                        Release(range);
                        Release(control);
                    }
                }
            }
            finally { Release(documentControls); }

            foreach (var storyType in SupportedStoryTypes())
            {
                Word.Range? story = null;
                try { story = document.StoryRanges[storyType]; }
                catch (COMException) { continue; }
                while (story != null)
                {
                    Word.Range? next = null;
                    try
                    {
                        CaptureStoryProtectedSpans(story, snapshots, seen, operation);
                        next = story.NextStoryRange;
                    }
                    finally { Release(story); }
                    story = next;
                }
            }

            return snapshots;
        }

        private static void CaptureStoryProtectedSpans(Word.Range story,
            ICollection<WordProtectedSpanSnapshot> snapshots, ISet<string> seen,
            DocumentOperationSession? operation)
        {
            Word.Fields? fields = null;
            Word.OMaths? maths = null;
            Word.Hyperlinks? hyperlinks = null;
            try
            {
                fields = story.Fields;
                var count = fields.Count;
                for (var index = 1; index <= count; index++)
                {
                    operation?.Checkpoint();
                    Word.Field? field = null;
                    Word.Range? code = null;
                    Word.Range? result = null;
                    try
                    {
                        field = fields[index];
                        code = field.Code;
                        result = field.Result;
                        AddProtectedSpan(snapshots, seen, code, "FieldCode");
                        AddProtectedSpan(snapshots, seen, result, "FieldResult");
                    }
                    catch (COMException) { }
                    finally { Release(result); Release(code); Release(field); }
                }
            }
            catch (COMException) { }
            finally { Release(fields); }

            try
            {
                maths = story.OMaths;
                var count = maths.Count;
                for (var index = 1; index <= count; index++)
                {
                    Word.OMath? math = null;
                    Word.Range? range = null;
                    try
                    {
                        math = maths[index];
                        range = math.Range;
                        AddProtectedSpan(snapshots, seen, range, "OMath");
                    }
                    catch (COMException) { }
                    finally { Release(range); Release(math); }
                }
            }
            catch (COMException) { }
            finally { Release(maths); }

            try
            {
                hyperlinks = story.Hyperlinks;
                var count = hyperlinks.Count;
                for (var index = 1; index <= count; index++)
                {
                    Word.Hyperlink? hyperlink = null;
                    Word.Range? range = null;
                    try
                    {
                        hyperlink = hyperlinks[index];
                        range = hyperlink.Range;
                        AddProtectedSpan(snapshots, seen, range, "Hyperlink");
                    }
                    catch (COMException) { }
                    finally { Release(range); Release(hyperlink); }
                }
            }
            catch (COMException) { }
            finally { Release(hyperlinks); }
        }

        private static void AddProtectedSpan(ICollection<WordProtectedSpanSnapshot> snapshots,
            ISet<string> seen, Word.Range range, string kind)
        {
            var storyType = range.StoryType.ToString();
            var key = storyType + ":" + range.Start.ToString(CultureInfo.InvariantCulture) + ":" +
                range.End.ToString(CultureInfo.InvariantCulture) + ":" + kind;
            if (!seen.Add(key)) return;
            snapshots.Add(new WordProtectedSpanSnapshot(storyType, ReadSectionIndex(range), range.Start,
                Math.Max(0, range.End - range.Start), kind));
        }

        private static IReadOnlyList<WordGraphicObjectSnapshot> CaptureGraphicObjects(
            Word.Document document, IReadOnlyList<WordParagraphSnapshot> paragraphs,
            IReadOnlyList<WordProtectedSpanSnapshot> protectedSpans,
            DocumentOperationSession? operation)
        {
            var result = new List<WordGraphicObjectSnapshot>();
            Word.InlineShapes? inlineShapes = null;
            Word.Shapes? shapes = null;
            try
            {
                inlineShapes = document.InlineShapes;
                var count = inlineShapes.Count;
                for (var index = 1; index <= count; index++)
                {
                    if (index == 1 || index % 25 == 0) operation?.Checkpoint();
                    Word.InlineShape? shape = null;
                    Word.Range? range = null;
                    try
                    {
                        shape = inlineShapes[index];
                        range = shape.Range.Duplicate;
                        result.Add(new WordGraphicObjectSnapshot(result.Count + 1,
                            "InlineShape:" + shape.Type, true, range.StoryType.ToString(),
                            ReadSectionIndex(range), range.Start, range.End,
                            FindAnchorParagraph(paragraphs, range.StoryType.ToString(), range.Start),
                            IntersectsProtectedSpan(protectedSpans, range.StoryType.ToString(), range.Start, range.End)));
                    }
                    catch (COMException) { }
                    finally { Release(range); Release(shape); }
                }
            }
            finally { Release(inlineShapes); }

            try
            {
                shapes = document.Shapes;
                var count = shapes.Count;
                for (var index = 1; index <= count; index++)
                {
                    if (index == 1 || index % 25 == 0) operation?.Checkpoint();
                    Word.Shape? shape = null;
                    Word.Range? anchor = null;
                    try
                    {
                        shape = shapes[index];
                        if ((int)shape.Type == 9) continue;
                        anchor = shape.Anchor.Duplicate;
                        result.Add(new WordGraphicObjectSnapshot(result.Count + 1,
                            "Shape:" + shape.Type, false, anchor.StoryType.ToString(),
                            ReadSectionIndex(anchor), anchor.Start, anchor.End,
                            FindAnchorParagraph(paragraphs, anchor.StoryType.ToString(), anchor.Start),
                            IntersectsProtectedSpan(protectedSpans, anchor.StoryType.ToString(), anchor.Start, anchor.End)));
                    }
                    catch (COMException) { }
                    finally { Release(anchor); Release(shape); }
                }
            }
            finally { Release(shapes); }
            return result;
        }

        private static bool IntersectsProtectedSpan(
            IEnumerable<WordProtectedSpanSnapshot> protectedSpans, string storyType, int start, int end) =>
            protectedSpans.Any(span => string.Equals(span.StoryType, storyType, StringComparison.Ordinal) &&
                span.AbsoluteStart < end && span.AbsoluteStart + span.Length > start);

        private static string CaptureFingerprint(
            Word.Document document,
            IReadOnlyList<WordSectionSnapshot> sections,
            IReadOnlyList<WordParagraphSnapshot> paragraphs,
            IReadOnlyList<WordLineShapeSnapshot> lineShapes,
            IReadOnlyList<WordTableSnapshot> tables,
            IReadOnlyList<WordGraphicObjectSnapshot> graphicObjects,
            IReadOnlyList<WordProtectedSpanSnapshot> protectedSpans)
        {
            using (var sha256 = SHA256.Create())
            {
                Hash(sha256, "chuanhoa-raw-word-snapshot-v3");
                foreach (var section in sections)
                {
                    Hash(sha256, section.Index.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, section.PageWidthPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, section.PageHeightPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, section.TopMarginPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, section.BottomMarginPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, section.LeftMarginPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, section.RightMarginPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, section.IsLandscape ? "1" : "0");
                    Hash(sha256, section.HasPageNumbers ? "1" : "0");
                    Hash(sha256, section.RestartPageNumbering ? "1" : "0");
                    Hash(sha256, FormatNullable(section.StartingPageNumber));
                    Hash(sha256, FormatNullable(section.PageNumberAlignment));
                }

                foreach (var paragraph in paragraphs)
                {
                    Hash(sha256, paragraph.StoryType);
                    Hash(sha256, paragraph.SectionIndex.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, paragraph.AbsoluteStart.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, paragraph.AbsoluteEnd.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, paragraph.Text);
                    Hash(sha256, paragraph.FontName ?? string.Empty);
                    Hash(sha256, FormatNullable(paragraph.FontSizePoints));
                    Hash(sha256, FormatNullable(paragraph.Bold));
                    Hash(sha256, FormatNullable(paragraph.Italic));
                    Hash(sha256, FormatNullable(paragraph.Alignment));
                    Hash(sha256, FormatNullable(paragraph.FontColor));
                    Hash(sha256, FormatNullable(paragraph.Underline));
                    Hash(sha256, paragraph.HasBottomBorder ? "1" : "0");
                    Hash(sha256, FormatNullable(paragraph.LineSpacingPoints));
                    Hash(sha256, FormatNullable(paragraph.LineSpacingRule));
                    Hash(sha256, FormatNullable(paragraph.OutlineLevel));
                    Hash(sha256, paragraph.PageNumber.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, FormatNullable(paragraph.PageLeftPoints));
                    Hash(sha256, FormatNullable(paragraph.PageTopPoints));
                    Hash(sha256, FormatNullable(paragraph.TextWidthPoints));
                    Hash(sha256, FormatNullable(paragraph.KeepWithNext));
                    Hash(sha256, FormatNullable(paragraph.WidowControl));
                    Hash(sha256, FormatNullable(paragraph.BuiltInStyleId));
                    Hash(sha256, paragraph.StyleName ?? string.Empty);
                    Hash(sha256, FormatNullable(paragraph.TableIndex));
                    Hash(sha256, FormatNullable(paragraph.RowIndex));
                    Hash(sha256, FormatNullable(paragraph.CellIndex));
                    Hash(sha256, paragraph.TableNestingDepth.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, FormatNullable(paragraph.HasField));
                    Hash(sha256, FormatNullable(paragraph.HasMathObject));
                    Hash(sha256, FormatNullable(paragraph.HasHyperlink));
                    Hash(sha256, FormatNullable(paragraph.HasContentControl));
                    Hash(sha256, paragraph.CaptionKind);
                }

                foreach (var shape in lineShapes)
                {
                    Hash(sha256, shape.Name);
                    Hash(sha256, shape.ShapeType.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, shape.AnchorStoryType);
                    Hash(sha256, shape.AnchorSectionIndex.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, shape.AnchorAbsoluteStart.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, FormatNullable(shape.AnchorParagraphIndex));
                    Hash(sha256, shape.AnchorPageNumber.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, shape.WidthPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, shape.HeightPoints.ToString("R", CultureInfo.InvariantCulture));
                    Hash(sha256, FormatNullable(shape.PageLeftPoints));
                    Hash(sha256, FormatNullable(shape.PageTopPoints));
                    Hash(sha256, shape.LineVisible ? "1" : "0");
                    Hash(sha256, FormatNullable(shape.DashStyle));
                    Hash(sha256, FormatNullable(shape.WeightPoints));
                    Hash(sha256, FormatNullable(shape.Color));
                    Hash(sha256, FormatNullable(shape.BeginArrowheadStyle));
                    Hash(sha256, FormatNullable(shape.EndArrowheadStyle));
                }

                foreach (var table in tables)
                {
                    Hash(sha256, table.Index.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.StoryType);
                    Hash(sha256, table.SectionIndex.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.AbsoluteStart.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.AbsoluteEnd.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.NestingDepth.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.RowCount.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.ColumnCount.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, FormatNullable(table.HasMergedCellsState));
                    foreach (var headerRow in table.HeaderRowIndexes)
                        Hash(sha256, headerRow.ToString(CultureInfo.InvariantCulture));
                    HashBorder(sha256, table.TopBorder);
                    HashBorder(sha256, table.BottomBorder);
                    HashBorder(sha256, table.LeftBorder);
                    HashBorder(sha256, table.RightBorder);
                    HashBorder(sha256, table.InsideHorizontalBorder);
                    HashBorder(sha256, table.InsideVerticalBorder);
                    HashBorder(sha256, table.HeaderSeparatorBorder);
                }

                foreach (var graphic in graphicObjects)
                {
                    Hash(sha256, graphic.Index.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, graphic.ObjectKind);
                    Hash(sha256, graphic.IsInline ? "1" : "0");
                    Hash(sha256, graphic.StoryType);
                    Hash(sha256, graphic.SectionIndex.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, graphic.AbsoluteStart.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, graphic.AbsoluteEnd.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, FormatNullable(graphic.AnchorParagraphIndex));
                    Hash(sha256, graphic.IsProtected ? "1" : "0");
                }

                foreach (var span in protectedSpans)
                {
                    Hash(sha256, span.StoryType);
                    Hash(sha256, span.AbsoluteStart.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, span.Length.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, span.Kind);
                }

                sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return "sha256:" + ToHex(sha256.Hash!);
            }
        }

        private static bool ShouldCapturePageLayout(Word.WdStoryType storyType, int paragraphIndex,
            string text, int alignment)
        {
            if (storyType != Word.WdStoryType.wdMainTextStory || paragraphIndex > 150) return false;
            if (string.IsNullOrWhiteSpace(text) || text.Length > 500) return false;
            // Opening lines can be malformed before normalization, so keep a small
            // text-based fallback for the three component families that own Line
            // Shapes. Do not paginate every opening paragraph: repeated layout reads
            // for legal bases/body/list items can trigger AppHangB1 in Word.
            if (paragraphIndex <= 20 && IsOpeningLineShapeCandidate(text)) return true;
            if (alignment != (int)Word.WdParagraphAlignment.wdAlignParagraphCenter) return false;

            // All NĐ30/HD05 components requiring a rendered Line Shape are centered
            // headings near the start of the main story. Limiting pagination metrics to
            // this bounded set preserves Line Shape checks without walking every table cell.
            return true;
        }

        private static bool IsOpeningLineShapeCandidate(string text)
        {
            var normalized = (text ?? string.Empty).Trim();
            if (normalized.Length == 0 || normalized.Length > 250) return false;
            if (normalized.IndexOf("Độc lập", StringComparison.OrdinalIgnoreCase) >= 0 &&
                normalized.IndexOf("Hạnh phúc", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (normalized.StartsWith("Về việc", StringComparison.OrdinalIgnoreCase) ||
                normalized.StartsWith("Trích yếu", StringComparison.OrdinalIgnoreCase))
                return true;

            var hasLetter = false;
            foreach (var value in normalized)
            {
                if (!char.IsLetter(value)) continue;
                hasLetter = true;
                if (char.IsLower(value)) return false;
            }
            return hasLetter;
        }

        private static bool IsPageLayoutCaptureSafe(Word.Document document)
        {
            Word.Range? content = null;
            Word.Tables? tables = null;
            try
            {
                // Page-relative Information forces synchronous repagination inside
                // WINWORD. Heavy/table-rich files use anchor-based Line Shape matching
                // instead so the Read data button cannot monopolize the UI thread.
                tables = document.Tables;
                if (tables.Count > 10) return false;
                content = document.Content;
                return content.End - content.Start <= 200000;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                Release(content);
                Release(tables);
            }
        }

        private static WordTableSnapshot? FindTableAt(
            IReadOnlyList<WordTableSnapshot> tables, Word.WdStoryType storyType, int start)
        {
            var storyName = storyType.ToString();
            return tables.FirstOrDefault(item =>
                string.Equals(item.StoryType, storyName, StringComparison.Ordinal) &&
                start >= item.AbsoluteStart && start < item.AbsoluteEnd);
        }

        private static Tuple<int?, int?, int?, int> ReadTableCoordinates(
            WordTableSnapshot? table, Word.Range range)
        {
            if (table == null)
                return Tuple.Create<int?, int?, int?, int>(null, null, null, 0);

            int? row = null;
            int? cell = null;
            try { row = range.get_Information(Word.WdInformation.wdStartOfRangeRowNumber); }
            catch (COMException) { }
            try { cell = range.get_Information(Word.WdInformation.wdStartOfRangeColumnNumber); }
            catch (COMException) { }
            return Tuple.Create<int?, int?, int?, int>(table.Index, row, cell, table.NestingDepth);
        }

        private static int ResolveSectionIndex(
            IReadOnlyList<SectionBoundary> boundaries, int start, int fallback)
        {
            foreach (var boundary in boundaries)
                if (start >= boundary.Start && start < boundary.End) return boundary.Index;
            return Math.Max(1, fallback);
        }

        private static bool ShouldCaptureStyleIdentity(
            string text, bool isInTable, bool? bold, int? outlineLevel)
        {
            if (outlineLevel.HasValue && outlineLevel.Value >= 1 && outlineLevel.Value <= 9)
                return true;
            if (isInTable || string.IsNullOrWhiteSpace(text) || text.Length > 250) return false;
            if (bold == true) return true;
            var hasLetter = false;
            foreach (var value in text)
            {
                if (!char.IsLetter(value)) continue;
                hasLetter = true;
                if (char.IsLower(value)) return false;
            }
            return hasLetter;
        }

        private static bool ShouldCaptureSemanticFacts(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            if (text.IndexOf('$') >= 0) return true;
            if (text.Length > 250) return false;
            return text.IndexOf("Bảng", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("Hình", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("Table", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("Figure", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int SafeInformation(Word.Range range, Word.WdInformation information)
        {
            try
            {
                return range.get_Information(information);
            }
            catch (COMException)
            {
                return 0;
            }
        }

        private static double? SafeInformationPoints(Word.Range range, Word.WdInformation information)
        {
            try
            {
                var value = range.get_Information(information);
                return value < 0 || value >= 999999 ? (double?)null : value;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static int? FindAnchorParagraph(IReadOnlyList<WordParagraphSnapshot> paragraphs,
            string storyType, int anchorStart)
        {
            var match = paragraphs
                .Where(item => string.Equals(item.StoryType, storyType, StringComparison.Ordinal) &&
                    item.AbsoluteStart <= anchorStart)
                .OrderByDescending(item => item.AbsoluteStart)
                .FirstOrDefault();
            return match == null ? (int?)null : match.Index;
        }

        private static WordSectionSnapshot? FindSection(IReadOnlyList<WordSectionSnapshot> sections, int index)
        {
            foreach (var section in sections)
                if (section.Index == index) return section;
            return null;
        }

        private static double? ResolvePageLeft(float left, int relativePosition, double? anchorLeft,
            WordSectionSnapshot? section)
        {
            // Word persists LayoutInCell=true for shapes anchored in a layout table,
            // even after the shape is page-relative. In that state Left is already a
            // page coordinate and must not have the cell origin added a second time.
            if (relativePosition == 1) return left;
            if (relativePosition == 0 && section != null) return section.LeftMarginPoints + left;
            if ((relativePosition == 2 || relativePosition == 3) && anchorLeft.HasValue) return anchorLeft.Value + left;
            return null;
        }

        private static double? ResolvePageTop(float top, int relativePosition, double? anchorTop,
            WordSectionSnapshot? section)
        {
            if (relativePosition == 1) return top;
            if (relativePosition == 0 && section != null) return section.TopMarginPoints + top;
            if ((relativePosition == 2 || relativePosition == 3) && anchorTop.HasValue) return anchorTop.Value + top;
            return null;
        }

        private static Tuple<bool, bool, int?, int?> ReadPageNumbering(Word.Section section)
        {
            Word.HeadersFooters? footers = null;
            Word.HeadersFooters? headers = null;
            try
            {
                footers = section.Footers;
                var footerResult = TryReadPageNumbering(footers);
                if (footerResult.Item1) return footerResult;
                headers = section.Headers;
                return TryReadPageNumbering(headers);
            }
            finally
            {
                Release(headers);
                Release(footers);
            }
        }

        private static Tuple<bool, bool, int?, int?> TryReadPageNumbering(Word.HeadersFooters collection)
        {
            foreach (Word.WdHeaderFooterIndex footerIndex in new[]
            {
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages
            })
            {
                Word.HeaderFooter? footer = null;
                Word.PageNumbers? numbers = null;
                Word.PageNumber? number = null;
                try
                {
                    footer = collection[footerIndex];
                    if (!footer.Exists) continue;
                    numbers = footer.PageNumbers;
                    if (numbers.Count <= 0) continue;
                    number = numbers[1];
                    return Tuple.Create(true, numbers.RestartNumberingAtSection,
                        (int?)numbers.StartingNumber, (int?)number.Alignment);
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(number);
                    Release(numbers);
                    Release(footer);
                }
            }
            return Tuple.Create(false, false, (int?)null, (int?)null);
        }

        private static bool ReadBottomBorder(Word.ParagraphFormat format)
        {
            Word.Borders? borders = null;
            Word.Border? border = null;
            try
            {
                borders = format.Borders;
                border = borders[Word.WdBorderType.wdBorderBottom];
                return border.LineStyle != Word.WdLineStyle.wdLineStyleNone;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                Release(border);
                Release(borders);
            }
        }

        private static int ReadSectionIndex(Word.Range range)
        {
            try
            {
                return Math.Max(1, range.get_Information(Word.WdInformation.wdActiveEndSectionNumber));
            }
            catch (COMException)
            {
                return 1;
            }
        }

        private static string ReadFileFormat(Word.Document document)
        {
            var fileName = document.FullName ?? string.Empty;
            var extension = System.IO.Path.GetExtension(fileName);
            return string.IsNullOrWhiteSpace(extension)
                ? document.SaveFormat.ToString()
                : extension.ToLowerInvariant();
        }

        private static string NormalizeParagraphText(string? value)
        {
            return (value ?? string.Empty).TrimEnd('\r', '\a');
        }

        private static string? ReadFontName(Word.Font font)
        {
            var name = font.Name;
            return string.IsNullOrWhiteSpace(name) || name == ""
                ? null
                : name;
        }

        private static IReadOnlyDictionary<string, int> CaptureBuiltInStyleNames(Word.Document document)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Word.Styles? styles = null;
            try
            {
                styles = document.Styles;
                var builtIns = new[]
                {
                    Word.WdBuiltinStyle.wdStyleNormal,
                    Word.WdBuiltinStyle.wdStyleHeading1,
                    Word.WdBuiltinStyle.wdStyleHeading2,
                    Word.WdBuiltinStyle.wdStyleHeading3,
                    Word.WdBuiltinStyle.wdStyleHeading4,
                    Word.WdBuiltinStyle.wdStyleHeading5,
                    Word.WdBuiltinStyle.wdStyleHeading6,
                    Word.WdBuiltinStyle.wdStyleHeading7,
                    Word.WdBuiltinStyle.wdStyleHeading8,
                    Word.WdBuiltinStyle.wdStyleHeading9
                };
                foreach (var builtIn in builtIns)
                {
                    Word.Style? style = null;
                    try
                    {
                        style = styles[builtIn];
                        var name = style.NameLocal;
                        if (!string.IsNullOrWhiteSpace(name)) result[name] = (int)builtIn;
                    }
                    catch (COMException) { }
                    finally { Release(style); }
                }
            }
            catch (COMException)
            {
            }
            finally { Release(styles); }
            return result;
        }

        private static void ReadStyleIdentity(Word.Range range,
            IReadOnlyDictionary<string, int> builtInStyles,
            out string? styleName, out int? builtInStyleId)
        {
            styleName = null;
            builtInStyleId = null;
            object? rawStyle = null;
            Word.Style? style = null;
            try
            {
                rawStyle = range.get_Style();
                style = rawStyle as Word.Style;
                styleName = style == null ? rawStyle?.ToString() : style.NameLocal;
                int id;
                if (styleName != null && !string.IsNullOrWhiteSpace(styleName) &&
                    builtInStyles.TryGetValue(styleName, out id))
                    builtInStyleId = id;
            }
            catch (COMException) { }
            finally
            {
                if (style != null) Release(style);
                else Release(rawStyle);
            }
        }

        private static void ReadRangeSemanticFacts(Word.Range range,
            out bool? hasField, out bool? hasMathObject, out bool? hasHyperlink,
            out bool? hasContentControl, out string? captionKind)
        {
            hasField = null;
            hasMathObject = null;
            hasHyperlink = null;
            hasContentControl = null;
            captionKind = null;
            Word.Fields? fields = null;
            Word.OMaths? maths = null;
            Word.Hyperlinks? hyperlinks = null;
            Word.ContentControls? controls = null;
            try
            {
                fields = range.Fields;
                hasField = fields.Count > 0;
                if (hasField == true)
                {
                    var fieldCount = fields.Count;
                    for (var index = 1; index <= fieldCount; index++)
                    {
                        Word.Field? field = null;
                        Word.Range? code = null;
                        try
                        {
                            field = fields[index];
                            code = field.Code;
                            var fieldCode = (code.Text ?? string.Empty).Trim();
                            if (fieldCode.IndexOf("SEQ Bảng", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                fieldCode.IndexOf("SEQ Table", StringComparison.OrdinalIgnoreCase) >= 0)
                                captionKind = "Table";
                            else if (fieldCode.IndexOf("SEQ Hình", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                     fieldCode.IndexOf("SEQ Figure", StringComparison.OrdinalIgnoreCase) >= 0)
                                captionKind = "Figure";
                        }
                        catch (COMException) { }
                        finally { Release(code); Release(field); }
                    }
                }
            }
            catch (COMException) { hasField = null; }
            finally { Release(fields); }

            try { maths = range.OMaths; hasMathObject = maths.Count > 0; }
            catch (COMException) { hasMathObject = null; }
            finally { Release(maths); }
            try { hyperlinks = range.Hyperlinks; hasHyperlink = hyperlinks.Count > 0; }
            catch (COMException) { hasHyperlink = null; }
            finally { Release(hyperlinks); }
            try { controls = range.ContentControls; hasContentControl = controls.Count > 0; }
            catch (COMException) { hasContentControl = null; }
            finally { Release(controls); }
        }

        private static double? ReadNullableFloat(float value)
        {
            return value > 999998f ? (double?)null : value;
        }

        private static bool? ReadNullableBoolean(int value)
        {
            return value == 9999999 ? (bool?)null : value != 0;
        }

        private static int? ReadNullableInteger(int value)
        {
            return value == 9999999 || value == -9999999 ? (int?)null : value;
        }

        private static IEnumerable<Word.WdStoryType> SupportedStoryTypes()
        {
            yield return Word.WdStoryType.wdMainTextStory;
            yield return Word.WdStoryType.wdPrimaryHeaderStory;
            yield return Word.WdStoryType.wdFirstPageHeaderStory;
            yield return Word.WdStoryType.wdEvenPagesHeaderStory;
            yield return Word.WdStoryType.wdPrimaryFooterStory;
            yield return Word.WdStoryType.wdFirstPageFooterStory;
            yield return Word.WdStoryType.wdEvenPagesFooterStory;
            yield return Word.WdStoryType.wdFootnotesStory;
            yield return Word.WdStoryType.wdEndnotesStory;
            yield return Word.WdStoryType.wdTextFrameStory;
        }

        private static string FormatNullable<T>(T? value) where T : struct
        {
            return value.HasValue
                ? Convert.ToString(value.Value, CultureInfo.InvariantCulture) ?? string.Empty
                : "null";
        }

        private static void Hash(HashAlgorithm hash, string value)
        {
            var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            var length = BitConverter.GetBytes(bytes.Length);
            hash.TransformBlock(length, 0, length.Length, null, 0);
            hash.TransformBlock(bytes, 0, bytes.Length, null, 0);
        }

        private static void HashBorder(HashAlgorithm hash, LocalBorderSnapshot border)
        {
            Hash(hash, ((int)border.State).ToString(CultureInfo.InvariantCulture));
            Hash(hash, FormatNullable(border.LineStyle));
            Hash(hash, FormatNullable(border.WeightPoints));
        }

        private static string ToHex(byte[] bytes)
        {
            var builder = new StringBuilder(bytes.Length * 2);
            foreach (var value in bytes)
            {
                builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
