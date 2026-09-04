using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

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
            double? textWidthPoints)
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
            IReadOnlyList<int> headerRowIndexes)
        {
            Index = index;
            RowCount = rowCount;
            ColumnCount = columnCount;
            HasMergedCells = hasMergedCells;
            IsNested = isNested;
            HeaderRowIndexes = headerRowIndexes;
        }

        public int Index { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        public bool HasMergedCells { get; }
        public bool IsNested { get; }
        public IReadOnlyList<int> HeaderRowIndexes { get; }
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
        private const int SnapshotSchemaVersion = 2;

        public WordDocumentSnapshot Build(
            Word.Document document,
            DocumentContext context,
            WordDocumentCapability capability)
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

            var sections = CaptureSections(document);
            var allowPageLayout = IsPageLayoutCaptureSafe(document);
            var paragraphs = CaptureParagraphs(document, allowPageLayout);
            var lineShapes = CaptureLineShapes(document, sections, paragraphs);
            var tables = CaptureTables(document);
            var protectedSpans = CaptureProtectedSpans(document);
            var fingerprint = CaptureFingerprint(document, sections, paragraphs, lineShapes, tables, protectedSpans);
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
                protectedSpans);
        }

        private static IReadOnlyList<WordSectionSnapshot> CaptureSections(Word.Document document)
        {
            var snapshots = new List<WordSectionSnapshot>();
            for (var index = 1; index <= document.Sections.Count; index++)
            {
                Word.Section? section = null;
                Word.PageSetup? setup = null;
                try
                {
                    section = document.Sections[index];
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

            return snapshots;
        }

        private static IReadOnlyList<WordParagraphSnapshot> CaptureParagraphs(Word.Document document,
            bool allowPageLayout)
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
                                allowPageLayout, application);
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
            }

            return snapshots;
        }

        private static void CaptureStoryParagraphs(
            Word.Range story,
            Word.WdStoryType storyType,
            ICollection<WordParagraphSnapshot> snapshots,
            ref int paragraphIndex,
            bool allowPageLayout,
            Word.Application application)
        {
            var paragraphs = story.Paragraphs;
            try
            {
                // Cache Count once. Re-reading this COM property for every item is
                // disproportionately expensive for stories containing large tables.
                var paragraphCount = paragraphs.Count;
                var largeStory = paragraphCount > 400;
                if (largeStory)
                {
                    CaptureLargeStoryParagraphs(story, storyType, snapshots, ref paragraphIndex,
                        application, paragraphCount);
                    return;
                }
                var currentSectionIndex = ReadSectionIndex(story);
                var storyParagraphIndex = 0;
                foreach (Word.Paragraph paragraphItem in paragraphs)
                {
                    storyParagraphIndex++;
                    if (storyParagraphIndex == 1 || storyParagraphIndex % 25 == 0)
                        PulseWordUi(application, paragraphIndex, paragraphCount, storyType);
                    Word.Paragraph? paragraph = paragraphItem;
                    Word.Range? range = null;
                    Word.Font? font = null;
                    Word.ParagraphFormat? format = null;
                    try
                    {
                        range = paragraph.Range.Duplicate;
                        paragraphIndex++;
                        var text = NormalizeParagraphText(range.Text);
                        var withInTable = range.get_Information(Word.WdInformation.wdWithInTable);
                        var tableCoordinates = ReadTableCoordinates(range, withInTable && !largeStory);
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
                            hasBottomBorder = ReadBottomBorder(format);
                            lineSpacing = ReadNullableFloat(format.LineSpacing);
                            lineSpacingRule = ReadNullableInteger((int)format.LineSpacingRule);
                            outlineLevel = ReadNullableInteger((int)format.OutlineLevel);
                        }
                        if (captureFormatting || storyParagraphIndex == 1 || storyParagraphIndex % 100 == 0)
                            currentSectionIndex = ReadSectionIndex(range);
                        double? textWidth = null;
                        double? pageLeft = null;
                        double? pageTop = null;
                        var pageNumber = 0;
                        if (allowPageLayout && paragraphCount <= 400 &&
                            ShouldCapturePageLayout(storyType, paragraphIndex, text, alignment.GetValueOrDefault(-1)))
                        {
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
                        }
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
                            range.Start,
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
                            textWidth));
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
        }

        private static void CaptureLargeStoryParagraphs(Word.Range story, Word.WdStoryType storyType,
            ICollection<WordParagraphSnapshot> snapshots, ref int paragraphIndex,
            Word.Application application, int expectedParagraphCount)
        {
            var tableSpans = CaptureTableSpans(story);
            var storyParagraphIndex = 0;
            var currentSectionIndex = ReadSectionIndex(story);
            Word.Paragraphs? paragraphs = null;
            try
            {
                paragraphs = story.Paragraphs;
                foreach (Word.Paragraph paragraphItem in paragraphs)
                {
                    Word.Paragraph? paragraph = paragraphItem;
                    Word.Range? range = null;
                    Word.Font? font = null;
                    Word.ParagraphFormat? format = null;
                    try
                    {
                        storyParagraphIndex++;
                        paragraphIndex++;
                        if (storyParagraphIndex == 1 || storyParagraphIndex % 25 == 0)
                            PulseWordUi(application, paragraphIndex, expectedParagraphCount, storyType);

                        // Range.Text cannot be used as a coordinate map for tables.
                        // A cell terminator is returned as CR + BEL (two chars), while
                        // Word's Start/End coordinate advances by either one or two
                        // positions depending on the table structure. Read the COM
                        // paragraph range itself so every later annotation has the
                        // authoritative Word coordinate.
                        range = paragraph.Range.Duplicate;
                        var absoluteStart = range.Start;
                        var text = NormalizeParagraphText(range.Text);
                        var isInTable = tableSpans.Any(span =>
                            absoluteStart >= span.Item1 && absoluteStart < span.Item2);
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
                            hasBottomBorder = ReadBottomBorder(format);
                            lineSpacing = ReadNullableFloat(format.LineSpacing);
                            lineSpacingRule = ReadNullableInteger((int)format.LineSpacingRule);
                            outlineLevel = ReadNullableInteger((int)format.OutlineLevel);
                        }
                        if (captureFormatting || storyParagraphIndex == 1 || storyParagraphIndex % 100 == 0)
                            currentSectionIndex = ReadSectionIndex(range);

                        snapshots.Add(new WordParagraphSnapshot(paragraphIndex, text, "Unknown", 0.0d,
                            fontName, fontSize, bold, italic, alignment, firstLineIndent, spaceBefore,
                            spaceAfter, isInTable, storyType.ToString(), currentSectionIndex, absoluteStart,
                            null, null, null, fontColor, underline, hasBottomBorder, lineSpacing,
                            lineSpacingRule, outlineLevel, 0, null, null, null));
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
                foreach (Word.Table table in tables)
                {
                    Word.Range? range = null;
                    try
                    {
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
            int storyParagraphs, Word.WdStoryType storyType)
        {
            try
            {
                application.StatusBar = string.Format(CultureInfo.CurrentCulture,
                    "Chuẩn hóa đang đọc dữ liệu: {0:N0} đoạn (vùng {1}, {2:N0} đoạn)...",
                    completedParagraphs, storyType, storyParagraphs);
            }
            catch (COMException)
            {
            }

            // Snapshot capture runs on Word's STA thread. Yielding its Windows message
            // queue at bounded intervals keeps the host responsive on long table stories;
            // RibbonRuntime's operation guard rejects re-entrant add-in commands.
            System.Windows.Forms.Application.DoEvents();
        }

        private static IReadOnlyList<WordLineShapeSnapshot> CaptureLineShapes(
            Word.Document document,
            IReadOnlyList<WordSectionSnapshot> sections,
            IReadOnlyList<WordParagraphSnapshot> paragraphs)
        {
            var snapshots = new List<WordLineShapeSnapshot>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            Word.Shapes? documentShapes = null;
            try
            {
                documentShapes = document.Shapes;
                CaptureShapeCollection(documentShapes, sections, paragraphs, snapshots, seen);
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(documentShapes);
            }

            for (var sectionIndex = 1; sectionIndex <= document.Sections.Count; sectionIndex++)
            {
                Word.Section? section = null;
                try
                {
                    section = document.Sections[sectionIndex];
                    CaptureHeaderFooterShapes(section.Headers, sections, paragraphs, snapshots, seen);
                    CaptureHeaderFooterShapes(section.Footers, sections, paragraphs, snapshots, seen);
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(section);
                }
            }

            return snapshots;
        }

        private static void CaptureHeaderFooterShapes(Word.HeadersFooters collection,
            IReadOnlyList<WordSectionSnapshot> sections, IReadOnlyList<WordParagraphSnapshot> paragraphs,
            ICollection<WordLineShapeSnapshot> snapshots, ISet<string> seen)
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
                    Word.HeaderFooter? headerFooter = null;
                    Word.Shapes? shapes = null;
                    try
                    {
                        headerFooter = collection[index];
                        if (!headerFooter.Exists) continue;
                        shapes = headerFooter.Shapes;
                        CaptureShapeCollection(shapes, sections, paragraphs, snapshots, seen);
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
            ICollection<WordLineShapeSnapshot> snapshots, ISet<string> seen)
        {
            for (var index = 1; index <= shapes.Count; index++)
            {
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
                    color = line.ForeColor;
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
                        (int)line.Visible != 0,
                        ReadNullableInteger((int)line.DashStyle),
                        ReadNullableFloat(line.Weight),
                        ReadNullableInteger(color.RGB),
                        ReadNullableInteger((int)line.BeginArrowheadStyle),
                        ReadNullableInteger((int)line.EndArrowheadStyle)));
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

        private static IReadOnlyList<WordTableSnapshot> CaptureTables(Word.Document document)
        {
            var snapshots = new List<WordTableSnapshot>();
            for (var index = 1; index <= document.Tables.Count; index++)
            {
                Word.Table? table = null;
                try
                {
                    table = document.Tables[index];
                    var headerRows = new List<int>();
                    for (var rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
                    {
                        Word.Row? row = null;
                        try
                        {
                            row = table.Rows[rowIndex];
                            if (row.HeadingFormat != 0)
                            {
                                headerRows.Add(rowIndex);
                            }
                        }
                        catch (COMException)
                        {
                            break;
                        }
                        finally
                        {
                            Release(row);
                        }
                    }

                    var hasMergedCells = false;
                    try
                    {
                        hasMergedCells = !table.Uniform;
                    }
                    catch (COMException)
                    {
                        hasMergedCells = true;
                    }

                    snapshots.Add(new WordTableSnapshot(
                        index,
                        table.Rows.Count,
                        table.Columns.Count,
                        hasMergedCells,
                        table.NestingLevel > 1,
                        headerRows));
                }
                finally
                {
                    Release(table);
                }
            }

            return snapshots;
        }

        private static IReadOnlyList<WordProtectedSpanSnapshot> CaptureProtectedSpans(Word.Document document)
        {
            var snapshots = new List<WordProtectedSpanSnapshot>();
            for (var index = 1; index <= document.ContentControls.Count; index++)
            {
                Word.ContentControl? control = null;
                Word.Range? range = null;
                try
                {
                    control = document.ContentControls[index];
                    if (control.LockContents)
                    {
                        range = control.Range.Duplicate;
                        snapshots.Add(new WordProtectedSpanSnapshot(
                            range.StoryType.ToString(),
                            ReadSectionIndex(range),
                            range.Start,
                            Math.Max(0, range.End - range.Start),
                            "ContentControl:" + control.Type));
                    }
                }
                finally
                {
                    Release(range);
                    Release(control);
                }
            }

            return snapshots;
        }

        private static string CaptureFingerprint(
            Word.Document document,
            IReadOnlyList<WordSectionSnapshot> sections,
            IReadOnlyList<WordParagraphSnapshot> paragraphs,
            IReadOnlyList<WordLineShapeSnapshot> lineShapes,
            IReadOnlyList<WordTableSnapshot> tables,
            IReadOnlyList<WordProtectedSpanSnapshot> protectedSpans)
        {
            using (var sha256 = SHA256.Create())
            {
                Hash(sha256, "chuanhoa-document-snapshot-v1");
                Hash(sha256, document.FullName ?? string.Empty);
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
                    Hash(sha256, table.RowCount.ToString(CultureInfo.InvariantCulture));
                    Hash(sha256, table.ColumnCount.ToString(CultureInfo.InvariantCulture));
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
            // The first 40 paragraphs include all header metadata (organ name, motto, code, subject)
            // even if their alignment is left, justify, or not yet normalized to center.
            if (paragraphIndex <= 40) return true;
            if (alignment != (int)Word.WdParagraphAlignment.wdAlignParagraphCenter) return false;

            // All NĐ30/HD05 components requiring a rendered Line Shape are centered
            // headings near the start of the main story. Limiting pagination metrics to
            // this bounded set preserves Line Shape checks without walking every table cell.
            return true;
        }

        private static bool IsPageLayoutCaptureSafe(Word.Document document)
        {
            Word.Range? content = null;
            try
            {
                // Page-relative Information forces synchronous repagination inside
                // WINWORD. Heavy/table-rich files use anchor-based Line Shape matching
                // instead so the Read data button cannot monopolize the UI thread.
                if (document.Tables.Count > 10) return false;
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
            }
        }

        private static Tuple<int?, int?, int?> ReadTableCoordinates(Word.Range range, bool withInTable)
        {
            if (!withInTable)
            {
                return Tuple.Create<int?, int?, int?>(null, null, null);
            }

            return Tuple.Create<int?, int?, int?>(
                null,
                range.get_Information(Word.WdInformation.wdStartOfRangeRowNumber),
                range.get_Information(Word.WdInformation.wdStartOfRangeColumnNumber));
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
