using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;
using QRCoder;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// Small, deterministic Word operations that are executed locally after a signed
    /// offline lease and rule pack have been validated. Only document-wide tone-placement
    /// commands create a recovery copy; local edits rely on Word's undo record.
    /// </summary>
    public sealed class WordLocalCommandRuntime
    {
        private const float PointsPerMillimeter = 72.0f / 25.4f;
        private readonly Word.Application _application;
        private readonly WordDocumentCapabilityProvider _capabilityProvider;
        private readonly LocalAccessManager _accessManager;
        private readonly Func<Word.Document?>? _documentProvider;
        private readonly DocumentRoleDetector _roleDetector = new DocumentRoleDetector();

        public WordLocalCommandRuntime(Word.Application application, LocalAccessManager accessManager)
            : this(application, accessManager, null)
        {
        }

        internal WordLocalCommandRuntime(
            Word.Application application,
            LocalAccessManager accessManager,
            Func<Word.Document?>? documentProvider)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _capabilityProvider = new WordDocumentCapabilityProvider(application);
            _accessManager = accessManager ?? throw new ArgumentNullException(nameof(accessManager));
            _documentProvider = documentProvider;
        }

        public string FormatPage()
        {
            return Execute("Định dạng trang giấy", (document, rules) =>
            {
                foreach (Word.Section section in document.Sections)
                {
                    try
                    {
                        section.PageSetup.PageWidth = (float)(rules.A4WidthMm * PointsPerMillimeter);
                        section.PageSetup.PageHeight = (float)(rules.A4HeightMm * PointsPerMillimeter);
                        section.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                        section.PageSetup.TopMargin = (float)(rules.TopMinMm * PointsPerMillimeter);
                        section.PageSetup.BottomMargin = (float)(rules.BottomMinMm * PointsPerMillimeter);
                        section.PageSetup.LeftMargin = (float)(rules.LeftMinMm * PointsPerMillimeter);
                        section.PageSetup.RightMargin = (float)(rules.RightMinMm * PointsPerMillimeter);
                    }
                    finally { Release(section); }
                }
            });
        }

        public string InsertSection(bool landscape)
        {
            return Execute(landscape ? "Chèn trang ngang" : "Chèn trang dọc", (document, rules) =>
            {
                var selection = _application.Selection;
                selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                var section = selection.Sections[1];
                try
                {
                    section.PageSetup.PageWidth = (float)(rules.A4WidthMm * PointsPerMillimeter);
                    section.PageSetup.PageHeight = (float)(rules.A4HeightMm * PointsPerMillimeter);
                    section.PageSetup.Orientation = landscape
                        ? Word.WdOrientation.wdOrientLandscape
                        : Word.WdOrientation.wdOrientPortrait;
                    section.PageSetup.TopMargin = (float)(rules.TopMinMm * PointsPerMillimeter);
                    section.PageSetup.BottomMargin = (float)(rules.BottomMinMm * PointsPerMillimeter);
                    section.PageSetup.LeftMargin = (float)(rules.LeftMinMm * PointsPerMillimeter);
                    section.PageSetup.RightMargin = (float)(rules.RightMinMm * PointsPerMillimeter);
                }
                finally { Release(section); }
            });
        }

        public string RemoveTrailingBlankParagraphs()
        {
            return Execute("Xóa trang thừa", (document, _) =>
            {
                WordTrailingBlankPageCleaner.Remove(document);
            });
        }

        public string KeepWithNext()
        {
            return Execute("Keep with next", (_, __) =>
                _application.Selection.ParagraphFormat.KeepWithNext = -1);
        }

        public string SetCharacterSpacing(float delta, bool reset)
        {
            try
            {
                var currentSelection = _application.Selection;
                if (currentSelection == null || currentSelection.Start == currentSelection.End)
                    throw new InvalidOperationException("Hãy chọn phần chữ cần co hoặc giãn.");
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException("Hãy mở tài liệu và chọn phần chữ cần co hoặc giãn.", exception);
            }

            return Execute(reset ? "Đưa giãn chữ về bình thường" : delta < 0 ? "Co chữ" : "Giãn chữ", (_, __) =>
            {
                var selection = _application.Selection;
                var range = selection.Range.Duplicate;
                try
                {
                    var current = range.Font.Spacing;
                    if (current > 999998f) current = 0f;
                    range.Font.Spacing = reset ? 0f : Math.Max(-5f, Math.Min(20f, current + delta));
                }
                finally { Release(range); }
            });
        }

        public string RepeatTableHeaders()
        {
            return Execute("Lặp tiêu đề bảng", (document, _) =>
            {
                Word.Range? originalSelection = null;
                try
                {
                    originalSelection = _application.Selection.Range.Duplicate;
                    foreach (Word.Table table in document.Tables)
                    {
                        try
                        {
                            if (table.NestingLevel != 1) continue;
                            RepeatDetectedHeaderRows(document, table);
                        }
                        finally { Release(table); }
                    }
                }
                finally
                {
                    if (originalSelection != null) originalSelection.Select();
                    Release(originalSelection);
                }
            });
        }

        private void RepeatDetectedHeaderRows(Word.Document document, Word.Table table)
        {
            var rows = CaptureTableRowStats(table);
            if (rows.Count < 2) return;

            var maximumRow = rows.Keys.Max();
            if (maximumRow < 2) return;
            var headerRowCount = 1;
            var maximumCandidate = Math.Min(maximumRow - 1, 6);
            for (var rowIndex = 2; rowIndex <= maximumCandidate; rowIndex++)
            {
                TableRowStats stats;
                if (!rows.TryGetValue(rowIndex, out stats) || stats.NonEmptyCells == 0) break;
                if (stats.AllCellsAlreadyHeading || stats.BoldCells * 2 >= stats.NonEmptyCells)
                    headerRowCount = rowIndex;
                else
                    break;
            }

            var alreadyApplied = rows.Where(item => item.Key <= headerRowCount)
                .All(item => item.Value.AllCellsAlreadyHeading);
            if (alreadyApplied) return;

            TableRowStats firstDataRow;
            if (!rows.TryGetValue(headerRowCount + 1, out firstDataRow)) return;
            Word.Range? tableRange = null;
            Word.Range? headerRange = null;
            try
            {
                tableRange = table.Range;
                headerRange = document.Range(tableRange.Start, firstDataRow.FirstCellStart);
                headerRange.Select();

                // Setting table.Rows[1].HeadingFormat throws when a header contains
                // vertical merges. Word's native command accepts a complete multi-row
                // selection and correctly marks every visual header row.
                _application.CommandBars.ExecuteMso("TableRepeatHeaderRows");
            }
            finally
            {
                Release(headerRange);
                Release(tableRange);
            }

            var verified = CaptureTableRowStats(table);
            if (verified.Where(item => item.Key <= headerRowCount)
                .Any(item => !item.Value.AllCellsAlreadyHeading))
                throw new InvalidOperationException(
                    "Word không thể đặt lặp đầy đủ " + headerRowCount.ToString(CultureInfo.InvariantCulture) +
                    " dòng tiêu đề của một bảng có ô gộp.");
        }

        private static Dictionary<int, TableRowStats> CaptureTableRowStats(Word.Table table)
        {
            var result = new Dictionary<int, TableRowStats>();
            Word.Range? tableRange = null;
            Word.Cells? cells = null;
            try
            {
                tableRange = table.Range;
                cells = tableRange.Cells;
                for (var cellIndex = 1; cellIndex <= cells.Count; cellIndex++)
                {
                    Word.Cell? cell = null;
                    Word.Range? range = null;
                    Word.Font? font = null;
                    Word.Rows? cellRows = null;
                    try
                    {
                        cell = cells[cellIndex];
                        range = cell.Range.Duplicate;
                        font = range.Font;
                        cellRows = range.Rows;
                        var rowIndex = cell.RowIndex;
                        TableRowStats stats;
                        if (!result.TryGetValue(rowIndex, out stats))
                        {
                            stats = new TableRowStats(range.Start);
                            result.Add(rowIndex, stats);
                        }
                        stats.FirstCellStart = Math.Min(stats.FirstCellStart, range.Start);
                        stats.TotalCells++;
                        var text = (range.Text ?? string.Empty).Trim('\r', '\a', ' ', '\t');
                        if (text.Length > 0)
                        {
                            stats.NonEmptyCells++;
                            if (font.Bold == -1) stats.BoldCells++;
                        }
                        if (cellRows.HeadingFormat != 0) stats.HeadingCells++;
                    }
                    finally
                    {
                        Release(cellRows);
                        Release(font);
                        Release(range);
                        Release(cell);
                    }
                }
                return result;
            }
            finally
            {
                Release(cells);
                Release(tableRange);
            }
        }

        public string InsertPageNumbers()
        {
            return Execute("Chèn số trang", (document, rules) =>
            {
                foreach (Word.Section section in document.Sections)
                {
                    Word.HeaderFooter? header = null;
                    Word.Range? range = null;
                    Word.Field? field = null;
                    try
                    {
                        var bookmarkName = "CHUANHOA_PAGE_NUMBER_S" + section.Index.ToString(CultureInfo.InvariantCulture);
                        if (document.Bookmarks.Exists(bookmarkName)) continue;
                        header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                        if (!header.Exists) header.Exists = true;
                        if (section.Index > 1 && header.LinkToPrevious &&
                            document.Bookmarks.Exists("CHUANHOA_PAGE_NUMBER_S" +
                                (section.Index - 1).ToString(CultureInfo.InvariantCulture)))
                            continue;
                        section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                        range = header.Range.Duplicate;
                        range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        field = document.Fields.Add(range, Word.WdFieldType.wdFieldPage, PreserveFormatting: true);
                        field.Result.Font.Name = rules.BodyFontName;
                        field.Result.Font.Size = 13f;
                        field.Result.Font.Bold = 0;
                        field.Result.Font.Italic = 0;
                        document.Bookmarks.Add(bookmarkName, field.Result);
                    }
                    finally
                    {
                        Release(field);
                        Release(range);
                        Release(header);
                        Release(section);
                    }
                }
            });
        }

        public string CenterTables()
        {
            return Execute("Căn giữa bảng", (document, _) =>
            {
                foreach (Word.Table table in document.Tables)
                {
                    try
                    {
                        if (IsHeaderLayoutTable(table))
                        {
                            NormalizeHeaderTable(table);
                            continue;
                        }
                        table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                    }
                    catch (COMException) { /* Continue with the remaining independent tables. */ }
                    finally { Release(table); }
                }
            });
        }

        private static bool IsHeaderLayoutTable(Word.Table table)
        {
            Word.Range? range = null;
            try
            {
                range = table.Range;
                var text = range.Text ?? string.Empty;
                return text.IndexOf("CỘNG HÒA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("CỘNG HOÀ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("ĐẢNG CỘNG SẢN", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (text.IndexOf("Độc lập", StringComparison.OrdinalIgnoreCase) >= 0 &&
                     text.IndexOf("Hạnh phúc", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch (COMException)
            {
                return false;
            }
            finally { Release(range); }
        }

        private static void NormalizeHeaderTable(Word.Table table)
        {
            Word.Range? tableRange = null;
            Word.Sections? sections = null;
            Word.Section? section = null;
            try
            {
                tableRange = table.Range;
                sections = tableRange.Sections;
                section = sections.Count > 0 ? sections[1] : null;
                var pageWidth = section != null ? section.PageSetup.PageWidth : 595.3f;
                var leftMargin = section != null ? section.PageSetup.LeftMargin : 85.05f;
                var rightMargin = section != null ? section.PageSetup.RightMargin : 42.5f;
                var availableWidth = Math.Max(200f, pageWidth - leftMargin - rightMargin);

                table.Rows.LeftIndent = 0f;
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                try { table.Borders.Enable = 0; } catch (COMException) { }
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = availableWidth;

                var col1Width = (float)Math.Round(availableWidth * 0.38d, 1);
                var col2Width = (float)Math.Round(availableWidth - col1Width, 1);

                for (var r = 1; r <= table.Rows.Count; r++)
                {
                    Word.Row? row = null;
                    try
                    {
                        row = table.Rows[r];
                        row.LeftIndent = 0f;
                        row.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                        if (row.Cells.Count == 2)
                        {
                            row.Cells[1].Width = col1Width;
                            row.Cells[2].Width = col2Width;
                        }
                    }
                    catch (COMException) { }
                    finally { Release(row); }
                }

                try
                {
                    if (table.Uniform && table.Columns.Count == 2)
                    {
                        table.Columns[1].SetWidth(col1Width, Word.WdRulerStyle.wdAdjustNone);
                        table.Columns[2].SetWidth(col2Width, Word.WdRulerStyle.wdAdjustNone);
                    }
                }
                catch (COMException) { }
            }
            catch (COMException) { }
            finally
            {
                Release(section);
                Release(sections);
                Release(tableRange);
            }
        }

        public string CenterImages()
        {
            return Execute("Căn giữa ảnh", (document, _) =>
            {
                foreach (Word.InlineShape shape in document.InlineShapes)
                {
                    Word.Range? range = null;
                    try
                    {
                        range = shape.Range.Duplicate;
                        range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        var section = range.Sections.Count > 0 ? range.Sections[1] : document.Sections[1];
                        try
                        {
                            var maximumWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
                            if (maximumWidth > 0f && shape.Width > maximumWidth)
                            {
                                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                                shape.Width = maximumWidth;
                            }
                        }
                        finally { Release(section); }
                    }
                    finally { Release(range); Release(shape); }
                }
            });
        }

        public string AlignCurrentCells(bool middle)
        {
            return Execute(middle ? "Căn giữa ô" : "Căn đỉnh ô", (_, __) =>
            {
                if (!_application.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    throw new InvalidOperationException("Hãy đặt con trỏ trong bảng hoặc chọn các ô cần căn.");
                foreach (Word.Cell cell in _application.Selection.Cells)
                {
                    try
                    {
                        cell.VerticalAlignment = middle
                            ? Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            : Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    }
                    finally { Release(cell); }
                }
            });
        }

        public string CleanExcelTableCharacters()
        {
            return Execute("Xóa ký tự thừa trong bảng", (_, __) =>
            {
                if (!_application.Selection.get_Information(Word.WdInformation.wdWithInTable))
                    throw new InvalidOperationException("Hãy đặt con trỏ trong bảng được dán từ Excel.");
                Word.Table? table = null;
                Word.Range? range = null;
                try
                {
                    table = _application.Selection.Tables[1];
                    range = table.Range.Duplicate;
                    ReplaceAll(range, "^l", " ");
                    ReplaceAll(range, "^t", " ");
                    ReplaceAll(range, "  ", " ");
                }
                finally { Release(range); Release(table); }
            });
        }

        public string ConvertDecimalSeparators()
        {
            return Execute("Đổi dấu phân cách số", (_, __) =>
            {
                var range = MutableSelectionOrParagraph();
                try
                {
                    var source = range.Text ?? string.Empty;
                    var converted = Regex.Replace(source, @"(?<!\d)(\d{1,3}(?:,\d{3})+)\.(\d+)(?!\d)", match =>
                        match.Groups[1].Value.Replace(',', '.') + "," + match.Groups[2].Value,
                        RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250));
                    if (!string.Equals(source, converted, StringComparison.Ordinal)) range.Text = converted;
                }
                finally { Release(range); }
            });
        }

        public string CleanWhitespaceAndPunctuation()
        {
            return Execute("Dọn khoảng trắng & Dấu câu", (document, _) =>
            {
                var selection = _application.Selection;
                if (selection.Start != selection.End)
                {
                    Word.Range? range = null;
                    try
                    {
                        range = selection.Range.Duplicate;
                        CleanWhitespaceInRange(range);
                    }
                    finally { Release(range); }
                }
                else
                {
                    foreach (var story in EditableStories(document))
                    {
                        try { CleanWhitespaceInStory(story); }
                        finally { Release(story); }
                    }
                }
            });
        }

        public string NormalizeQuotationMarks()
        {
            return Execute("Chuẩn hóa dấu ngoặc", (document, _) =>
            {
                var selection = _application.Selection;
                if (selection.Start != selection.End)
                {
                    Word.Range? range = null;
                    try
                    {
                        range = selection.Range.Duplicate;
                        NormalizeQuotationMarksInRange(range);
                    }
                    finally { Release(range); }
                }
                else
                {
                    foreach (var story in EditableStories(document))
                    {
                        try { NormalizeQuotationMarksInStory(story); }
                        finally { Release(story); }
                    }
                }
            });
        }

        public void OpenCustomDictionaryDialog()
        {
            CustomDictionaryDialog.Prompt(null);
        }

        private static void CleanWhitespaceInStory(Word.Range story)
        {
            foreach (Word.Paragraph paragraph in story.Paragraphs)
            {
                Word.Range? range = null;
                try
                {
                    range = paragraph.Range.Duplicate;
                    CleanWhitespaceInRange(range);
                }
                finally
                {
                    Release(range);
                    Release(paragraph);
                }
            }
        }

        private static void CleanWhitespaceInRange(Word.Range range)
        {
            var source = range.Text ?? string.Empty;
            var cleaned = VietnameseTypographyCleaner.CleanWhitespaceAndPunctuation(source);
            if (!string.Equals(source, cleaned, StringComparison.Ordinal))
            {
                range.Text = cleaned;
            }
        }

        private static void NormalizeQuotationMarksInStory(Word.Range story)
        {
            foreach (Word.Paragraph paragraph in story.Paragraphs)
            {
                Word.Range? range = null;
                try
                {
                    range = paragraph.Range.Duplicate;
                    NormalizeQuotationMarksInRange(range);
                }
                finally
                {
                    Release(range);
                    Release(paragraph);
                }
            }
        }

        private static void NormalizeQuotationMarksInRange(Word.Range range)
        {
            var source = range.Text ?? string.Empty;
            var cleaned = VietnameseTypographyCleaner.NormalizeQuotationMarks(source);
            if (!string.Equals(source, cleaned, StringComparison.Ordinal))
            {
                range.Text = cleaned;
            }
        }

        public string ConvertLegacyEncodingToUnicode()
        {
            return Execute("Chuyển đổi Unicode", (document, _) =>
            {
                var convertedRuns = 0;
                foreach (var story in EditableStories(document))
                {
                    try { convertedRuns += ConvertLegacyRuns(story); }
                    finally { Release(story); }
                }
                if (convertedRuns == 0)
                    throw new InvalidOperationException("Không tìm thấy đoạn chữ dùng font TCVN3 (.Vn...) hoặc VNI (VNI-...).");
            });
        }

        public string BuildStyleSet(DocumentContext context, float mainSize)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            context.RequireSnapshotAnalysis();
            return Execute("Dựng bộ Style cỡ " + mainSize.ToString("0", CultureInfo.InvariantCulture), (document, rules) =>
            {
                ConfigureBuiltInStyles(document, mainSize, rules.BodyFontName);
                if (context.LastLocalSnapshot!.Paragraphs.Any(item => !string.IsNullOrWhiteSpace(item.Text)))
                    ApplyFontSizeSetCore(document, context, mainSize, rules.BodyFontName);
            });
        }

        public string ApplyFontSizeSet(DocumentContext context, float mainSize)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            context.RequireSnapshotAnalysis();
            return Execute("Bộ cỡ chữ " + mainSize.ToString("0", CultureInfo.InvariantCulture),
                (document, rules) => ApplyFontSizeSetCore(document, context, mainSize, rules.BodyFontName));
        }

        public string NormalizeTonePlacement(VietnameseTonePlacementStyle style)
        {
            var title = style == VietnameseTonePlacementStyle.MainVowel ? "Kiểu oà, uý" : "Kiểu òa, úy";
            return Execute(title, (document, _) =>
            {
                foreach (var story in EditableStories(document))
                {
                    try { NormalizeToneInStory(story, style); }
                    finally { Release(story); }
                }
            }, createBackup: true);
        }

        public string InsertQrCode(string content)
        {
            if (string.IsNullOrWhiteSpace(content)) throw new ArgumentException("Nội dung mã QR không được để trống.", nameof(content));
            if (content.Length > 800) throw new ArgumentException("Nội dung mã QR không được vượt quá 800 ký tự.", nameof(content));
            return Execute("Chèn mã QR", (document, _) =>
            {
                var temporaryPath = Path.Combine(Path.GetTempPath(), "chuanhoa-qr-" + Guid.NewGuid().ToString("N") + ".png");
                Word.Range? insertionRange = null;
                Word.InlineShape? shape = null;
                try
                {
                    using (var generator = new QRCodeGenerator())
                    using (var data = generator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M))
                    using (var qr = new PngByteQRCode(data))
                        File.WriteAllBytes(temporaryPath, qr.GetGraphic(20));

                    insertionRange = _application.Selection.Range.Duplicate;
                    insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    shape = document.InlineShapes.AddPicture(temporaryPath, false, true, insertionRange);
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.Width = 50f * PointsPerMillimeter;
                    shape.Height = 50f * PointsPerMillimeter;
                }
                finally
                {
                    Release(shape);
                    Release(insertionRange);
                    try { if (File.Exists(temporaryPath)) File.Delete(temporaryPath); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            });
        }

        private void ApplyFontSizeSetCore(Word.Document document, DocumentContext context, float mainSize, string fontName)
        {
            if (mainSize != 13f && mainSize != 14f && mainSize != 15f)
                throw new ArgumentOutOfRangeException(nameof(mainSize));
            var local = context.LastLocalSnapshot!;
            var roles = _roleDetector.Detect(local);
            var sizes = FontSizeProfile.For(mainSize);

            foreach (var story in EditableStories(document))
            {
                try { story.Font.Name = fontName; }
                catch (COMException) { }
                finally { Release(story); }
            }

            foreach (var paragraph in local.Paragraphs.Where(p =>
                string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal)))
            {
                Word.Range? range = null;
                try
                {
                    var end = Math.Min(document.Content.End, paragraph.AbsoluteStart + paragraph.Text.Length + 1);
                    range = document.Range(paragraph.AbsoluteStart, Math.Max(paragraph.AbsoluteStart, end));
                    string role;
                    roles.TryGetValue(paragraph.Index, out role!);
                    range.Font.Size = sizes.SizeFor(role ?? string.Empty);
                }
                finally { Release(range); }
            }

            foreach (Word.Section section in document.Sections)
            {
                try
                {
                    foreach (Word.HeaderFooter header in section.Headers)
                    {
                        try
                        {
                            foreach (Word.Field field in header.Range.Fields)
                            {
                                try
                                {
                                    if (field.Type == Word.WdFieldType.wdFieldPage) field.Result.Font.Size = sizes.PageNumber;
                                }
                                finally { Release(field); }
                            }
                        }
                        finally { Release(header); }
                    }
                    foreach (Word.HeaderFooter footer in section.Footers)
                    {
                        try
                        {
                            foreach (Word.Field field in footer.Range.Fields)
                            {
                                try
                                {
                                    if (field.Type == Word.WdFieldType.wdFieldPage) field.Result.Font.Size = sizes.PageNumber;
                                }
                                finally { Release(field); }
                            }
                        }
                        finally { Release(footer); }
                    }
                }
                finally { Release(section); }
            }
        }

        private static int ConvertLegacyRuns(Word.Range story)
        {
            var runs = new List<LegacyRun>();
            Word.Range? character = null;
            try
            {
                var count = story.Characters.Count;
                LegacyRun? current = null;
                for (var index = 1; index <= count; index++)
                {
                    Release(character);
                    character = story.Characters[index];
                    var fontName = character.Font.Name ?? string.Empty;
                    if (!VietnameseLegacyEncodingConverter.IsLegacyFont(fontName))
                    {
                        current = null;
                        continue;
                    }
                    if (current != null && current.End == character.Start &&
                        string.Equals(current.FontName, fontName, StringComparison.OrdinalIgnoreCase))
                    {
                        current.End = character.End;
                    }
                    else
                    {
                        current = new LegacyRun(character.Start, character.End, fontName);
                        runs.Add(current);
                    }
                }
            }
            finally { Release(character); }

            for (var index = runs.Count - 1; index >= 0; index--)
            {
                Word.Range? segment = null;
                try
                {
                    segment = story.Duplicate;
                    segment.SetRange(runs[index].Start, runs[index].End);
                    var converted = VietnameseLegacyEncodingConverter.Convert(runs[index].FontName, segment.Text ?? string.Empty);
                    if (converted.UnmappedCharacters > 0)
                        throw new InvalidOperationException("Có " + converted.UnmappedCharacters.ToString(CultureInfo.InvariantCulture) +
                            " ký tự mã cũ chưa ánh xạ được; tài liệu chưa bị chuyển đổi một phần.");
                    segment.Text = converted.Text;
                    segment.Font.Name = "Times New Roman";
                }
                finally { Release(segment); }
            }
            return runs.Count;
        }

        private static void NormalizeToneInStory(Word.Range story, VietnameseTonePlacementStyle style)
        {
            foreach (Word.Paragraph paragraph in story.Paragraphs)
            {
                Word.Range? range = null;
                try
                {
                    range = paragraph.Range.Duplicate;
                    NormalizeToneInRange(range, style);
                }
                finally { Release(range); Release(paragraph); }
            }
        }

        private static void NormalizeToneInRange(Word.Range range, VietnameseTonePlacementStyle style)
        {
            var source = range.Text ?? string.Empty;
            var normalized = VietnameseTonePlacementNormalizer.Normalize(source, style);
            if (string.Equals(source, normalized, StringComparison.Ordinal)) return;
            for (var index = source.Length - 1; index >= 0; index--)
            {
                if (source[index] == normalized[index]) continue;
                Word.Range? character = null;
                try
                {
                    if (index + 1 > range.Characters.Count) continue;
                    character = range.Characters[index + 1];
                    character.Text = normalized[index].ToString();
                }
                finally { Release(character); }
            }
        }

        private static IEnumerable<Word.Range> EditableStories(Word.Document document)
        {
            var types = new[]
            {
                Word.WdStoryType.wdMainTextStory,
                Word.WdStoryType.wdFootnotesStory,
                Word.WdStoryType.wdEndnotesStory,
                Word.WdStoryType.wdPrimaryHeaderStory,
                Word.WdStoryType.wdPrimaryFooterStory,
                Word.WdStoryType.wdEvenPagesHeaderStory,
                Word.WdStoryType.wdEvenPagesFooterStory,
                Word.WdStoryType.wdFirstPageHeaderStory,
                Word.WdStoryType.wdFirstPageFooterStory,
                Word.WdStoryType.wdTextFrameStory
            };
            foreach (var type in types)
            {
                Word.Range? current = null;
                try { current = document.StoryRanges[type]; }
                catch (COMException) { }
                while (current != null)
                {
                    Word.Range? next = null;
                    try { next = current.NextStoryRange; }
                    catch (COMException) { }
                    yield return current;
                    current = next;
                }
            }
        }

        private static void ConfigureBuiltInStyles(Word.Document document, float mainSize, string fontName)
        {
            var definitions = StyleDefinition.All(mainSize);
            foreach (var definition in definitions)
            {
                Word.Style? style = null;
                try
                {
                    style = document.Styles[definition.BuiltIn];
                    style.Font.Name = fontName;
                    style.Font.Size = definition.FontSize;
                    style.Font.Bold = definition.Bold ? -1 : 0;
                    style.Font.Color = Word.WdColor.wdColorAutomatic;
                    style.Font.Kerning = 0f;
                    style.LanguageID = Word.WdLanguageID.wdVietnamese;
                    if (!definition.IsParagraph) continue;
                    style.ParagraphFormat.Alignment = definition.Alignment;
                    style.ParagraphFormat.SpaceBefore = definition.SpaceBefore;
                    style.ParagraphFormat.SpaceAfter = definition.SpaceAfter;
                    style.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    style.ParagraphFormat.LineSpacing = 12f;
                    style.ParagraphFormat.FirstLineIndent = definition.FirstLineIndent;
                    style.ParagraphFormat.LeftIndent = definition.LeftIndent;
                    style.ParagraphFormat.KeepWithNext = definition.KeepWithNext ? -1 : 0;
                    style.ParagraphFormat.KeepTogether = definition.KeepTogether ? -1 : 0;
                    style.ParagraphFormat.OutlineLevel = definition.OutlineLevel;
                    style.ParagraphFormat.TabStops.ClearAll();
                    if (definition.RightTab > 0f)
                        style.ParagraphFormat.TabStops.Add(definition.RightTab,
                            Word.WdTabAlignment.wdAlignTabRight, Word.WdTabLeader.wdTabLeaderDots);
                }
                finally { Release(style); }
            }

            Word.Style? tableGrid = null;
            try
            {
                tableGrid = document.Styles[(Word.WdBuiltinStyle)(-155)];
                foreach (Word.WdBorderType borderType in new[]
                {
                    Word.WdBorderType.wdBorderTop, Word.WdBorderType.wdBorderLeft,
                    Word.WdBorderType.wdBorderBottom, Word.WdBorderType.wdBorderRight,
                    Word.WdBorderType.wdBorderHorizontal, Word.WdBorderType.wdBorderVertical
                })
                {
                    Word.Border? border = null;
                    try
                    {
                        border = tableGrid.Table.Borders[borderType];
                        border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        border.LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                        border.Color = Word.WdColor.wdColorAutomatic;
                    }
                    finally { Release(border); }
                }
            }
            finally { Release(tableGrid); }
        }

        private sealed class LegacyRun
        {
            public LegacyRun(int start, int end, string fontName)
            {
                Start = start; End = end; FontName = fontName;
            }
            public int Start { get; }
            public int End { get; set; }
            public string FontName { get; }
        }

        private sealed class FontSizeProfile
        {
            private readonly IDictionary<string, float> _sizes;
            private FontSizeProfile(float body, float pageNumber, IDictionary<string, float> sizes)
            {
                Body = body; PageNumber = pageNumber; _sizes = sizes;
            }
            public float Body { get; }
            public float PageNumber { get; }
            public float SizeFor(string role)
            {
                float size;
                return _sizes.TryGetValue(role, out size) ? size : Body;
            }
            public static FontSizeProfile For(float mainSize)
            {
                if (mainSize == 13f) return Create(13, 13, 12, 13, 15, 12, 12, 13, 13, 13, 13, 12, 13, 13, 13, 12, 11, 14, 13);
                if (mainSize == 15f) return Create(15, 14, 13, 14, 15, 14, 14, 14, 14, 16, 15, 12, 15, 14, 14, 14, 12, 14, 14);
                return Create(14, 13, 13, 14, 15, 13, 13, 13, 14, 14, 14, 13, 14, 14, 14, 12, 11, 14, 14);
            }
            private static FontSizeProfile Create(float body, float page, float nationalTitle, float nationalMotto,
                float partyTitle, float superiorOrgan, float organ, float code, float date, float typeName,
                float subject, float officialSubject, float legalBasis, float signer, float salutation,
                float recipientLabel, float recipientList, float appendixLabel, float appendixTitle)
            {
                return new FontSizeProfile(body, page, new Dictionary<string, float>(StringComparer.Ordinal)
                {
                    ["nationalTitle"] = nationalTitle, ["nationalMotto"] = nationalMotto, ["partyTitle"] = partyTitle,
                    ["superiorOrganName"] = superiorOrgan, ["organName"] = organ, ["codeNumber"] = code,
                    ["placeAndIssuedDate"] = date, ["typeName"] = typeName, ["subject"] = subject,
                    ["officialLetterSubject"] = officialSubject, ["legalBasis"] = legalBasis,
                    ["signerAuthority"] = signer, ["signerAuthorityTitle"] = signer,
                    ["recipientSalutation"] = salutation, ["recipientSalutationList"] = salutation,
                    ["recipientSalutationInline"] = salutation, ["recipientSalutationInlineContent"] = body,
                    ["recipientLabel"] = recipientLabel, ["recipientList"] = recipientList,
                    ["appendixLabel"] = appendixLabel, ["appendixTitle"] = appendixTitle,
                    ["appendixReference"] = appendixTitle,
                    ["appendixDigitalSignatureInfo"] = 10
                });
            }
        }

        private sealed class StyleDefinition
        {
            private StyleDefinition(Word.WdBuiltinStyle builtIn, bool isParagraph, float size, bool bold,
                Word.WdParagraphAlignment alignment, float before, float after, float first, float left,
                bool keepNext, bool keepTogether, Word.WdOutlineLevel outline, float rightTab)
            {
                BuiltIn = builtIn; IsParagraph = isParagraph; FontSize = size; Bold = bold; Alignment = alignment;
                SpaceBefore = before; SpaceAfter = after; FirstLineIndent = first; LeftIndent = left;
                KeepWithNext = keepNext; KeepTogether = keepTogether; OutlineLevel = outline; RightTab = rightTab;
            }
            public Word.WdBuiltinStyle BuiltIn { get; }
            public bool IsParagraph { get; }
            public float FontSize { get; }
            public bool Bold { get; }
            public Word.WdParagraphAlignment Alignment { get; }
            public float SpaceBefore { get; }
            public float SpaceAfter { get; }
            public float FirstLineIndent { get; }
            public float LeftIndent { get; }
            public bool KeepWithNext { get; }
            public bool KeepTogether { get; }
            public Word.WdOutlineLevel OutlineLevel { get; }
            public float RightTab { get; }
            public static IReadOnlyList<StyleDefinition> All(float size)
            {
                var body = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                return new[]
                {
                    P(Word.WdBuiltinStyle.wdStyleNormal,size,false,Word.WdParagraphAlignment.wdAlignParagraphJustify,12,0,28.35f,0,false,false,body),
                    P(Word.WdBuiltinStyle.wdStyleHeading1,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,12,0,0,0,true,true,Word.WdOutlineLevel.wdOutlineLevel1),
                    P(Word.WdBuiltinStyle.wdStyleHeading2,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,12,0,0,0,true,true,Word.WdOutlineLevel.wdOutlineLevel2),
                    P(Word.WdBuiltinStyle.wdStyleHeading3,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,12,0,0,0,true,true,Word.WdOutlineLevel.wdOutlineLevel3),
                    P(Word.WdBuiltinStyle.wdStyleHeading4,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,12,0,0,0,true,true,Word.WdOutlineLevel.wdOutlineLevel4),
                    P(Word.WdBuiltinStyle.wdStyleHeading5,size,true,Word.WdParagraphAlignment.wdAlignParagraphJustify,12,0,28.35f,0,true,true,Word.WdOutlineLevel.wdOutlineLevel5),
                    P(Word.WdBuiltinStyle.wdStyleHeading6,size,true,Word.WdParagraphAlignment.wdAlignParagraphJustify,12,0,28.35f,0,true,true,Word.WdOutlineLevel.wdOutlineLevel6),
                    P(Word.WdBuiltinStyle.wdStyleTitle,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,24,0,0,0,true,true,body),
                    P(Word.WdBuiltinStyle.wdStyleSubtitle,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,0,0,0,0,true,true,body),
                    P((Word.WdBuiltinStyle)(-155),size,false,Word.WdParagraphAlignment.wdAlignParagraphLeft,0,0,0,0,false,false,body),
                    P((Word.WdBuiltinStyle)(-158),size,false,Word.WdParagraphAlignment.wdAlignParagraphLeft,0,0,0,0,false,false,body),
                    P(Word.WdBuiltinStyle.wdStyleCaption,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,6,6,0,0,false,false,body),
                    P(Word.WdBuiltinStyle.wdStyleHeader,size,false,Word.WdParagraphAlignment.wdAlignParagraphCenter,0,0,0,0,false,false,body),
                    C(Word.WdBuiltinStyle.wdStylePageNumber,size),
                    P(Word.WdBuiltinStyle.wdStyleFootnoteText,11,false,Word.WdParagraphAlignment.wdAlignParagraphJustify,0,0,28.35f,0,false,false,body),
                    P(Word.WdBuiltinStyle.wdStyleTocHeading,size,true,Word.WdParagraphAlignment.wdAlignParagraphCenter,24,12,0,0,true,true,body),
                    P(Word.WdBuiltinStyle.wdStyleTOC1,size,true,Word.WdParagraphAlignment.wdAlignParagraphLeft,6,0,0,0,false,false,body,439.4f),
                    P(Word.WdBuiltinStyle.wdStyleTOC2,size,false,Word.WdParagraphAlignment.wdAlignParagraphLeft,0,0,0,14.2f,false,false,body,439.4f),
                    P(Word.WdBuiltinStyle.wdStyleTOC3,size,false,Word.WdParagraphAlignment.wdAlignParagraphLeft,0,0,0,28.35f,false,false,body,439.4f)
                };
            }
            private static StyleDefinition P(Word.WdBuiltinStyle id, float size, bool bold,
                Word.WdParagraphAlignment alignment, float before, float after, float first, float left,
                bool keepNext, bool keep, Word.WdOutlineLevel outline, float tab = 0f) =>
                new StyleDefinition(id, true, size, bold, alignment, before, after, first, left, keepNext, keep, outline, tab);
            private static StyleDefinition C(Word.WdBuiltinStyle id, float size) =>
                new StyleDefinition(id, false, size, false, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                    0, 0, 0, 0, false, false, Word.WdOutlineLevel.wdOutlineLevelBodyText, 0);
        }

        private string Execute(
            string title,
            Action<Word.Document, LocalRulePack> operation,
            bool createBackup = false)
        {
            var document = _documentProvider == null ? _application.ActiveDocument : _documentProvider();
            if (document == null)
                throw new InvalidOperationException("Hãy mở một tài liệu Word trước khi sử dụng Chuẩn hóa.");
            var capability = _capabilityProvider.Evaluate(document);
            if (!capability.CanReadDocument) throw new InvalidOperationException(capability.Reason);
            if (capability.IsReadOnly || capability.IsProtected || capability.TrackChangesEnabled)
                throw new InvalidOperationException("Tài liệu phải cho phép chỉnh sửa, không bảo vệ và tắt Track Changes.");
            var rules = _accessManager.GetRulePack(LocalAccessManager.DocumentToolsFeature);
            var backup = createBackup ? CreateBackup(document) : string.Empty;
            var previousScreenUpdating = _application.ScreenUpdating;
            var undoStarted = false;
            try
            {
                _application.ScreenUpdating = false;
                if (WordMajorVersion() >= 15)
                {
                    _application.UndoRecord.StartCustomRecord("Chuẩn hóa: " + title);
                    undoStarted = true;
                }
                operation(document, rules);
                return backup;
            }
            catch (Exception exception)
            {
                if (undoStarted)
                {
                    try { _application.UndoRecord.EndCustomRecord(); }
                    catch (COMException) { }
                    undoStarted = false;
                }
                try
                {
                    object times = 1;
                    document.Undo(ref times);
                }
                catch { /* Word may have nothing left to undo. */ }
                var recoveryMessage = string.IsNullOrWhiteSpace(backup)
                    ? string.Empty
                    : "\nBản sao khôi phục: " + backup;
                throw new InvalidOperationException(exception.Message + recoveryMessage, exception);
            }
            finally
            {
                if (undoStarted)
                {
                    try { _application.UndoRecord.EndCustomRecord(); }
                    catch (COMException) { }
                }
                _application.ScreenUpdating = previousScreenUpdating;
            }
        }

        private static string CreateBackup(Word.Document document)
        {
            return WordRecoveryCopyManager.Create(document.Application, document, "tone");
        }

        private Word.Range MutableSelectionOrParagraph()
        {
            var selection = _application.Selection;
            if (selection.Start != selection.End) return selection.Range.Duplicate;
            return selection.Paragraphs[1].Range.Duplicate;
        }

        private static void ReplaceAll(Word.Range range, string findText, string replacement)
        {
            var find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Execute(FindText: findText, ReplaceWith: replacement,
                Replace: Word.WdReplace.wdReplaceAll, Wrap: Word.WdFindWrap.wdFindStop);
        }

        private int WordMajorVersion()
        {
            var value = _application.Version ?? string.Empty;
            var dot = value.IndexOf('.');
            int major;
            return int.TryParse(dot < 0 ? value : value.Substring(0, dot), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out major) ? major : 0;
        }

        private sealed class TableRowStats
        {
            public TableRowStats(int firstCellStart)
            {
                FirstCellStart = firstCellStart;
            }

            public int FirstCellStart { get; set; }
            public int TotalCells { get; set; }
            public int NonEmptyCells { get; set; }
            public int BoldCells { get; set; }
            public int HeadingCells { get; set; }
            public bool AllCellsAlreadyHeading => TotalCells > 0 && HeadingCells == TotalCells;
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }

    /// <summary>
    /// Deletes only removable empty paragraph marks immediately before Word's required
    /// final paragraph. The previous implementation repeatedly tried to delete that
    /// required sentinel; Word kept the paragraph count unchanged and the loop never
    /// terminated. This implementation finds a bounded contiguous tail and performs one
    /// deletion, so repagination is triggered at most once.
    /// </summary>
    internal static class WordTrailingBlankPageCleaner
    {
        private const int MaximumParagraphsPerRun = 4096;

        public static int Remove(Word.Document document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            Word.Range? documentRange = null;
            Word.Range? deleteRange = null;
            try
            {
                paragraphs = document.Paragraphs;
                var count = paragraphs.Count;
                if (count <= 1) return 0;

                // Count is Word's immutable final paragraph. Start immediately before it.
                var deleteStart = -1;
                var inspected = 0;
                for (var index = count - 1; index >= 1 && inspected < MaximumParagraphsPerRun;
                     index--, inspected++)
                {
                    Release(paragraphRange);
                    paragraphRange = null;
                    Release(paragraph);
                    paragraph = null;

                    paragraph = paragraphs[index];
                    paragraphRange = paragraph.Range.Duplicate;
                    var raw = paragraphRange.Text ?? string.Empty;
                    // A table-cell marker (\a), manual page break (\f) or manual line
                    // break (\v) is structural content and must never be deleted here.
                    if (raw.IndexOf('\a') >= 0 || raw.IndexOf('\f') >= 0 || raw.IndexOf('\v') >= 0 ||
                        raw.Trim('\r', ' ', '\t').Length != 0)
                        break;
                    deleteStart = paragraphRange.Start;
                }

                if (deleteStart < 0) return 0;
                documentRange = document.Content.Duplicate;
                var deleteEnd = documentRange.End - 1; // Preserve Word's final sentinel.
                if (deleteEnd <= deleteStart) return 0;
                var before = documentRange.End;
                deleteRange = document.Range(deleteStart, deleteEnd);
                deleteRange.Delete();

                Release(documentRange);
                documentRange = document.Content.Duplicate;
                return Math.Max(0, before - documentRange.End);
            }
            finally
            {
                Release(deleteRange);
                Release(documentRange);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
            }
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
