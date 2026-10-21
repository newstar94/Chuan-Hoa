using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    internal static class WordTextMeasurement
    {
        public static double MeasureParagraphWidth(Word.Range range, string? fontName,
            double? fontSizePoints, bool? bold, bool? italic)
        {
            var raw = (range.Text ?? string.Empty).TrimEnd('\r', '\a');
            var text = raw.Trim('\r', '\a', '\v', ' ', '\t');
            if (text.Length == 0) return 0d;

            // Measure the line directly above the separator: the last Shift+Enter
            // segment of a two-line organ name, not both lines added up.
            var lastBreak = text.LastIndexOfAny(new[] { '\v', '\n' });
            if (lastBreak >= 0) text = text.Substring(lastBreak + 1).Trim(' ', '\t');
            if (text.Length == 0) return 0d;
            var typographic = MeasureTypographically(range, text, fontName, fontSizePoints, bold, italic);

            // Word's own layout is authoritative (kerning, rounding) only while it
            // agrees with the font metrics. Information(wdHorizontalPosition...) can
            // report the start of a centred line at the line's left boundary instead
            // of its first glyph; that reading is (column + text) / 2 wide and drew
            // separators almost twice as long as the motto.
            var native = MeasureLastSegmentInWord(range, raw);
            if (native.HasValue && Math.Abs(native.Value - typographic) <= typographic * .1d)
                return native.Value;
            return typographic;
        }

        /// <summary>
        /// Width of <paramref name="text"/> from the font metrics (within 1% of
        /// Word for Times New Roman), adjusted for the range's character scaling
        /// and spacing and capped at the width available to the paragraph.
        /// </summary>
        private static double MeasureTypographically(Word.Range range, string text, string? fontName,
            double? fontSizePoints, bool? bold, bool? italic)
        {
            var family = string.IsNullOrWhiteSpace(fontName) ? "Times New Roman" : fontName;
            var size = (float)Math.Max(6d, fontSizePoints.GetValueOrDefault(13d));
            var style = FontStyle.Regular;
            if (bold.GetValueOrDefault()) style |= FontStyle.Bold;
            if (italic.GetValueOrDefault()) style |= FontStyle.Italic;
            double measured;
            using (var bitmap = new Bitmap(1, 1))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var font = new Font(family, size, style, GraphicsUnit.Point))
            using (var format = (StringFormat)StringFormat.GenericTypographic.Clone())
            {
                graphics.PageUnit = GraphicsUnit.Point;
                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                measured = graphics.MeasureString(text, font, int.MaxValue, format).Width;
            }

            Word.Font? rangeFont = null;
            try
            {
                rangeFont = range.Font;
                var scaling = rangeFont.Scaling;
                if (scaling > 0 && scaling < 9999999) measured *= scaling / 100d;
                var spacing = rangeFont.Spacing;
                if (spacing > -999999f && spacing < 999999f) measured += spacing * text.Length;
            }
            catch (COMException)
            {
            }
            finally { Release(rangeFont); }

            var available = ReadAvailableWidth(range);
            return available.HasValue ? Math.Min(measured, available.Value) : measured;
        }

        /// <summary>
        /// Width of the last text segment measured with Range.Information on the
        /// collapsed positions before its first and after its last character. The
        /// difference is independent of a stale absolute X after table resizing.
        /// Returns null when the segment wraps or Word cannot report positions.
        /// </summary>
        private static double? MeasureLastSegmentInWord(Word.Range range, string printable)
        {
            if (printable.Length == 0 || printable.Length > 250) return null;
            var segmentStart = printable.LastIndexOfAny(new[] { '\v', '\n' }) + 1;
            while (segmentStart < printable.Length && IsTrailingOrLeading(printable[segmentStart]))
                segmentStart++;
            var segmentEnd = printable.Length;
            while (segmentEnd > segmentStart && IsTrailingOrLeading(printable[segmentEnd - 1]))
                segmentEnd--;
            if (segmentEnd <= segmentStart) return null;

            Word.Range? first = null;
            Word.Range? last = null;
            try
            {
                first = range.Duplicate;
                first.SetRange(range.Start + segmentStart, range.Start + segmentStart);
                last = range.Duplicate;
                last.SetRange(range.Start + segmentEnd, range.Start + segmentEnd);
                var firstLine = SafeInformation(first, Word.WdInformation.wdFirstCharacterLineNumber);
                var lastLine = SafeInformation(last, Word.WdInformation.wdFirstCharacterLineNumber);
                if (!firstLine.HasValue || !lastLine.HasValue || firstLine.Value != lastLine.Value)
                    return null;
                var left = SafeInformation(first, Word.WdInformation.wdHorizontalPositionRelativeToPage);
                var right = SafeInformation(last, Word.WdInformation.wdHorizontalPositionRelativeToPage);
                if (!left.HasValue || !right.HasValue) return null;
                var width = right.Value - left.Value;
                return width > 1d && width < 2000d ? width : (double?)null;
            }
            catch (COMException)
            {
                return null;
            }
            finally { Release(last); Release(first); }
        }

        private static bool IsTrailingOrLeading(char value) =>
            value == '\r' || value == '\a' || value == '\v' || value == ' ' || value == '\t';

        public static double? ReadAvailableWidth(Word.Range range)
        {
            if (range.get_Information(Word.WdInformation.wdWithInTable))
            {
                Word.Cells? cells = null;
                Word.Cell? cell = null;
                try
                {
                    cells = range.Cells;
                    if (cells.Count == 0) return null;
                    cell = cells[1];
                    return Math.Max(12d, cell.Width -
                        ValidIndent(cell.LeftPadding) - ValidIndent(cell.RightPadding));
                }
                catch (COMException)
                {
                    return null;
                }
                finally { Release(cell); Release(cells); }
            }

            Word.Sections? sections = null;
            Word.Section? section = null;
            Word.PageSetup? setup = null;
            try
            {
                sections = range.Sections;
                if (sections.Count == 0) return null;
                section = sections[1];
                setup = section.PageSetup;
                return Math.Max(12d, setup.PageWidth - setup.LeftMargin - setup.RightMargin);
            }
            catch (COMException)
            {
                return null;
            }
            finally { Release(setup); Release(section); Release(sections); }
        }

        /// <summary>
        /// Length of a paragraph's bottom border as Word draws it: from the
        /// leftmost indent (a hanging first line included) to the right indent of
        /// the column or cell, whatever the width of the text above it.
        /// </summary>
        public static double? MeasureBottomBorderLength(Word.Range range)
        {
            var available = ReadAvailableWidth(range);
            if (!available.HasValue) return null;
            Word.ParagraphFormat? paragraphFormat = null;
            try
            {
                paragraphFormat = range.ParagraphFormat;
                var leftIndent = ValidIndent(paragraphFormat.LeftIndent);
                var firstLineIndent = ValidIndent(paragraphFormat.FirstLineIndent);
                var rightIndent = ValidIndent(paragraphFormat.RightIndent);
                var length = available.Value - Math.Min(leftIndent, leftIndent + firstLineIndent) - rightIndent;
                return length > 0d ? length : (double?)null;
            }
            catch (COMException)
            {
                return null;
            }
            finally { Release(paragraphFormat); }
        }

        public static double? ReadHorizontalCenter(Word.Range range)
        {
            Word.ParagraphFormat? paragraphFormat = null;
            var leftIndent = 0d;
            var rightIndent = 0d;
            var firstLineIndent = 0d;
            try
            {
                paragraphFormat = range.ParagraphFormat;
                leftIndent = ValidIndent(paragraphFormat.LeftIndent);
                rightIndent = ValidIndent(paragraphFormat.RightIndent);
                firstLineIndent = ValidIndent(paragraphFormat.FirstLineIndent);
            }
            catch (COMException)
            {
            }
            finally { Release(paragraphFormat); }

            if (range.get_Information(Word.WdInformation.wdWithInTable))
            {
                Word.Cells? cells = null;
                Word.Cell? cell = null;
                try
                {
                    cells = range.Cells;
                    if (cells.Count == 0) return null;
                    cell = cells[1];
                    var left = ReadCellContentLeft(cell);
                    if (!left.HasValue) return null;
                    var availableWidth = Math.Max(12d,
                        cell.Width - ValidIndent(cell.LeftPadding) - ValidIndent(cell.RightPadding));
                    var contentLeft = left.Value + leftIndent + firstLineIndent;
                    var contentRight = left.Value + availableWidth - rightIndent;
                    return contentRight > contentLeft
                        ? contentLeft + (contentRight - contentLeft) / 2d
                        : left.Value + availableWidth / 2d;
                }
                catch (COMException)
                {
                    return null;
                }
                finally
                {
                    Release(cell);
                    Release(cells);
                }
            }

            Word.Sections? sections = null;
            Word.Section? section = null;
            Word.PageSetup? setup = null;
            try
            {
                sections = range.Sections;
                if (sections.Count == 0) return null;
                section = sections[1];
                setup = section.PageSetup;
                var contentLeft = setup.LeftMargin + leftIndent + firstLineIndent;
                var contentRight = setup.PageWidth - setup.RightMargin - rightIndent;
                return contentRight > contentLeft
                    ? contentLeft + (contentRight - contentLeft) / 2d
                    : setup.LeftMargin + (setup.PageWidth - setup.LeftMargin - setup.RightMargin) / 2d;
            }
            catch (COMException)
            {
                return null;
            }
            finally { Release(setup); Release(section); Release(sections); }
        }

        /// <summary>
        /// Page X of the text area (after the left cell padding) of the table cell
        /// holding <paramref name="range"/>; null outside a table.
        /// </summary>
        public static double? ReadCellContentLeft(Word.Range range)
        {
            if (!range.get_Information(Word.WdInformation.wdWithInTable)) return null;
            Word.Cells? cells = null;
            Word.Cell? cell = null;
            try
            {
                cells = range.Cells;
                if (cells.Count == 0) return null;
                cell = cells[1];
                return ReadCellContentLeft(cell);
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Release(cell);
                Release(cells);
            }
        }

        private static double? ReadCellContentLeft(Word.Cell cell)
        {
            Word.Range? cellRange = null;
            Word.Range? cellStart = null;
            Word.Tables? tables = null;
            Word.Table? table = null;
            Word.Rows? rows = null;
            Word.Row? row = null;
            Word.Cells? rowCells = null;
            Word.Sections? cellSections = null;
            Word.Section? cellSection = null;
            Word.PageSetup? cellSetup = null;
            try
            {
                cellRange = cell.Range.Duplicate;
                tables = cellRange.Tables;
                table = tables.Count > 0 ? tables[1] : null;
                cellSections = cellRange.Sections;
                cellSection = cellSections.Count > 0 ? cellSections[1] : null;
                cellSetup = cellSection?.PageSetup;
                if (table != null)
                {
                    rows = table.Rows;
                    // Rows[i] throws for a table with vertically merged cells.
                    try { row = rows[cell.RowIndex]; }
                    catch (COMException) { row = null; }
                }

                // Information(wdHorizontalPositionRelativeToPage) can keep the
                // pre-resize X coordinate until Word saves or reopens the file.
                // For the left-aligned header layout table, derive the cell text
                // origin from current table geometry instead. This keeps a line
                // centred after 1-Click changes the table indent/column widths.
                if (row != null && cellSetup != null &&
                    row.Alignment == Word.WdRowAlignment.wdAlignRowLeft)
                {
                    rowCells = row.Cells;
                    var precedingWidth = 0d;
                    var firstLeftPadding = 0d;
                    for (var index = 1; index <= rowCells.Count; index++)
                    {
                        Word.Cell? candidate = null;
                        Word.Range? candidateRange = null;
                        try
                        {
                            candidate = rowCells[index];
                            if (index == 1) firstLeftPadding = ValidIndent(candidate.LeftPadding);
                            candidateRange = candidate.Range;
                            if (candidateRange.Start >= cellRange.Start) break;
                            precedingWidth += candidate.Width;
                        }
                        finally { Release(candidateRange); Release(candidate); }
                    }
                    // Read this row, not table.Rows: the collection reports wdUndefined
                    // (9999999) once any row differs, which read as 0 and moved every
                    // header line right by the whole negative table indent.
                    var rowIndent = ValidIndent(row.LeftIndent);
                    // Layouts before Word 2013 (compatibility mode below 15) measure
                    // the indent to the first cell's text, so its padding hangs out.
                    var outdent = UsesLegacyTableIndent(cellRange) ? firstLeftPadding : 0d;
                    return cellSetup.LeftMargin + rowIndent - outdent + precedingWidth +
                        ValidIndent(cell.LeftPadding);
                }

                cellStart = cellRange.Duplicate;
                cellStart.SetRange(cellRange.Start, Math.Min(cellRange.Start + 1, cellRange.End));
                // A one-character cell range returns the start of the cell's text
                // area, after the left cell padding. Treating it as the outer cell
                // edge adds the padding twice and shifts every centred Line Shape
                // to the right by roughly half the combined cell padding.
                return SafeInformation(cellStart, Word.WdInformation.wdHorizontalPositionRelativeToPage);
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Release(cellSetup);
                Release(cellSection);
                Release(cellSections);
                Release(rowCells);
                Release(row);
                Release(rows);
                Release(table);
                Release(tables);
                Release(cellStart);
                Release(cellRange);
            }
        }

        private static bool UsesLegacyTableIndent(Word.Range range)
        {
            try
            {
                // The Document RCW is shared with the caller; it is not released here.
                return range.Document.CompatibilityMode < 15;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static double ValidIndent(float value) =>
            value > -999999f && value < 999999f ? value : 0d;

        public static double? ReadLastTextLineTop(Word.Range range)
        {
            Word.Range? last = null;
            try
            {
                var text = range.Text ?? string.Empty;
                var trailing = 0;
                for (var index = text.Length - 1; index >= 0; index--)
                {
                    var value = text[index];
                    if (value != '\r' && value != '\a' && value != '\v' &&
                        value != ' ' && value != '\t')
                        break;
                    trailing++;
                }
                if (text.Length - trailing <= 0)
                    return SafeInformation(range, Word.WdInformation.wdVerticalPositionRelativeToPage);
                var end = Math.Max(range.Start + 1, range.End - trailing);
                last = range.Duplicate;
                last.SetRange(end - 1, end);
                return SafeInformation(last, Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            catch (COMException)
            {
                return SafeInformation(range, Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            finally { Release(last); }
        }

        public static double? ReadFirstTextLineTop(Word.Range range)
        {
            Word.Range? first = null;
            try
            {
                var text = range.Text ?? string.Empty;
                var leading = 0;
                while (leading < text.Length && IsTrailingOrLeading(text[leading])) leading++;
                if (leading >= text.Length)
                    return SafeInformation(range, Word.WdInformation.wdVerticalPositionRelativeToPage);
                var start = range.Start + leading;
                first = range.Duplicate;
                first.SetRange(start, Math.Min(start + 1, range.End));
                return SafeInformation(first, Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            catch (COMException)
            {
                return SafeInformation(range, Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            finally { Release(first); }
        }

        public static double? SafeInformation(Word.Range range, Word.WdInformation information)
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

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
