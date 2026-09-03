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
            var text = (range.Text ?? string.Empty).Trim('\r', '\a', '\v', ' ', '\t');
            if (text.Length == 0) return 0d;
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

            var available = ReadAvailableWidth(range);
            return available.HasValue ? Math.Min(measured, available.Value) : measured;
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
                Word.Range? cellRange = null;
                Word.Range? cellStart = null;
                try
                {
                    cells = range.Cells;
                    if (cells.Count == 0) return null;
                    cell = cells[1];
                    cellRange = cell.Range.Duplicate;
                    cellStart = cellRange.Duplicate;
                    cellStart.SetRange(cellRange.Start, Math.Min(cellRange.Start + 1, cellRange.End));
                    // A one-character cell range returns the start of the cell's text
                    // area, after the left cell padding. Treating it as the outer cell
                    // edge adds the padding twice and shifts every centred Line Shape
                    // to the right by roughly half the combined cell padding.
                    var left = SafeInformation(cellStart,
                        Word.WdInformation.wdHorizontalPositionRelativeToPage);
                    if (!left.HasValue) return null;
                    var leftPadding = ValidIndent(cell.LeftPadding);
                    var rightPadding = ValidIndent(cell.RightPadding);
                    var availableWidth = Math.Max(12d, cell.Width - leftPadding - rightPadding);
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
                finally { Release(cellStart); Release(cellRange); Release(cell); Release(cells); }
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
