using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// Keeps an appendix heading block with its first data table without changing the
    /// portrait/landscape choice made by the user. A redundant next-page section break
    /// may be changed to continuous only when both adjacent sections have identical
    /// page geometry. Different orientations are deliberately left untouched.
    /// </summary>
    internal static class WordAppendixPaginationNormalizer
    {
        public static int Normalize(Word.Document document, IReadOnlyDictionary<int, string> roles)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (roles == null || roles.Count == 0) return 0;

            var changes = 0;
            var ordered = roles
                .Where(item => IsAppendixHeadingRole(item.Value))
                .OrderBy(item => item.Key)
                .ToArray();
            foreach (var label in ordered.Where(item => item.Value == "appendixLabel"))
            {
                var nextLabel = ordered.FirstOrDefault(item =>
                    item.Key > label.Key && item.Value == "appendixLabel");
                var endExclusive = nextLabel.Equals(default(KeyValuePair<int, string>))
                    ? int.MaxValue
                    : nextLabel.Key;
                var block = ordered.Where(item => item.Key >= label.Key && item.Key < endExclusive)
                    .ToArray();
                if (block.Length == 0) continue;

                Word.Paragraph? lastHeading = null;
                Word.Range? lastHeadingRange = null;
                Word.Table? table = null;
                Word.Range? tableRange = null;
                Word.Range? gap = null;
                Word.Section? headingSection = null;
                Word.Section? tableSection = null;
                try
                {
                    foreach (var item in block)
                    {
                        Word.Paragraph? heading = null;
                        try
                        {
                            if (item.Key < 1 || item.Key > document.Paragraphs.Count) continue;
                            heading = document.Paragraphs[item.Key];
                            // The complete label/title/reference block moves together. If
                            // there is genuinely not enough room, Word moves the block with
                            // the table instead of leaving an orphan heading page.
                            heading.Format.KeepWithNext = -1;
                        }
                        finally { Release(heading); }
                    }

                    var lastIndex = block.Max(item => item.Key);
                    if (lastIndex < 1 || lastIndex > document.Paragraphs.Count) continue;
                    lastHeading = document.Paragraphs[lastIndex];
                    lastHeadingRange = lastHeading.Range.Duplicate;
                    table = FindFirstFollowingTable(document, lastHeadingRange.End);
                    if (table == null) continue;
                    tableRange = table.Range.Duplicate;
                    gap = document.Range(lastHeadingRange.End, tableRange.Start);
                    if (!ContainsOnlyLayoutBoundary(gap.Text ?? string.Empty)) continue;

                    headingSection = lastHeadingRange.Sections[1];
                    tableSection = tableRange.Sections[1];
                    if (tableSection.Index != headingSection.Index + 1 ||
                        !HasSamePageGeometry(headingSection, tableSection))
                        continue;

                    if (tableSection.PageSetup.SectionStart != Word.WdSectionStart.wdSectionContinuous)
                    {
                        tableSection.PageSetup.SectionStart = Word.WdSectionStart.wdSectionContinuous;
                        changes++;
                    }
                }
                catch (COMException)
                {
                    // A protected or structurally complex appendix remains unchanged.
                }
                finally
                {
                    Release(tableSection);
                    Release(headingSection);
                    Release(gap);
                    Release(tableRange);
                    Release(table);
                    Release(lastHeadingRange);
                    Release(lastHeading);
                }
            }
            return changes;
        }

        private static Word.Table? FindFirstFollowingTable(Word.Document document, int after)
        {
            Word.Tables? tables = null;
            try
            {
                tables = document.Tables;
                for (var index = 1; index <= tables.Count; index++)
                {
                    Word.Table? table = null;
                    Word.Range? range = null;
                    var selected = false;
                    try
                    {
                        table = tables[index];
                        range = table.Range;
                        if (range.Start >= after)
                        {
                            selected = true;
                            return table;
                        }
                    }
                    finally
                    {
                        Release(range);
                        if (!selected) Release(table);
                    }
                }
                return null;
            }
            finally { Release(tables); }
        }

        private static bool HasSamePageGeometry(Word.Section left, Word.Section right)
        {
            Word.PageSetup? a = null;
            Word.PageSetup? b = null;
            try
            {
                a = left.PageSetup;
                b = right.PageSetup;
                return a.Orientation == b.Orientation &&
                    Near(a.PageWidth, b.PageWidth) && Near(a.PageHeight, b.PageHeight) &&
                    Near(a.TopMargin, b.TopMargin) && Near(a.BottomMargin, b.BottomMargin) &&
                    Near(a.LeftMargin, b.LeftMargin) && Near(a.RightMargin, b.RightMargin) &&
                    Near(a.HeaderDistance, b.HeaderDistance) && Near(a.FooterDistance, b.FooterDistance);
            }
            finally { Release(b); Release(a); }
        }

        private static bool ContainsOnlyLayoutBoundary(string value) =>
            value.All(character => character == '\r' || character == '\f' ||
                character == '\v' || character == ' ' || character == '\t');

        private static bool IsAppendixHeadingRole(string role) =>
            role == "appendixLabel" || role == "appendixTitle" ||
            role == "appendixReference" || role == "appendixDigitalSignatureInfo";

        private static bool Near(float left, float right) => Math.Abs(left - right) <= 1f;

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
