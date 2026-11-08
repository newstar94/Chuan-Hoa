using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// Joins lines that were broken only by the line width of the source (text
    /// copied from a PDF) inside a selection. Each break, with the blanks around
    /// it, is replaced on its own: the words of the lines keep their character
    /// formatting, and a merged paragraph keeps the format of its first line.
    /// Table cells, and manual line breaks in paragraphs holding fields, are
    /// left as they are.
    /// </summary>
    internal static class WordBrokenLineJoiner
    {
        private sealed class PlannedJoin
        {
            public PlannedJoin(int start, string expectedText, string replacement, int? mergedParagraph)
            {
                Start = start;
                ExpectedText = expectedText;
                Replacement = replacement;
                MergedParagraph = mergedParagraph;
            }

            public int Start { get; }
            public string ExpectedText { get; }
            public string Replacement { get; }

            /// <summary>Index of the paragraph whose mark is removed; null for a manual line break.</summary>
            public int? MergedParagraph { get; }
        }

        /// <summary>Returns the number of breaks joined.</summary>
        internal static int Join(Word.Range selection)
        {
            if (selection == null) throw new ArgumentNullException(nameof(selection));
            var selectionStart = selection.Start;
            var selectionEnd = selection.End;
            var ranges = new List<Word.Range>();
            var savedFormats = new Dictionary<int, Word.ParagraphFormat>();
            Word.Paragraphs? paragraphs = null;
            try
            {
                paragraphs = selection.Paragraphs;
                foreach (Word.Paragraph paragraph in paragraphs)
                {
                    try { ranges.Add(paragraph.Range.Duplicate); }
                    finally { Release(paragraph); }
                }

                var count = ranges.Count;
                var texts = new string[count];
                var starts = new int[count];
                var ends = new int[count];
                var skipped = new bool[count];
                var hasFields = new bool[count];
                for (var index = 0; index < count; index++)
                {
                    var range = ranges[index];
                    texts[index] = range.Text ?? string.Empty;
                    starts[index] = range.Start;
                    ends[index] = range.End;
                    bool inTable = range.get_Information(Word.WdInformation.wdWithInTable);
                    skipped[index] = inTable || !texts[index].EndsWith("\r", StringComparison.Ordinal);
                    hasFields[index] = range.Fields.Count > 0;
                }

                var plans = new List<PlannedJoin>();
                for (var index = 0; index < count; index++)
                {
                    if (skipped[index]) continue;
                    var lines = Body(texts[index]).Split('\v');

                    // Manual line breaks (Shift+Enter) inside the paragraph. Their
                    // offsets map to positions only without fields before them.
                    if (!hasFields[index])
                    {
                        var offset = 0;
                        for (var line = 0; line < lines.Length - 1; line++)
                        {
                            offset += lines[line].Length;
                            var position = starts[index] + offset;
                            if (position >= selectionStart && position + 1 < selectionEnd)
                                Plan(plans, lines[line], lines[line + 1], position, '\v', null);
                            offset++;
                        }
                    }

                    // The paragraph mark, when the next paragraph is selected too.
                    var next = index + 1;
                    if (next >= count || skipped[next] || starts[next] != ends[index]) continue;
                    var markPosition = ends[index] - 1;
                    if (markPosition < selectionStart || starts[next] >= selectionEnd) continue;
                    var nextFirstLine = Body(texts[next]).Split('\v')[0];
                    Plan(plans, lines[lines.Length - 1], nextFirstLine, markPosition, '\r', index);
                }
                if (plans.Count == 0) return 0;

                // A run of merged paragraphs keeps the paragraph format of its first
                // line, whichever mark Word keeps when the others are removed.
                var merged = new HashSet<int>(plans.Where(plan => plan.MergedParagraph.HasValue)
                    .Select(plan => plan.MergedParagraph!.Value));
                foreach (var index in merged.Where(index => !merged.Contains(index - 1)))
                    savedFormats[index] = ranges[index].ParagraphFormat.Duplicate;

                var joined = 0;
                var appliedRuns = new HashSet<int>();
                foreach (var plan in plans.OrderByDescending(item => item.Start))
                {
                    if (!Apply(selection, plan)) continue;
                    joined++;
                    if (plan.MergedParagraph.HasValue)
                    {
                        var first = plan.MergedParagraph.Value;
                        while (!savedFormats.ContainsKey(first)) first--;
                        appliedRuns.Add(first);
                    }
                }
                foreach (var first in appliedRuns)
                    ranges[first].ParagraphFormat = savedFormats[first];
                return joined;
            }
            finally
            {
                foreach (var format in savedFormats.Values) Release(format);
                foreach (var range in ranges) Release(range);
                Release(paragraphs);
            }
        }

        /// <summary>
        /// Paragraph roles whose text is running prose. Every other role is part
        /// of the document frame (national title, organ, number, place and date,
        /// type name, subject, recipients, signature, chapter/section headings,
        /// appendix labels), where line breaks are deliberate.
        /// </summary>
        private static readonly HashSet<string> ProseRoles =
            new HashSet<string>(StringComparer.Ordinal) { string.Empty, "legalBasis" };

        /// <summary>
        /// Below this length the line before a paragraph mark did not reach the
        /// margin of an A4 page in 13–14 pt, so in a whole-document run the mark is
        /// removed only when the next paragraph continues in lowercase.
        /// </summary>
        private const int MinimumWrappedParagraphLineLength = 60;

        /// <summary>
        /// Whole-document pass used by 1-Click normalization: joins wrapped lines
        /// (manual line breaks, and paragraph marks between two prose paragraphs)
        /// in the main story. Positions come from <paramref name="snapshot"/>,
        /// which must have been read from the document as it is now; every span is
        /// verified against the document before it is replaced.
        /// </summary>
        internal static int JoinWrappedBodyLines(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (roles == null) throw new ArgumentNullException(nameof(roles));

            var paragraphs = snapshot.Paragraphs
                .Where(item => string.Equals(item.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                .OrderBy(item => item.Index)
                .ToList();
            var plans = new List<PlannedJoin>();
            for (var index = 0; index < paragraphs.Count; index++)
            {
                var paragraph = paragraphs[index];
                if (!IsProse(paragraph, snapshot, roles)) continue;
                var lines = paragraph.Text.Split('\v');
                var offset = 0;
                for (var line = 0; line < lines.Length - 1; line++)
                {
                    offset += lines[line].Length;
                    Plan(plans, lines[line], lines[line + 1], paragraph.AbsoluteStart + offset, '\v', null);
                    offset++;
                }

                if (index + 1 >= paragraphs.Count) continue;
                var next = paragraphs[index + 1];
                if (next.Index != paragraph.Index + 1 || next.AbsoluteStart != paragraph.AbsoluteEnd ||
                    !IsProse(next, snapshot, roles))
                    continue;
                var lastLine = lines[lines.Length - 1];
                var nextFirstLine = next.Text.Split('\v')[0];
                var nextStart = nextFirstLine.TrimStart(' ', '\t', ' ');
                var continuesInLowercase = nextStart.Length > 0 && char.IsLower(nextStart[0]);
                if (lastLine.Trim(' ', '\t', ' ').Length < MinimumWrappedParagraphLineLength &&
                    !continuesInLowercase)
                    continue;
                Plan(plans, lastLine, nextFirstLine, paragraph.AbsoluteEnd - 1, '\r', paragraph.Index);
            }
            if (plans.Count == 0) return 0;

            Word.Range? content = null;
            try
            {
                content = document.Content;
                var joined = 0;
                foreach (var plan in plans.OrderByDescending(item => item.Start))
                    if (Apply(content, plan)) joined++;
                return joined;
            }
            finally { Release(content); }
        }

        private static bool IsProse(LocalParagraphSnapshot paragraph, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            string? role;
            roles.TryGetValue(paragraph.Index, out role);
            return ProseRoles.Contains(role ?? string.Empty) && !paragraph.IsInTable &&
                !string.IsNullOrWhiteSpace(paragraph.Text) &&
                paragraph.HasField != true && paragraph.HasHyperlink != true &&
                paragraph.HasContentControl != true && paragraph.HasMathObject != true &&
                !snapshot.IntersectsProtectedSpan(paragraph);
        }

        private static void Plan(List<PlannedJoin> plans, string previousLine, string nextLine,
            int breakPosition, char breakCharacter, int? mergedParagraph)
        {
            var decision = BrokenLineJoiner.Decide(previousLine, nextLine);
            if (decision == LineBreakDecision.Keep) return;
            var trailing = BrokenLineJoiner.TrailingBlankCount(previousLine);
            var leading = BrokenLineJoiner.LeadingBlankCount(nextLine);
            var expected = previousLine.Substring(previousLine.Length - trailing) + breakCharacter +
                nextLine.Substring(0, leading);
            plans.Add(new PlannedJoin(breakPosition - trailing, expected,
                decision == LineBreakDecision.JoinWithSpace ? " " : string.Empty, mergedParagraph));
        }

        private static bool Apply(Word.Range story, PlannedJoin plan)
        {
            Word.Range? span = null;
            try
            {
                span = story.Duplicate;
                span.SetRange(plan.Start, plan.Start + plan.ExpectedText.Length);
                // The span holds only blanks and the break: anything else means the
                // text offsets did not map to document positions.
                if (!string.Equals(span.Text ?? string.Empty, plan.ExpectedText, StringComparison.Ordinal) ||
                    span.Fields.Count > 0 || span.InlineShapes.Count > 0 || span.ContentControls.Count > 0)
                    return false;
                span.Text = plan.Replacement;
                return true;
            }
            finally { Release(span); }
        }

        private static string Body(string paragraphText) =>
            paragraphText.EndsWith("\r", StringComparison.Ordinal)
                ? paragraphText.Substring(0, paragraphText.Length - 1)
                : paragraphText;

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
