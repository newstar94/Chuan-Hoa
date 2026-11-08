using System;
using System.Runtime.InteropServices;
using ChuanHoa.Client.Core.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// Writes a text transformation back to Word by changing only the characters
    /// that differ. Assigning <c>Range.Text</c> for a whole paragraph gives every
    /// character the formatting of the first one (a bold or italic word loses
    /// it), flattens fields and hyperlinks into plain text and replaces the
    /// paragraph mark; small span edits leave every untouched character as it was.
    /// </summary>
    internal static class WordMinimalTextEditor
    {
        /// <summary>
        /// Transforms the text of <paramref name="range"/> without its trailing
        /// paragraph/cell markers and applies the difference as span edits.
        /// Returns the number of spans changed: 0 when nothing differs, or when
        /// the range holds fields, content controls or equations, or text whose
        /// characters do not map one-to-one to document positions (hidden text).
        /// </summary>
        internal static int Apply(Word.Range range, Func<string, string> transform)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));
            if (transform == null) throw new ArgumentNullException(nameof(transform));

            var text = range.Text ?? string.Empty;
            var coreLength = text.Length;
            while (coreLength > 0 && (text[coreLength - 1] == '\r' || text[coreLength - 1] == '\a'))
                coreLength--;
            var core = text.Substring(0, coreLength);
            if (string.IsNullOrWhiteSpace(core)) return 0;

            var target = transform(core) ?? string.Empty;
            if (string.Equals(core, target, StringComparison.Ordinal)) return 0;
            if (range.Fields.Count > 0 || range.ContentControls.Count > 0 || range.OMaths.Count > 0)
                return 0;

            var edits = MinimalTextEditPlanner.Plan(core, target);
            if (edits == null || edits.Count == 0) return 0;

            var start = range.Start;
            if (!SpanTextEquals(range, start, start + coreLength, core)) return 0;

            // From the end backwards, so the offsets of earlier edits stay valid.
            var applied = 0;
            for (var index = edits.Count - 1; index >= 0; index--)
            {
                var edit = edits[index];
                Word.Range? span = null;
                try
                {
                    span = range.Duplicate;
                    span.SetRange(start + edit.Offset, start + edit.Offset + edit.Length);
                    if (edit.Length > 0 &&
                        (!string.Equals(span.Text ?? string.Empty, edit.ExpectedText, StringComparison.Ordinal) ||
                         span.InlineShapes.Count > 0))
                        continue;
                    span.Text = edit.Replacement;
                    applied++;
                }
                finally { Release(span); }
            }
            return applied;
        }

        private static bool SpanTextEquals(Word.Range story, int start, int end, string expected)
        {
            Word.Range? span = null;
            try
            {
                span = story.Duplicate;
                span.SetRange(start, end);
                return span.Start == start && span.End == end &&
                    string.Equals(span.Text ?? string.Empty, expected, StringComparison.Ordinal);
            }
            finally { Release(span); }
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
