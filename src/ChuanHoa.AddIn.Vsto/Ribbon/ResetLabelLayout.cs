using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ChuanHoa.AddIn.Vsto.Ribbon
{
    /// <summary>
    /// Ribbon small buttons align native labels at their leading edge. Balance
    /// the six character-tool labels to the natural width of 100%. Color pickers
    /// live in a separate third row and never determine a character-column width.
    /// </summary>
    internal static class ResetLabelLayout
    {
        private static readonly object Gate = new object();
        private static string? _fontKey;
        private static Dictionary<string, string>? _labels;

        private static readonly KeyValuePair<string, string>[] CharacterLabels =
        {
            new KeyValuePair<string, string>("btnScaleGiam", "A−"),
            new KeyValuePair<string, string>("btnScale100", "100%"),
            new KeyValuePair<string, string>("btnScaleTang", "A+"),
            new KeyValuePair<string, string>("btnCoChu", "◀"),
            new KeyValuePair<string, string>("btnGianChuNormal", "⬤"),
            new KeyValuePair<string, string>("btnGianChuRa", "▶"),
        };

        internal static string GetLabel(string controlId)
        {
            lock (Gate)
            {
                EnsureLabels();
                return _labels![controlId];
            }
        }

        private static void EnsureLabels()
        {
            var font = SystemFonts.MenuFont;
            var key = font.Name + "|" + font.SizeInPoints + "|" + font.Style;
            if (_labels != null && key == _fontKey) return;
            _labels = MeasureLabels(font);
            _fontKey = key;
        }

        internal static Dictionary<string, string> MeasureLabels(Font font)
        {
            // Use only the widest natural label. The old extra two symbol
            // advances made the whole three-column block unnecessarily wide.
            var target = Width("100%", font);
            // Different Unicode spaces have different advances. Search both
            // whole native labels: GDI rounding/kerning is not additive.
            var pads = PaddingCandidates();
            var labels = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var entry in CharacterLabels)
            {
                var best = entry.Value;
                var bestError = double.MaxValue;
                foreach (var pad in pads)
                {
                    var candidate = pad + entry.Value + pad;
                    var widthError = Math.Abs(Width(candidate, font) - target);
                    // Equal padding centers each glyph in the same label width.
                    var error = widthError * 100 + pad.Length;
                    if (error >= bestError) continue;
                    best = candidate;
                    bestError = error;
                }
                labels.Add(entry.Key, best);
            }
            return labels;
        }

        private static List<string> PaddingCandidates()
        {
            var pads = new List<string> { string.Empty };
            foreach (var space in new[] { '\u00a0', '\u2002', '\u2009', '\u200a' })
                for (var count = 1; count <= 24; count++)
                    pads.Add(new string(space, count));
            foreach (var baseSpace in new[] { '\u00a0', '\u2002', '\u2009' })
                for (var baseCount = 1; baseCount <= 8; baseCount++)
                    for (var hairCount = 1; hairCount <= 12; hairCount++)
                        pads.Add(new string(baseSpace, baseCount) + new string('\u200a', hairCount));
            return pads;
        }

        private static int Width(string text, Font font) => TextRenderer.MeasureText(text, font,
            new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding |
            TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
    }
}
