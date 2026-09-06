using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ChuanHoa.AddIn.Vsto.Ribbon
{
    /// <summary>
    /// Ribbon small buttons align native labels at their leading edge. Balance
    /// the reset labels using measured non-breaking typographic spaces, keeping
    /// the original glyph/font and avoiding Office bitmap rescaling entirely.
    /// </summary>
    internal static class ResetLabelLayout
    {
        private static readonly object Gate = new object();
        private static string? _fontKey;
        private static string[]? _labels;

        internal static string GetLabel(string controlId)
        {
            lock (Gate)
            {
                var font = SystemFonts.MenuFont;
                var key = font.Name + "|" + font.SizeInPoints + "|" + font.Style;
                if (_labels == null || key != _fontKey)
                {
                    _labels = MeasureLabels(font);
                    _fontKey = key;
                }
                return _labels[controlId == "btnScale100" ? 0 : 1];
            }
        }

        internal static string[] MeasureLabels(Font font)
        {
            var target = Width("100%", font);
            var best = "⬤";
            var bestError = double.MaxValue;
            // Different Unicode spaces have different advances. Search both
            // whole native labels: GDI rounding/kerning is not additive.
            var pads = new List<string> { string.Empty };
            foreach (var space in new[] { '\u00a0', '\u2002', '\u2009', '\u200a' })
                for (var count = 1; count <= 8; count++)
                    pads.Add(new string(space, count));
            foreach (var pad in pads)
            {
                var candidate = pad + "⬤" + pad;
                var widthError = Math.Abs(Width(candidate, font) - target);
                // Keep equal padding on both sides, so the dot is centered in
                // its own label while its total width tracks the upper label.
                var error = widthError * 100 + pad.Length;
                if (error >= bestError) continue;
                best = candidate;
                bestError = error;
            }
            return new[] { "100%", best };
        }

        private static int Width(string text, Font font) => TextRenderer.MeasureText(text, font,
            new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding |
            TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
    }
}
