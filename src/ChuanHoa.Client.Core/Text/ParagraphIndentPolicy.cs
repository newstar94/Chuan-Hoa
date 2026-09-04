using System;

namespace ChuanHoa.Client.Core.Text
{
    /// <summary>
    /// Deterministic paragraph-indent policy shared by the Word formatter and tests.
    /// Values are expressed in millimetres so the Word adapter owns the COM conversion.
    /// </summary>
    public static class ParagraphIndentPolicy
    {
        public const double BodyFirstLineMillimeters = 10d;
        public const double ListMarkerMillimeters = 10d;
        public const double ListTextMillimeters = 15d;

        public static bool IsDashListParagraph(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var value = text!.TrimStart();
            if (value.Length < 2 || !IsListMarker(value[0])) return false;

            // A marker must be followed by whitespace. This avoids treating negative
            // values such as "-5" as list items.
            return char.IsWhiteSpace(value[1]);
        }

        private static bool IsListMarker(char value)
        {
            return value == '-' || value == '\u2013' || value == '\u2014' ||
                   value == '\u2022' || value == '\u25AA' || value == '\u25E6';
        }
    }
}
