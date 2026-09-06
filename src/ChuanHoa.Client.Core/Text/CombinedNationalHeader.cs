using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Text
{
    public static class CombinedNationalHeader
    {
        // Only split complete, consecutive header lines. Never split ordinary
        // prose, fields or an arbitrary occurrence of the national title.
        public static int[] GetBreakOffsets(string text)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<int>();
            var match = Regex.Match(text,
                @"^\s*CỘNG\s+H(?:ÒA|OÀ)\s+XÃ\s+HỘI\s+CHỦ\s+NGHĨA\s+VIỆT\s+NAM[ \t]*(?<break>[\v\n])[ \t]*Độc\s+lập[ \t]*[-–—][ \t]*Tự\s+do[ \t]*[-–—][ \t]*Hạnh\s+phúc[ \t]*(?:(?<break>[\v\n])[ \t]*[-_.–—]{3,}[ \t]*)?[\r\a]*$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success
                ? match.Groups["break"].Captures.Cast<Capture>().Select(c => c.Index).ToArray()
                : Array.Empty<int>();
        }
    }
}
