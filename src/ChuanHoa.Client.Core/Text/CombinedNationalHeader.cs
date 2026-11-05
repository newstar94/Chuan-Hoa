using System;
using System.Linq;

namespace ChuanHoa.Client.Core.Text
{
    public static class CombinedNationalHeader
    {
        // Only split complete, consecutive header lines. Never split ordinary
        // prose, fields or an arbitrary occurrence of the national title.
        public static int[] GetBreakOffsets(string text)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<int>();
            var breaks = Enumerable.Range(0, text.Length)
                .Where(i => text[i] == '\v' || text[i] == '\n').ToArray();
            if (breaks.Length < 1 || breaks.Length > 2) return Array.Empty<int>();
            // Compare normalized components, but return offsets in the ORIGINAL
            // UTF-16 text so NFD and invisible characters never move Word anchors.
            if (!ComponentTextNormalizer.IsNationalTitle(text.Substring(0, breaks[0])) ||
                !ComponentTextNormalizer.IsNationalMotto(text.Substring(breaks[0] + 1)))
                return Array.Empty<int>();
            if (breaks.Length == 2 &&
                ComponentTextNormalizer.Normalize(text.Substring(breaks[1] + 1)).Length != 0)
                return Array.Empty<int>();
            return breaks;
        }
    }
}
