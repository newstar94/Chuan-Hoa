using System;
using System.Globalization;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Recognizes only the explicit names emitted by released Chuẩn hóa builds.
    /// A generic Word Shape name must never become owned merely because it is near
    /// a legal component or happens to begin with a similar string.
    /// </summary>
    public static class LineShapeOwnership
    {
        private static readonly string[] CurrentPrefixes =
        {
            "CHUANHOA2_ORG_", "CHUANHOA2_SUBJ_",
            "CHUANHOA2_PARTY_", "CHUANHOA2_MOTTO_"
        };

        private static readonly string[] LegacyPrefixes =
        {
            "CHUANHOA_ORG_", "CHUANHOA_SUBJ_",
            "CHUANHOA_PARTY_", "CHUANHOA_MOTTO_"
        };

        public static bool IsCurrentOwned(string? name) => HasKnownPrefix(name, CurrentPrefixes);

        public static bool IsLegacyOwned(string? name) => HasKnownPrefix(name, LegacyPrefixes);

        public static bool IsOwned(string? name) => IsCurrentOwned(name) || IsLegacyOwned(name);

        public static bool IsOwnedForParagraph(string? name, int paragraphIndex)
        {
            if (!IsOwned(name)) return false;
            var marker = "P" + paragraphIndex.ToString(CultureInfo.InvariantCulture);
            return name!.EndsWith(marker, StringComparison.Ordinal) ||
                name!.IndexOf(marker + "_", StringComparison.Ordinal) >= 0;
        }

        private static bool HasKnownPrefix(string? name, string[] prefixes)
        {
            if (string.IsNullOrEmpty(name)) return false;
            foreach (var prefix in prefixes)
                if (name!.StartsWith(prefix, StringComparison.Ordinal)) return true;
            return false;
        }
    }
}
