using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>One deterministic vocabulary shared by role and citation parsing.</summary>
    internal static class VietnameseLegalDocumentVocabulary
    {
        private static readonly string[] Names =
        {
            "Nghị quyết liên tịch", "Thông tư liên tịch", "Hiến pháp", "Bộ luật",
            "Pháp lệnh", "Nghị quyết", "Nghị định", "Quyết định", "Chỉ thị",
            "Thông tư", "Công văn", "Tờ trình", "Luật", "Lệnh"
        };

        public static readonly string RegexAlternation = string.Join("|", Names
            .OrderByDescending(value => value.Length)
            .Select(value => Regex.Escape(value).Replace(@"\ ", @"\s+")));

        public static bool Contains(string text)
        {
            return Names.Any(name => (text ?? string.Empty).IndexOf(name,
                StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public static IReadOnlyList<string> All => Names;
    }
}
