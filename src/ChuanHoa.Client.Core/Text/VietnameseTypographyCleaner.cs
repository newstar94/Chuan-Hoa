using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Text
{
    /// <summary>
    /// Thuật toán chuẩn hóa khoảng trắng, dấu câu và dấu ngoặc kép tiếng Việt.
    /// Hoạt động thuần túy trên chuỗi ký tự, siêu tốc và độc lập với Word COM.
    /// </summary>
    public static class VietnameseTypographyCleaner
    {
        private static readonly Regex MultipleSpacesRegex = new Regex(@"[ ]{2,}", RegexOptions.Compiled);
        private static readonly Regex SpaceBeforePunctuationRegex = new Regex(@"[ ]+([,.:;!?])", RegexOptions.Compiled);
        // Exclude space insertion when a digit follows a comma (decimal: 1,5)
        // or a digit follows a colon (time: 14:30). Supports \uE000 placeholder for shielded URLs/emails.
        private static readonly Regex SpaceAfterPunctuationRegex = new Regex(@"([,](?!\d)|[;!?]|:(?!//|\d))([A-Za-zÀ-ỹ]|\uE000)", RegexOptions.Compiled);
        // Do not split compound abbreviations without spaces (e.g. v.v, PGS.TS, GS.TS, BS.CKI, BS.CKII, Th.S, TP.HCM).
        // For standalone titles (e.g. ThS. Nguyễn, TS. Lê), space is inserted normally.
        private static readonly Regex PeriodFollowedByLetterRegex = new Regex(
            @"(?<!\b(?:v|PGS|GS|BS|Th|TP))\.(?!(?:v\b|TS\b|S\b|CKI\b|CKII\b|HCM\b))([A-Za-zÀ-ỹ]|\uE000)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex UrlOrEmailRegex = new Regex(
            @"(?:https?://[^\s<>""'()]+|www\.[^\s<>""'()]+|[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|(?<![@\w.-])(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+(?:vn|com|org|net|gov|edu)(?:\.vn)?(?:/[^\s<>""'()]*)?)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex InsideParenthesesOpenRegex = new Regex(@"\(\s+", RegexOptions.Compiled);
        private static readonly Regex InsideParenthesesCloseRegex = new Regex(@"\s+\)", RegexOptions.Compiled);
        private static readonly Regex InsideBracketsOpenRegex = new Regex(@"\[\s+", RegexOptions.Compiled);
        private static readonly Regex InsideBracketsCloseRegex = new Regex(@"\s+\]", RegexOptions.Compiled);

        /// <summary>
        /// Dọn dẹp khoảng trắng và chuẩn hóa khoảng cách trước/sau dấu câu.
        /// </summary>
        public static string CleanWhitespaceAndPunctuation(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            // Xử lý bảo toàn ký tự ngắt đoạn của Word nếu có (\r, \n, \a)
            var suffix = string.Empty;
            var core = text!;
            if (core.EndsWith("\r\n", StringComparison.Ordinal))
            {
                suffix = "\r\n";
                core = core.Substring(0, core.Length - 2);
            }
            else if (core.EndsWith("\r", StringComparison.Ordinal) || core.EndsWith("\n", StringComparison.Ordinal) || core.EndsWith("\a", StringComparison.Ordinal))
            {
                suffix = core.Substring(core.Length - 1);
                core = core.Substring(0, core.Length - 1);
            }

            // Thay ký tự non-breaking space (U+00A0) thành space thông thường
            core = core.Replace('\u00A0', ' ');

            // Bảo tồn URL và Email trước khi xử lý khoảng trắng và dấu câu
            var urlPlaceholders = new System.Collections.Generic.List<string>();
            core = UrlOrEmailRegex.Replace(core, match =>
            {
                var val = match.Value;
                var trailing = string.Empty;
                while (val.Length > 0 && (val[val.Length - 1] == '.' || val[val.Length - 1] == ',' ||
                                          val[val.Length - 1] == ';' || val[val.Length - 1] == ':' ||
                                          val[val.Length - 1] == '!' || val[val.Length - 1] == '?'))
                {
                    trailing = val[val.Length - 1] + trailing;
                    val = val.Substring(0, val.Length - 1);
                }
                if (val.Length == 0) return match.Value;
                var placeholder = "\uE000__URL_EMAIL_" + urlPlaceholders.Count + "__\uE001";
                urlPlaceholders.Add(val);
                return placeholder + trailing;
            });

            // 1. Gộp 2+ dấu cách liên tiếp thành 1 dấu cách
            core = MultipleSpacesRegex.Replace(core, " ");

            // 2. Xóa khoảng trắng đứng trước dấu câu (, . : ; ! ?)
            core = SpaceBeforePunctuationRegex.Replace(core, "$1");

            // 3. Thêm khoảng trắng sau dấu câu nếu người gõ viết liền (ngoại trừ URL như http://)
            core = SpaceAfterPunctuationRegex.Replace(core, "$1 $2");
            core = PeriodFollowedByLetterRegex.Replace(core, ". $1");

            // 4. Xóa khoảng cách thừa sát mép trong dấu ngoặc đơn và ngoặc vuông
            core = InsideParenthesesOpenRegex.Replace(core, "(");
            core = InsideParenthesesCloseRegex.Replace(core, ")");
            core = InsideBracketsOpenRegex.Replace(core, "[");
            core = InsideBracketsCloseRegex.Replace(core, "]");

            // 5. Khôi phục URL và Email đã bảo vệ
            if (urlPlaceholders.Count > 0)
            {
                for (var i = 0; i < urlPlaceholders.Count; i++)
                {
                    core = core.Replace("\uE000__URL_EMAIL_" + i + "__\uE001", urlPlaceholders[i]);
                }
            }

            // 6. Cắt khoảng trắng thừa ở đầu và cuối dòng
            core = core.Trim();

            return core + suffix;
        }

        /// <summary>
        /// Chuẩn hóa dấu ngoặc kép thẳng (") thành dấu ngoặc kép cong mở/đóng chuẩn tiếng Việt (“ ”).
        /// Tự động xóa khoảng cách thừa sát mép trong của dấu ngoặc.
        /// </summary>
        public static string NormalizeQuotationMarks(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var suffix = string.Empty;
            var core = text!;
            if (core.EndsWith("\r\n", StringComparison.Ordinal))
            {
                suffix = "\r\n";
                core = core.Substring(0, core.Length - 2);
            }
            else if (core.EndsWith("\r", StringComparison.Ordinal) || core.EndsWith("\n", StringComparison.Ordinal) || core.EndsWith("\a", StringComparison.Ordinal))
            {
                suffix = core.Substring(core.Length - 1);
                core = core.Substring(0, core.Length - 1);
            }

            // An unmatched ASCII quote is ambiguous. Keep the input intact rather
            // than let one missing quote invert all remaining open/close marks.
            if (core.Count(character => character == '"') % 2 != 0)
            {
                return core + suffix;
            }

            var sb = new StringBuilder(core.Length);
            var inQuotes = false;

            for (var i = 0; i < core.Length; i++)
            {
                var c = core[i];
                if (c == '"')
                {
                    if (!inQuotes)
                    {
                        sb.Append('“'); // U+201C: Left double quotation mark
                        inQuotes = true;
                    }
                    else
                    {
                        sb.Append('”'); // U+201D: Right double quotation mark
                        inQuotes = false;
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }

            var result = sb.ToString();

            // Dọn khoảng trắng sát mép trong ngoặc kép: “ abc ” ➔ “abc”
            result = Regex.Replace(result, @"“\s+", "“");
            result = Regex.Replace(result, @"\s+”", "”");

            // Đảm bảo có khoảng trắng phía ngoài ngoặc kép nếu liền kề chữ cái
            result = Regex.Replace(result, @"([A-Za-zÀ-ỹ0-9])(“)", "$1 $2");
            result = Regex.Replace(result, @"(”)([A-Za-zÀ-ỹ0-9])", "$1 $2");

            return result + suffix;
        }
    }
}
