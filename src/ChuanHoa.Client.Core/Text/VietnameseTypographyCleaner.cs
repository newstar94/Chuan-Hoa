using System;
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
        private static readonly Regex SpaceAfterPunctuationRegex = new Regex(@"([,;!?]|:(?!\/\/))([A-Za-zÀ-ỹ])", RegexOptions.Compiled);
        private static readonly Regex PeriodFollowedByLetterRegex = new Regex(@"(?<!\b(?:v\.v|tp|gs|ts|pgs|th\b)\.)\.([A-Za-zÀ-ỹ])", RegexOptions.Compiled | RegexOptions.IgnoreCase);
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

            // 5. Cắt khoảng trắng thừa ở đầu và cuối dòng
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
