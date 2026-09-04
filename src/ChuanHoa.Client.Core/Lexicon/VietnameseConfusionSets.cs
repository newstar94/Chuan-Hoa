using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ChuanHoa.Client.Core.Lexicon
{
    /// <summary>
    /// Provides phonetic, typing, and regional confusion sets commonly occurring in
    /// Vietnamese administrative and daily texts.
    /// </summary>
    public static class VietnameseConfusionSets
    {
        private static readonly CultureInfo VietnameseCulture = CultureInfo.GetCultureInfo("vi-VN");

        public static readonly IReadOnlyDictionary<string, string[]> InitialConsonantConfusion =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "s", new[] { "x" } },
                { "x", new[] { "s" } },
                { "ch", new[] { "tr" } },
                { "tr", new[] { "ch" } },
                { "d", new[] { "gi", "r" } },
                { "gi", new[] { "d", "r" } },
                { "r", new[] { "d", "gi" } },
                { "l", new[] { "n" } },
                { "n", new[] { "l" } }
            };

        public static readonly IReadOnlyDictionary<string, string[]> FinalConsonantConfusion =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "c", new[] { "t" } },
                { "t", new[] { "c" } },
                { "n", new[] { "ng" } },
                { "ng", new[] { "n" } }
            };

        private static readonly IReadOnlyDictionary<char, char> ToneHoiNgaMap =
            new Dictionary<char, char>
            {
                { 'ả', 'ã' }, { 'ã', 'ả' },
                { 'ẳ', 'ẵ' }, { 'ẵ', 'ẳ' },
                { 'ẩ', 'ẫ' }, { 'ẫ', 'ẩ' },
                { 'ẻ', 'ẽ' }, { 'ẽ', 'ẻ' },
                { 'ể', 'ễ' }, { 'ễ', 'ể' },
                { 'ỉ', 'ĩ' }, { 'ĩ', 'ỉ' },
                { 'ỏ', 'õ' }, { 'õ', 'ỏ' },
                { 'ổ', 'ỗ' }, { 'ỗ', 'ổ' },
                { 'ở', 'ỡ' }, { 'ỡ', 'ở' },
                { 'ủ', 'ũ' }, { 'ũ', 'ủ' },
                { 'ử', 'ữ' }, { 'ữ', 'ử' },
                { 'ỷ', 'ỹ' }, { 'ỹ', 'ỷ' },
                { 'Ả', 'Ã' }, { 'Ã', 'Ả' },
                { 'Ẳ', 'Ẵ' }, { 'Ẵ', 'Ẳ' },
                { 'Ẩ', 'Ẫ' }, { 'Ẫ', 'Ẩ' },
                { 'Ẻ', 'Ẽ' }, { 'Ẽ', 'Ẻ' },
                { 'Ể', 'Ễ' }, { 'Ễ', 'Ể' },
                { 'Ỉ', 'Ĩ' }, { 'Ĩ', 'Ỉ' },
                { 'Ỏ', 'Õ' }, { 'Õ', 'Ỏ' },
                { 'Ổ', 'Ỗ' }, { 'Ỗ', 'Ổ' },
                { 'Ở', 'Ỡ' }, { 'Ỡ', 'Ở' },
                { 'Ủ', 'Ũ' }, { 'Ũ', 'Ủ' },
                { 'Ử', 'Ữ' }, { 'Ữ', 'Ử' },
                { 'Ỷ', 'Ỹ' }, { 'Ỹ', 'Ỷ' }
            };

        public static readonly IReadOnlyDictionary<string, string> AdministrativeConfusionPairs =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // L / N confusion
                { "tài niệu", "tài liệu" },
                { "lăng lực", "năng lực" },
                { "lội dung", "nội dung" },
                { "kỷ nuật", "kỷ luật" },
                { "kỉ nuật", "kỉ luật" },
                { "niên quan", "liên quan" },
                { "lập nuận", "lập luận" },
                { "nực lượng", "lực lượng" },
                { "độc nập", "độc lập" },
                { "nưu ý", "lưu ý" },
                { "lỗ lực", "nỗ lực" },
                { "lề nếp", "nề nếp" },
                { "lêu gương", "nêu gương" },
                { "lăng suất", "năng suất" },
                { "ninh hoạt", "linh hoạt" },
                { "lòng cốt", "nòng cốt" },
                { "nãnh đạo", "lãnh đạo" },
                { "nghị nực", "nghị lực" },
                { "quyết niệt", "quyết liệt" },
                { "quản ný", "quản lý" },
                { "xử ný", "xử lý" },
                { "khả lăng", "khả năng" },
                { "kỹ lăng", "kỹ năng" },
                { "kĩ lăng", "kĩ năng" },
                { "chức lăng", "chức năng" },
                { "tiềm lăng", "tiềm năng" },

                // S / X confusion
                { "sử lý", "xử lý" },
                { "bổ xung", "bổ sung" },
                { "sắp sếp", "sắp xếp" },
                { "sát nhập", "sáp nhập" },
                { "chia sẽ", "chia sẻ" },
                { "sơ xuất", "sơ suất" },
                { "sâu xát", "sâu sát" },
                { "sét duyệt", "xét duyệt" },
                { "suất sắc", "xuất sắc" },
                { "xai xót", "sai sót" },
                { "xửa chữa", "sửa chữa" },
                { "xản xuất", "sản xuất" },
                { "súc tiến", "xúc tiến" },
                { "sây dựng", "xây dựng" },
                { "giám xát", "giám sát" },
                { "cơ xở", "cơ sở" },
                { "phát xinh", "phát sinh" },

                // TR / CH confusion
                { "chân trọng", "trân trọng" },
                { "trính trực", "chính trực" },
                { "tập chung", "tập trung" },
                { "chỉ chích", "chỉ trích" },
                { "chung thực", "trung thực" },
                { "chách nhiệm", "trách nhiệm" },
                { "tri trả", "chi trả" },
                { "tri tiết", "chi tiết" },
                { "trính sách", "chính sách" },
                { "trủ trương", "chủ trương" },
                { "chực thuộc", "trực thuộc" },
                { "cha cứu", "tra cứu" },
                { "triển giao", "chuyển giao" },
                { "chiển khai", "triển khai" },
                { "chọng tâm", "trọng tâm" },

                // D / GI / R confusion
                { "bàn dao", "bàn giao" },
                { "dút gọn", "rút gọn" },
                { "dút kinh nghiệm", "rút kinh nghiệm" },
                { "dao nhận", "giao nhận" },
                { "dao dịch", "giao dịch" },
                { "dành mạch", "rành mạch" },
                { "dữ dìn", "giữ gìn" },
                { "giàn xếp", "dàn xếp" },
                { "dàng buộc", "ràng buộc" },
                { "dõ dàng", "rõ ràng" },
                { "dám đốc", "giám đốc" },
                { "da hạn", "gia hạn" },

                // Hỏi / Ngã & Vần
                { "điều khoảng", "điều khoản" },
                { "hướng dẩn", "hướng dẫn" },
                { "chuẩn bị kĩ", "chuẩn bị kỹ" },
                { "nghỉa vụ", "nghĩa vụ" },
                { "diển biến", "diễn biến" },
                { "tiến triễn", "tiến triển" },
                { "kiễm điểm", "kiểm điểm" },
                { "kết qủa", "kết quả" }
            };

        /// <summary>
        /// Generates phonetically swapped candidate syllables for a given syllable word.
        /// </summary>
        public static IEnumerable<string> GeneratePhoneticAlternates(string syllable)
        {
            if (string.IsNullOrWhiteSpace(syllable))
                yield break;

            var normalized = syllable.Normalize(NormalizationForm.FormC);
            var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1. Swap hỏi / ngã
            var chars = normalized.ToCharArray();
            var toneSwapped = false;
            for (var i = 0; i < chars.Length; i++)
            {
                if (ToneHoiNgaMap.TryGetValue(chars[i], out var swappedChar))
                {
                    chars[i] = swappedChar;
                    toneSwapped = true;
                }
            }
            if (toneSwapped)
            {
                var alternate = new string(chars);
                if (results.Add(alternate))
                    yield return alternate;
            }

            // 2. Swap initial consonant
            foreach (var pair in InitialConsonantConfusion)
            {
                if (normalized.StartsWith(pair.Key, StringComparison.OrdinalIgnoreCase))
                {
                    var isFirstUpper = char.IsUpper(normalized[0]);
                    var remainder = normalized.Substring(pair.Key.Length);
                    foreach (var replacement in pair.Value)
                    {
                        var candidate = isFirstUpper
                            ? char.ToUpper(replacement[0], VietnameseCulture) + replacement.Substring(1) + remainder
                            : replacement + remainder;
                        if (results.Add(candidate))
                            yield return candidate;
                    }
                }
            }

            // 3. Swap final consonant
            foreach (var pair in FinalConsonantConfusion)
            {
                if (normalized.EndsWith(pair.Key, StringComparison.OrdinalIgnoreCase) && normalized.Length > pair.Key.Length)
                {
                    var prefix = normalized.Substring(0, normalized.Length - pair.Key.Length);
                    foreach (var replacement in pair.Value)
                    {
                        var candidate = prefix + replacement;
                        if (results.Add(candidate))
                            yield return candidate;
                    }
                }
            }
        }
    }
}
