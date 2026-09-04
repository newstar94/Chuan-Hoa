using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Lexicon;
using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.Client.Core.Scanning
{
    public sealed class LocalDocumentScanner
    {
        private readonly CanonicalRuleScanner _canonical;
        private const double PointsPerMillimeter = 72.0d / 25.4d;
        private static readonly Regex RepeatedWhitespace = new Regex("[ \\t]{2,}", RegexOptions.CultureInvariant);
        private static readonly Regex WhitespaceBeforePunctuation = new Regex("[ \\t]+(?=[,.;:!?])", RegexOptions.CultureInvariant);
        private static readonly Regex SentenceLowercase = new Regex(@"[.!?][ \t]+(?<letter>\p{Ll})", RegexOptions.CultureInvariant);
        private static readonly Regex LowercaseSignerAbbreviation = new Regex(@"(?<!\p{L})(?<abbr>tm|kt|tl|tuq|q)\.(?=\s|$)", RegexOptions.CultureInvariant);
        private static readonly Regex CodeNumber = new Regex(@"^\s*Số\s*:?[ \t]*(?<number>\d+)(?<notation>[^\r\n]*)$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        private static readonly Regex PlaceDate = new Regex(@"^\s*(?<place>[^,]+)(?<comma>,?)[ \t]*ngày[ \t]+(?<day>\d{1,2})[ \t]+tháng[ \t]+(?<month>\d{1,2})[ \t]+năm[ \t]+(?<year>\d{4})", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public LocalDocumentScanner(PersonalDictionaryManager? personalDictionary = null)
        {
            _canonical = new CanonicalRuleScanner(personalDictionary);
        }

        public LocalScanResult ScanFormat(LocalScanSnapshot snapshot, LocalRulePack rules,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            Validate(snapshot, rules);
            var findings = _canonical.ScanFormat(snapshot, rules, cancellationToken);
            return Result("format", snapshot, rules, findings);
        }

        public LocalScanResult ScanSpelling(LocalScanSnapshot snapshot, LocalRulePack rules,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            Validate(snapshot, rules);
            return Result("spelling", snapshot, rules,
                _canonical.ScanSpelling(snapshot, rules, cancellationToken));
        }

        private static IEnumerable<LocalParagraphSnapshot> Scannable(LocalScanSnapshot snapshot) =>
            snapshot.Paragraphs.Where(item => !string.IsNullOrWhiteSpace(item.Text));

        private static void Validate(LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            if (string.IsNullOrWhiteSpace(snapshot.DocumentFingerprint)) throw new InvalidOperationException("DOCUMENT_FINGERPRINT_REQUIRED");
        }

        private static LocalScanResult Result(string lane, LocalScanSnapshot snapshot, LocalRulePack rules,
            IReadOnlyList<AnnotationFinding> findings) =>
            new LocalScanResult(Guid.NewGuid().ToString("D"), lane, rules.PackId, snapshot.DocumentFingerprint,
                snapshot.Revision, findings);

        private static void AddRegexFindings(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph,
            Regex regex, string code, string issue, string expected, string packId, string? replacement)
        {
            foreach (Match match in regex.Matches(paragraph.Text))
            {
                if (!match.Success || match.Length <= 0) continue;
                findings.Add(TextFinding(code, paragraph, match.Index, match.Length, issue, expected,
                    "Gói quy tắc " + packId));
            }
        }

        private static void AddComponentFormatFinding(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            var normalized = CollapseWhitespace(paragraph.Text).Trim();
            var issues = new List<string>();
            string component;
            double minSize;
            double maxSize;
            bool requireBold;
            bool requireItalic;
            if (string.Equals(normalized, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", StringComparison.OrdinalIgnoreCase))
            {
                component = "Quốc hiệu"; minSize = 12; maxSize = 13; requireBold = true; requireItalic = false;
                if (!string.Equals(normalized, normalized.ToUpper(new CultureInfo("vi-VN")), StringComparison.Ordinal)) issues.Add("phải viết hoa");
            }
            else if (string.Equals(normalized, "ĐỘC LẬP - TỰ DO - HẠNH PHÚC", StringComparison.OrdinalIgnoreCase))
            {
                component = "Tiêu ngữ"; minSize = 13; maxSize = 14; requireBold = true; requireItalic = false;
                if (!string.Equals(normalized, "Độc lập - Tự do - Hạnh phúc", StringComparison.Ordinal)) issues.Add("sai viết hoa hoặc khoảng cách quanh gạch nối");
            }
            else if (string.Equals(normalized, "ĐẢNG CỘNG SẢN VIỆT NAM", StringComparison.OrdinalIgnoreCase))
            {
                component = "Tiêu đề Đảng"; minSize = 15; maxSize = 15; requireBold = true; requireItalic = false;
                if (!string.Equals(normalized, normalized.ToUpper(new CultureInfo("vi-VN")), StringComparison.Ordinal)) issues.Add("phải viết hoa");
            }
            else return;

            if (!string.IsNullOrWhiteSpace(paragraph.FontName) &&
                !string.Equals(paragraph.FontName, rules.BodyFontName, StringComparison.OrdinalIgnoreCase))
                issues.Add("phông chữ phải là " + rules.BodyFontName);
            if (paragraph.FontSizePoints.HasValue &&
                (paragraph.FontSizePoints.Value < minSize - 0.1d || paragraph.FontSizePoints.Value > maxSize + 0.1d))
                issues.Add("cỡ chữ phải trong khoảng " + Range(minSize, maxSize) + " pt");
            if (requireBold && paragraph.Bold.HasValue && !paragraph.Bold.Value) issues.Add("phải in đậm");
            if (!requireItalic && paragraph.Italic.HasValue && paragraph.Italic.Value) issues.Add("không được in nghiêng");
            if (paragraph.Alignment.HasValue && paragraph.Alignment.Value != 1) issues.Add("phải căn giữa");
            if (issues.Count == 0) return;
            findings.Add(ParagraphFinding("FORMAT-COMPONENT-STYLE", paragraph,
                component + " chưa đúng thể thức: " + string.Join(", ", issues) + ".",
                "Định dạng lại " + component + " theo bộ quy tắc.", "Gói quy tắc " + rules.PackId));
        }

        private static void AddBodyStyleFinding(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            if (!IsLikelyBodyParagraph(paragraph)) return;
            var issues = new List<string>();
            if (!string.IsNullOrWhiteSpace(paragraph.FontName) &&
                !string.Equals(paragraph.FontName, rules.BodyFontName, StringComparison.OrdinalIgnoreCase))
                issues.Add("phông chữ phải là " + rules.BodyFontName);
            if (paragraph.FontSizePoints.HasValue &&
                (paragraph.FontSizePoints.Value < rules.BodyFontMinPoints - 0.1d ||
                 paragraph.FontSizePoints.Value > rules.BodyFontMaxPoints + 0.1d))
                issues.Add("cỡ chữ ngoài " + Range(rules.BodyFontMinPoints, rules.BodyFontMaxPoints) + " pt");
            if (paragraph.Alignment.HasValue && paragraph.Alignment.Value != rules.BodyAlignment)
                issues.Add("chưa căn đều hai lề");
            if (paragraph.FirstLineIndentPoints.HasValue)
            {
                var indentMm = paragraph.FirstLineIndentPoints.Value / PointsPerMillimeter;
                if (indentMm < rules.BodyFirstLineIndentMinMm - 0.5d || indentMm > rules.BodyFirstLineIndentMaxMm + 0.5d)
                    issues.Add("thụt đầu dòng ngoài " + Range(rules.BodyFirstLineIndentMinMm, rules.BodyFirstLineIndentMaxMm) + " mm");
            }
            if (paragraph.SpaceAfterPoints.HasValue && paragraph.SpaceAfterPoints.Value + 0.1d < rules.BodySpaceAfterMinPoints)
                issues.Add("khoảng cách sau đoạn nhỏ hơn " + rules.BodySpaceAfterMinPoints.ToString("0.#", CultureInfo.InvariantCulture) + " pt");
            if (issues.Count == 0) return;
            findings.Add(ParagraphFinding("FORMAT-BODY-STYLE", paragraph,
                "Đoạn nội dung chưa đúng thể thức: " + string.Join(", ", issues) + ".",
                "Căn đều hai lề, thụt đầu dòng và dùng cỡ chữ/khoảng cách theo gói quy tắc.",
                "Gói quy tắc " + rules.PackId));
        }

        private static void AddCodeNumberFindings(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            var match = CodeNumber.Match(paragraph.Text);
            if (!match.Success) return;
            int number;
            if (int.TryParse(match.Groups["number"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out number) &&
                number >= 0 && number < 10 && match.Groups["number"].Value.Length < 2)
                findings.Add(TextFinding("FORMAT-CODE-NUMBER-PAD", paragraph, match.Groups["number"].Index,
                    match.Groups["number"].Length, "Số văn bản nhỏ hơn 10 chưa có số 0 phía trước.",
                    "Viết đủ hai chữ số, ví dụ 05.", "Gói quy tắc " + rules.PackId));
            var notation = match.Groups["notation"];
            if (notation.Success && notation.Length > 0 && notation.Value.Any(char.IsLower))
                findings.Add(TextFinding("FORMAT-CODE-NOTATION-UPPERCASE", paragraph, notation.Index, notation.Length,
                    "Ký hiệu văn bản có chữ thường.", "Viết hoa các nhóm ký hiệu văn bản.", "Gói quy tắc " + rules.PackId));
        }

        private static void AddPlaceDateFindings(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            var match = PlaceDate.Match(paragraph.Text);
            if (!match.Success) return;
            if (match.Groups["comma"].Length == 0)
                findings.Add(TextFinding("FORMAT-PLACE-DATE-COMMA", paragraph, match.Groups["place"].Index,
                    match.Groups["place"].Length, "Thiếu dấu phẩy sau địa danh.",
                    "Thêm dấu phẩy sau địa danh.", "Gói quy tắc " + rules.PackId));
            foreach (var name in new[] { "day", "month" })
            {
                var group = match.Groups[name];
                int value;
                if (group.Length == 1 && int.TryParse(group.Value, out value) && value < 10)
                    findings.Add(TextFinding("FORMAT-PLACE-DATE-PAD", paragraph, group.Index, group.Length,
                        (name == "day" ? "Ngày" : "Tháng") + " nhỏ hơn 10 chưa có số 0 phía trước.",
                        "Viết đủ hai chữ số.", "Gói quy tắc " + rules.PackId));
            }
        }

        private static void AddCapitalizationFindings(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, string packId)
        {
            var text = paragraph.Text;
            if (string.IsNullOrWhiteSpace(text) || StartsWithLowercaseAllowedComponent(text)) return;
            var first = FirstLetterIndex(text);
            if (first >= 0 && char.IsLower(text[first]))
                findings.Add(TextFinding("SPELLING-SENTENCE-CAPITALIZATION", paragraph, first, 1,
                    "Đầu đoạn chưa viết hoa.", "Viết hoa chữ cái đầu câu.", "Gói quy tắc " + packId));
            foreach (Match match in SentenceLowercase.Matches(text))
            {
                var letter = match.Groups["letter"];
                var signer = LowercaseSignerAbbreviation.Match(text, letter.Index);
                if (!letter.Success || IsKnownAbbreviationBefore(text, match.Index) ||
                    (signer.Success && signer.Index == letter.Index)) continue;
                findings.Add(TextFinding("SPELLING-SENTENCE-CAPITALIZATION", paragraph, letter.Index, 1,
                    "Chữ cái đầu câu chưa viết hoa.", "Viết hoa chữ cái đầu câu.", "Gói quy tắc " + packId));
            }
            foreach (Match match in LowercaseSignerAbbreviation.Matches(text))
            {
                var abbreviation = match.Groups["abbr"];
                findings.Add(TextFinding("SPELLING-ABBREVIATION-UPPERCASE", paragraph, abbreviation.Index,
                    abbreviation.Length, "Chữ viết tắt về thẩm quyền ký đang viết thường.",
                    "Viết hoa chữ viết tắt này.", "Gói quy tắc " + packId));
            }
        }

        private static bool IsLikelyBodyParagraph(LocalParagraphSnapshot paragraph)
        {
            if (paragraph.IsInTable || !string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal)) return false;
            var text = CollapseWhitespace(paragraph.Text).Trim();
            if (text.Length < 40 || text.Length > 2000) return false;
            if (string.Equals(text, text.ToUpper(new CultureInfo("vi-VN")), StringComparison.Ordinal)) return false;
            if (CodeNumber.IsMatch(text) || PlaceDate.IsMatch(text)) return false;
            if (Regex.IsMatch(text, @"^(Điều|Khoản|Chương|Phần|Mục|Nơi nhận|Kính gửi|TM\.|KT\.|TL\.|TUQ\.)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) return false;
            return true;
        }

        private static bool IsRecognizedComponent(string text)
        {
            var normalized = CollapseWhitespace(text).Trim();
            return string.Equals(normalized, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "ĐỘC LẬP - TỰ DO - HẠNH PHÚC", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "ĐẢNG CỘNG SẢN VIỆT NAM", StringComparison.OrdinalIgnoreCase);
        }

        private static int FirstLetterIndex(string text)
        {
            for (var index = 0; index < text.Length; index++) if (char.IsLetter(text[index])) return index;
            return -1;
        }

        private static bool StartsWithLowercaseAllowedComponent(string text)
        {
            var value = text.TrimStart();
            return value.StartsWith("về việc", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("V/v", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("e-mail", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsKnownAbbreviationBefore(string text, int punctuationIndex)
        {
            var start = punctuationIndex - 1;
            while (start >= 0 && !char.IsWhiteSpace(text[start])) start--;
            var token = text.Substring(start + 1, punctuationIndex - start);
            return new[] { "TM.", "KT.", "TL.", "TUQ.", "Q.", "TS.", "ThS.", "PGS.", "GS.", "TP.", "P.", "v.v." }
                .Any(item => string.Equals(item, token, StringComparison.OrdinalIgnoreCase));
        }

        private static string CollapseWhitespace(string value) => Regex.Replace(value ?? string.Empty, @"\s+", " ");

        private static IEnumerable<Tuple<int, int>> WholePhraseMatches(string text, string phrase)
        {
            if (string.IsNullOrEmpty(phrase)) yield break;
            var offset = 0;
            while (offset <= text.Length - phrase.Length)
            {
                var index = CultureInfo.GetCultureInfo("vi-VN").CompareInfo.IndexOf(
                    text, phrase, offset, CompareOptions.IgnoreCase);
                if (index < 0) yield break;
                var beforeOk = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
                var end = index + phrase.Length;
                var afterOk = end == text.Length || !char.IsLetterOrDigit(text[end]);
                if (beforeOk && afterOk) yield return Tuple.Create(index, phrase.Length);
                offset = index + Math.Max(1, phrase.Length);
            }
        }

        private static string ApplyCase(string source, string target)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target)) return target;
            if (source.All(value => !char.IsLetter(value) || char.IsUpper(value))) return target.ToUpper(new CultureInfo("vi-VN"));
            if (char.IsUpper(source[0])) return char.ToUpper(target[0], new CultureInfo("vi-VN")) + target.Substring(1);
            return target;
        }

        private static AnnotationFinding SectionFinding(string code, int section, string issue, string expected, string citation) =>
            new AnnotationFinding(Id(code, section, 0), code, "Error", issue, expected, citation,
                new AnnotationAnchor(AnnotationAnchorKind.Section, "wdMainTextStory", null, null, null, string.Empty, section));

        private static AnnotationFinding ParagraphFinding(string code, LocalParagraphSnapshot paragraph, string issue,
            string expected, string citation) =>
            new AnnotationFinding(Id(code, paragraph.Index, 0), code, "Warning", issue, expected, citation,
                Anchor(AnnotationAnchorKind.Paragraph, paragraph, null, null, string.Empty));

        private static AnnotationFinding TextFinding(string code, LocalParagraphSnapshot paragraph, int offset, int length,
            string issue, string expected, string citation) =>
            new AnnotationFinding(Id(code, paragraph.Index, offset), code, "Warning", issue, expected, citation,
                Anchor(AnnotationAnchorKind.TextSpan, paragraph, offset, length, paragraph.Text.Substring(offset, length)));

        private static AnnotationAnchor Anchor(AnnotationAnchorKind kind, LocalParagraphSnapshot paragraph,
            int? offset, int? length, string expectedText) =>
            new AnnotationAnchor(kind, paragraph.StoryType, paragraph.Index, offset, length, expectedText,
                paragraph.SectionIndex, paragraph.TableIndex, paragraph.RowIndex, paragraph.CellIndex);

        private static string Id(string code, int owner, int offset) => code + "-" + owner + "-" + offset;
        private static bool Near(double left, double right, double tolerance) => Math.Abs(left - right) <= tolerance;
        private static bool BetweenMm(double points, double min, double max)
        {
            var value = points / PointsPerMillimeter;
            return value >= min - 0.5d && value <= max + 0.5d;
        }
        private static string Range(double min, double max) => min.ToString("0.#", CultureInfo.InvariantCulture) + "–" + max.ToString("0.#", CultureInfo.InvariantCulture);
    }
}
