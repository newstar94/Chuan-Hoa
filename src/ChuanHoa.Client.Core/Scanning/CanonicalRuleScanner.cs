using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.Client.Core.Scanning
{
    public sealed class CanonicalRuleScanner
    {
        private const double PointsPerMillimeter = 72d / 25.4d;
        private const double TolerancePoints = 1.1d;
        private static readonly CultureInfo Vietnamese = CultureInfo.GetCultureInfo("vi-VN");
        private static readonly Regex CodeNumber = Rx(@"^\s*Số\s*(?<colon>:?)\s*(?<number>\d+)(?<separator>[^\p{L}\d]?)(?<notation>[^\r\n]*)$");
        private static readonly Regex PlaceDate = Rx(@"^\s*(?<place>[\p{L}][\p{L}\s.]{0,70}?)(?<comma>,?)\s+ngày\s+(?<day>\d{1,2})\s+tháng\s+(?<month>\d{1,2})\s+năm\s+(?<year>\d{4})\s*$", true);
        private static readonly Regex Article = Rx(@"^\s*Điều\s+(?<number>\d+)\s*\.?(?<title>.*)$", true);
        private static readonly Regex Clause = Rx(@"^\s*(?<number>\d+)\.\s+", false);
        private static readonly Regex Point = Rx(@"^\s*(?<letter>[a-zđ])\)\s+", true);
        private static readonly Regex Citation = Rx(@"(?<type>nghị\s+quyết|nghị\s+định|quyết\s+định|chỉ\s+thị|thông\s+tư|luật|pháp\s+lệnh)(?<so>\s+số)?\s+(?<code>\d+[\wĐđ/-]+)", true);
        private static readonly Regex AbbreviatedDate = Rx(@"ngày\s+(?<day>\d{1,2})[/.-](?<month>\d{1,2})[/.-](?<year>\d{4})", true);
        private static readonly Regex LegalBasisShortDate = Rx(@"\b(?<day>\d{1,2})[/.-](?<month>\d{1,2})[/.-](?<year>\d{4})\b", true);
        private static readonly Regex LeadingListMarkers = Rx(
            @"^\s*(?:(?:[-–—+•·▪◦‣⁃]\s+)|(?:\(\s*(?:\d+(?:\.\d+)*|[a-zđ])\s*\)\s*)|(?:(?:\d+(?:\.\d+)*|[a-zđ])[.)]\s+))+(?<letter>\p{L})",
            true);
        private static readonly Regex VietnameseToken = Rx(@"(?<![\p{L}\p{M}])[\p{L}\p{M}]+(?![\p{L}\p{M}])");
        private static readonly HashSet<string> AllowedForeignWords = new HashSet<string>(
            new[] { "email", "e-mail", "website", "online", "internet", "software", "hardware", "file", "link", "wifi" },
            StringComparer.OrdinalIgnoreCase);

        public static readonly IReadOnlyList<string> RegisteredRuleCodes = new[]
        {
            "LOCAL-TYPO-DICT", "LOCAL-TYPO-LEXICON", "LOCAL-TYPO-HIDDEN", "LOCAL-TYPO-PUNCT", "LOCAL-TYPO-SPACE", "LOCAL-TYPO-TELEX",
            "ND30-PL1-M1-K1", "ND30-PL1-M1-K3", "ND30-PL1-M1-K4-COLOR", "ND30-PL1-M1-K4-FONT", "ND30-PL1-M1-K7",
            "ND30-PL1-M2-K1-C", "ND30-PL1-M2-K1-QH", "ND30-PL1-M2-K1-TN", "ND30-PL1-M2-K1-TN-SEP",
            "ND30-PL1-M2-K2-ORG", "ND30-PL1-M2-K2-SUP", "ND30-PL1-M2-K3-ABBR", "ND30-PL1-M2-K3-CASE", "ND30-PL1-M2-K3-PAD",
            "ND30-PL1-M2-K1-TN-LINE", "ND30-PL1-M2-K2-ORG-LINE", "ND30-PL1-M2-K5A-SUBJ-LINE", "HD05-M1-TITLE-LINE",
            "ND30-PL1-M2-K3-PREFIX", "ND30-PL1-M2-K3-SEP", "ND30-PL1-M2-K3-SPACE", "ND30-PL1-M2-K4-CASE", "ND30-PL1-M2-K4-COMMA",
            "ND30-PL1-M2-K4-PAD", "ND30-PL1-M2-K4-STYLE", "ND30-PL1-M2-K5A-SUBJ", "ND30-PL1-M2-K5A-TYPE",
            "ND30-PL1-M2-K5B-SPACE", "ND30-PL1-M2-K5B-STYLE", "ND30-PL1-M2-K6A-PUNCT", "ND30-PL1-M2-K6A-STYLE",
            "ND30-PL1-M2-K6B-CITE", "ND30-PL1-M2-K6B-DATE", "ND30-PL1-M2-K6B-SO", "ND30-PL1-M2-K6D-ALPHABET",
            "ND30-PL1-M2-K6D-ARTICLE", "ND30-PL1-M2-K6D-CLAUSE", "ND30-PL1-M2-K6D-POINT", "ND30-PL1-M2-K6D-TITLE",
            "ND30-PL1-M2-K6E-ALIGN", "ND30-PL1-M2-K6E-DOTSLASH", "ND30-PL1-M2-K6E-INDENT", "ND30-PL1-M2-K6E-LINESPACING",
            "ND30-PL1-M2-K6E-SPACEAFTER", "ND30-PL1-M2-K7B-AUTH", "ND30-PL1-M2-K7D-STYLE", "ND30-PL1-M2-K9A-COLON",
            "ND30-PL1-M2-K9A-INLINE-END", "ND30-PL1-M2-K9A-LAYOUT", "ND30-PL1-M2-K9A-PUNCT", "ND30-PL1-M2-K9B-LABEL",
            "ND30-PL1-M2-K9B-LIST", "ND30-PL1-M2-K9B-LUU", "ND30-PL1-M3-K1A-NUM", "ND30-PL1-M3-K1A-REF",
            "ND30-PL1-M3-K1B", "ND30-PL1-M3-K1C", "ND30-PL1-M3-K1D", "ND30-PL1-MV-CT1", "ND30-PL2-M1",
            "ND30-PL2-M2-K1", "ND30-PL2-M3-K1A", "ND30-PL2-M3-K1B", "ND30-PL2-M3-K1C", "ND30-PL2-M3-K1D",
            "ND30-PL2-M3-K1E", "ND30-PL2-M4-K1A", "ND30-PL2-M4-K1B", "ND30-PL2-M5-K5", "ND30-PL2-M5-K7", "ND30-PL2-M5-K8A"
        };

        private readonly DocumentRoleDetector _roleDetector = new DocumentRoleDetector();

        public IReadOnlyList<AnnotationFinding> ScanFormat(LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            var findings = new List<AnnotationFinding>();
            var roles = _roleDetector.Detect(snapshot);
            CheckPageSetup(findings, snapshot, rules);
            CheckBodyTypography(findings, snapshot, rules, roles);
            CheckComponents(findings, snapshot, rules, roles);
            CheckHeaderFontSizeTier(findings, snapshot, rules, roles);
            CheckRequiredLineShapes(findings, snapshot, rules, roles);
            CheckCodeNumber(findings, snapshot, rules, roles);
            CheckPlaceDate(findings, snapshot, rules, roles);
            CheckTypeAndSubject(findings, snapshot, rules, roles);
            CheckLegalBasisAndCitations(findings, snapshot, rules, roles);
            CheckStructure(findings, snapshot, rules, roles);
            CheckSignerAndRecipients(findings, snapshot, rules, roles);
            CheckAppendices(findings, snapshot, rules, roles);
            CheckFontSizeConsistency(findings, snapshot, rules, roles);
            return findings;
        }

        public IReadOnlyList<AnnotationFinding> ScanSpelling(LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            var findings = new List<AnnotationFinding>();
            var paragraphs = Scannable(snapshot).ToArray();
            var roles = _roleDetector.Detect(snapshot);
            var lexicon = new VietnameseLexiconSpellChecker(rules.Lexicon);
            foreach (var paragraph in paragraphs)
            {
                AddMatches(findings, paragraph, Rx(@"[ \t]{2,}"), "LOCAL-TYPO-SPACE", "Có khoảng trắng thừa.", "Chỉ để một khoảng trắng.", rules);
                AddMatches(findings, paragraph, Rx(@"[ \t]+(?=[,.;:!?])"), "LOCAL-TYPO-PUNCT", "Có khoảng trắng trước dấu câu.", "Xóa khoảng trắng trước dấu câu.", rules);
                CheckHiddenCharacters(findings, paragraph, rules);
                CheckDictionary(findings, paragraph, rules);
                CheckLexicon(findings, paragraph, rules, lexicon);
                CheckTelex(findings, paragraph, rules);
                CheckBareShortDates(findings, paragraph, rules);
                CheckSentenceCapitalization(findings, paragraph, rules, roles);
                CheckPersonNames(findings, paragraph, rules);
                CheckArticleClauseCapitalization(findings, paragraph, rules);
            }
            CheckConfiguredCapitalizations(findings, paragraphs, rules, roles);
            CheckAdministrativeNumerals(findings, paragraphs, rules);
            return findings;
        }

        private static void CheckPageSetup(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            var party = IsParty(snapshot);
            foreach (var section in snapshot.Sections)
            {
                var shortSide = Math.Min(section.PageWidthPoints, section.PageHeightPoints);
                var longSide = Math.Max(section.PageWidthPoints, section.PageHeightPoints);
                var isA4 = Near(shortSide, Math.Min(rules.A4WidthMm, rules.A4HeightMm) * PointsPerMillimeter, 2d) &&
                    Near(longSide, Math.Max(rules.A4WidthMm, rules.A4HeightMm) * PointsPerMillimeter, 2d);
                var isA5 = party && Near(shortSide, 148d * PointsPerMillimeter, 2d) &&
                    Near(longSide, 210d * PointsPerMillimeter, 2d);
                if (!isA4 && !isA5)
                    findings.Add(Section("ND30-PL1-M1-K1", section.Index, "Khổ giấy không phải A4.", "Dùng khổ A4 210 × 297 mm.", rules));
                var marginsValid = party
                    ? Near(section.TopMarginPoints, 20d * PointsPerMillimeter, 1.5d) &&
                        Near(section.BottomMarginPoints, 20d * PointsPerMillimeter, 1.5d) &&
                        Near(section.LeftMarginPoints, 30d * PointsPerMillimeter, 1.5d) &&
                        Near(section.RightMarginPoints, 15d * PointsPerMillimeter, 1.5d)
                    : Between(section.TopMarginPoints, rules.TopMinMm, rules.TopMaxMm) &&
                        Between(section.BottomMarginPoints, rules.BottomMinMm, rules.BottomMaxMm) &&
                        Between(section.LeftMarginPoints, rules.LeftMinMm, rules.LeftMaxMm) &&
                        Between(section.RightMarginPoints, rules.RightMinMm, rules.RightMaxMm);
                if (!marginsValid)
                    findings.Add(Section("ND30-PL1-M1-K3", section.Index, "Lề trang nằm ngoài khoảng quy định.",
                        party ? "Lề trên/dưới 20 mm, trái 30 mm, phải 15 mm." : "Lề trên/dưới 20–25 mm, trái 30–35 mm, phải 15–20 mm.", rules));
                if (!section.HasPageNumbers)
                    findings.Add(Section("ND30-PL1-M1-K7", section.Index, "Section chưa có số trang.", "Đánh số trang bằng chữ số Ả Rập theo thể thức.", rules));
            }
        }

        private static void CheckBodyTypography(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            foreach (var paragraph in Scannable(snapshot))
            {
                if (!IsBody(paragraph, roles)) continue;
                if (!string.IsNullOrWhiteSpace(paragraph.FontName) && !Eq(paragraph.FontName, rules.BodyFontName))
                    findings.Add(Paragraph("ND30-PL1-M1-K4-FONT", paragraph, "Phông chữ không đúng quy định.", "Dùng " + rules.BodyFontName + ".", rules));
                if (paragraph.FontColor.HasValue && paragraph.FontColor.Value != 0 && paragraph.FontColor.Value != -16777216)
                    findings.Add(Paragraph("ND30-PL1-M1-K4-COLOR", paragraph, "Nội dung không dùng màu đen.", "Dùng màu đen tự động.", rules));
                if (paragraph.Alignment.HasValue && paragraph.Alignment.Value != rules.BodyAlignment)
                    findings.Add(Paragraph("ND30-PL1-M2-K6E-ALIGN", paragraph, "Đoạn nội dung chưa căn đều hai lề.", "Căn đều hai lề.", rules));
                var minimumSize = party ? 14d : rules.BodyFontMinPoints;
                var maximumSize = party ? 15d : rules.BodyFontMaxPoints;
                if (paragraph.FontSizePoints.HasValue &&
                    (paragraph.FontSizePoints.Value < minimumSize - .1d || paragraph.FontSizePoints.Value > maximumSize + .1d))
                    findings.Add(Paragraph("ND30-PL1-MV-CT1", paragraph, "Cỡ chữ nội dung không đúng chế độ đã chọn.",
                        "Dùng cỡ " + minimumSize.ToString("0.#", CultureInfo.InvariantCulture) + "–" + maximumSize.ToString("0.#", CultureInfo.InvariantCulture) + ".", rules));
                if (paragraph.FirstLineIndentPoints.HasValue && !Between(paragraph.FirstLineIndentPoints.Value,
                    rules.BodyFirstLineIndentMinMm, rules.BodyFirstLineIndentMaxMm))
                    findings.Add(Paragraph("ND30-PL1-M2-K6E-INDENT", paragraph, "Thụt đầu dòng không đúng.", "Thụt đầu dòng 1–1,27 cm.", rules));
                if (paragraph.SpaceAfterPoints.HasValue && paragraph.SpaceAfterPoints.Value < rules.BodySpaceAfterMinPoints - TolerancePoints)
                    findings.Add(Paragraph("ND30-PL1-M2-K6E-SPACEAFTER", paragraph, "Khoảng cách sau đoạn quá nhỏ.", "Đặt tối thiểu 6 pt.", rules));
                if (paragraph.LineSpacingPoints.HasValue && paragraph.LineSpacingRule.HasValue)
                {
                    var invalid = party
                        ? paragraph.LineSpacingRule.Value != 4 || paragraph.LineSpacingPoints.Value < 18d - .1d || paragraph.LineSpacingPoints.Value > 22d + .1d
                        : !IsNd30LineSpacingValid(paragraph);
                    if (invalid)
                        findings.Add(Paragraph("ND30-PL1-M2-K6E-LINESPACING", paragraph, "Giãn dòng ngoài khoảng cho phép.",
                            party ? "Dùng Exactly từ 18–22 pt." : "Dùng dòng đơn đến 1,5 dòng.", rules));
                }
            }
            var endBoundary = FirstRole(roles,
                "recipientLabel", "recipientList", "recipientSalutation", "recipientSalutationList",
                "signerAuthority", "signerRole", "signerName", "appendixLabel", "appendixTitle");

            if (!endBoundary.HasValue)
            {
                var recipientPara = Scannable(snapshot)
                    .Where(p => Rx(@"^Nơi\s+nhận\s*:", true).IsMatch(p.Text) ||
                                Rx(@"^(?:TM\.|KT\.|TL\.|TUQ\.)\s*", true).IsMatch(p.Text))
                    .OrderBy(p => p.Index)
                    .FirstOrDefault();
                if (recipientPara != null)
                    endBoundary = recipientPara.Index;
            }

            var headerRoles = new HashSet<string>(StringComparer.Ordinal)
            {
                "nationalTitle", "nationalMotto", "superiorOrganName", "organName",
                "partyTitle", "codeNumber", "placeAndIssuedDate", "typeName",
                "subject", "subjectContinuation", "legalBasis"
            };

            var contentParagraphs = Scannable(snapshot)
                .Where(p => !p.IsInTable &&
                            string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                            !string.IsNullOrWhiteSpace(p.Text) &&
                            (!endBoundary.HasValue || p.Index < endBoundary.Value) &&
                            (!roles.TryGetValue(p.Index, out var role) || !headerRoles.Contains(role)))
                .OrderBy(p => p.Index)
                .ToArray();

            var last = contentParagraphs.LastOrDefault();
            if (last != null)
            {
                var text = last.Text.Trim();
                if (text.Length > 0 &&
                    !text.EndsWith(":", StringComparison.Ordinal) &&
                    (!text.EndsWith(".", StringComparison.Ordinal) ||
                     text.EndsWith("./.", StringComparison.Ordinal) ||
                     text.EndsWith(". / .", StringComparison.Ordinal)))
                {
                    var printable = TrimParagraphTerminator(last.Text).TrimEnd();
                    var offset = Math.Max(0, printable.Length - Math.Min(3, printable.Length));
                    var length = Math.Max(1, printable.Length - offset);
                    findings.Add(Span("ND30-PL1-M2-K6E-DOTSLASH", last, offset, length,
                        "Kết thúc nội dung không đúng.", "Văn bản hành chính kết thúc bằng dấu chấm.", rules));
                }
            }
        }

        private static void CheckComponents(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            foreach (var paragraph in Scannable(snapshot))
            {
                string role;
                if (!roles.TryGetValue(paragraph.Index, out role)) continue;
                if (party && role == "partyTitle")
                    CheckStyle(findings, "ND30-PL1-M2-K1-QH", paragraph, rules, 15, 15, true, false, 1, "Tiêu đề Đảng");
                else if (!party && role == "nationalTitle")
                    CheckStyle(findings, "ND30-PL1-M2-K1-QH", paragraph, rules, 12, 13, true, false, 1, "Quốc hiệu");
                else if (!party && role == "nationalMotto")
                {
                    CheckStyle(findings, "ND30-PL1-M2-K1-TN", paragraph, rules, 13, 14, true, false, 1, "Tiêu ngữ");
                    var expected = "Độc lập - Tự do - Hạnh phúc";
                    if (!string.Equals(Collapse(paragraph.Text), expected, StringComparison.Ordinal))
                        findings.Add(Paragraph("ND30-PL1-M2-K1-TN-SEP", paragraph, "Tiêu ngữ sai chữ hoa hoặc dấu nối.", "Trình bày: " + expected + ".", rules));
                    if (paragraph.SpaceBeforePoints.HasValue && paragraph.SpaceBeforePoints.Value > TolerancePoints)
                        findings.Add(Paragraph("ND30-PL1-M2-K1-C", paragraph, "Quốc hiệu và Tiêu ngữ cách nhau quá một dòng đơn.", "Không thêm khoảng cách trước Tiêu ngữ.", rules));
                }
                else if (role == "superiorOrganName")
                    CheckStyle(findings, "ND30-PL1-M2-K2-SUP", paragraph, rules, party ? 14 : 12, party ? 14 : 13, false, false, 1, "Tên cơ quan chủ quản");
                else if (role == "organName")
                    CheckStyle(findings, "ND30-PL1-M2-K2-ORG", paragraph, rules, party ? 14 : 12, party ? 14 : 13, true, false, 1, "Tên cơ quan ban hành");
            }
        }

        private static void CheckRequiredLineShapes(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            var motto = FindComponent(snapshot, roles, "nationalMotto",
                text => text.IndexOf("Độc lập", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    text.IndexOf("Hạnh phúc", StringComparison.OrdinalIgnoreCase) >= 0);
            if (!party && motto != null)
                AddRequiredLineFinding(findings, snapshot, motto, rules.MottoLineMinRatio,
                    rules.MottoLineMaxRatio, "ND30-PL1-M2-K1-TN-LINE", "Tiêu ngữ",
                    "dài bằng dòng Tiêu ngữ", rules);

            foreach (var organ in WithRole(snapshot, roles, "organName"))
                if (!party)
                    AddRequiredLineFinding(findings, snapshot, organ, rules.OrganLineMinRatio,
                        rules.OrganLineMaxRatio, "ND30-PL1-M2-K2-ORG-LINE", "tên cơ quan ban hành",
                        "dài từ 1/3 đến 1/2 dòng tên cơ quan", rules);

            foreach (var subject in WithRole(snapshot, roles, "subject"))
                if (!party)
                {
                    var lineAnchor = LastSubjectParagraph(snapshot, roles, subject);
                    AddRequiredLineFinding(findings, snapshot, lineAnchor, rules.SubjectLineMinRatio,
                        rules.SubjectLineMaxRatio, "ND30-PL1-M2-K5A-SUBJ-LINE", "trích yếu",
                        "dài từ 1/3 đến 1/2 dòng trích yếu", rules);
                }

            var partyTitle = FindComponent(snapshot, roles, "partyTitle",
                text => string.Equals(Collapse(text), "ĐẢNG CỘNG SẢN VIỆT NAM", StringComparison.OrdinalIgnoreCase));
            if (partyTitle != null)
                AddRequiredLineFinding(findings, snapshot, partyTitle, rules.PartyTitleLineMinRatio,
                    rules.PartyTitleLineMaxRatio, "HD05-M1-TITLE-LINE", "tiêu đề Đảng",
                    "dài bằng tiêu đề ĐẢNG CỘNG SẢN VIỆT NAM", rules);
        }

        private static void CheckHeaderFontSizeTier(ICollection<AnnotationFinding> findings,
            LocalScanSnapshot snapshot, LocalRulePack rules, IDictionary<int, string> roles)
        {
            if (IsParty(snapshot)) return;
            var tier = Nd30HeaderFontSizeTierResolver.Resolve(snapshot, roles);
            CheckTierRole(findings, snapshot, rules, roles, "nationalTitle", tier.NationalTitle,
                "Quốc hiệu", tier);
            CheckTierRole(findings, snapshot, rules, roles, "nationalMotto", tier.NationalMotto,
                "Tiêu ngữ", tier);
            CheckTierRole(findings, snapshot, rules, roles, "placeAndIssuedDate", tier.PlaceAndIssuedDate,
                "Địa danh và ngày tháng", tier);
        }

        private static void CheckTierRole(ICollection<AnnotationFinding> findings,
            LocalScanSnapshot snapshot, LocalRulePack rules, IDictionary<int, string> roles,
            string role, double expectedSize, string label, Nd30HeaderFontSizeTier tier)
        {
            foreach (var paragraph in WithRole(snapshot, roles, role))
            {
                if (!paragraph.FontSizePoints.HasValue ||
                    Math.Abs(paragraph.FontSizePoints.Value - expectedSize) <= .1d)
                    continue;
                findings.Add(Paragraph("ND30-PL1-MV-CT1", paragraph,
                    label + " chưa thống nhất bậc cỡ chữ: đang dùng cỡ " +
                    paragraph.FontSizePoints.Value.ToString("0.#", CultureInfo.InvariantCulture) + ".",
                    "Đặt Quốc hiệu cỡ " + tier.NationalTitle.ToString("0", CultureInfo.InvariantCulture) +
                    ", Tiêu ngữ cỡ " + tier.NationalMotto.ToString("0", CultureInfo.InvariantCulture) +
                    ", địa danh và ngày tháng cỡ " +
                    tier.PlaceAndIssuedDate.ToString("0", CultureInfo.InvariantCulture) + ".", rules));
            }
        }

        private static void AddRequiredLineFinding(ICollection<AnnotationFinding> findings,
            LocalScanSnapshot snapshot, LocalParagraphSnapshot paragraph, double minimumWidthRatio,
            double maximumWidthRatio, string ruleCode, string componentName, string expectedLength,
            LocalRulePack rules)
        {
            var status = EvaluateRequiredLine(snapshot, paragraph, minimumWidthRatio, maximumWidthRatio);
            if (status == RequiredLineStatus.Valid) return;

            string issue;
            switch (status)
            {
                case RequiredLineStatus.InvalidStyle:
                    issue = "Đã có Line Shape dưới " + componentName +
                        " nhưng đường kẻ không phải nét liền, đang có mũi tên hoặc đang bị ẩn.";
                    break;
                case RequiredLineStatus.InvalidPosition:
                    issue = "Đã có Line Shape gần " + componentName +
                        " nhưng chưa đặt đúng vị trí bên dưới.";
                    break;
                case RequiredLineStatus.InvalidGeometry:
                    issue = "Đã có Line Shape dưới " + componentName +
                        " nhưng độ dài, độ ngang hoặc vị trí cân giữa chưa đúng.";
                    break;
                default:
                    issue = "Thiếu đường kẻ Line Shape nét liền dưới " + componentName + ".";
                    break;
            }

            findings.Add(Paragraph(ruleCode, paragraph, issue,
                "Dùng Line Shape nét liền, không mũi tên, " + expectedLength +
                " và đặt cân đối bên dưới; underline hoặc viền đoạn không thay thế được.",
                rules, "Error"));
        }

        private static LocalParagraphSnapshot FindComponent(LocalScanSnapshot snapshot,
            IDictionary<int, string> roles, string role, Func<string, bool> fallback)
        {
            var byRole = WithRole(snapshot, roles, role).FirstOrDefault();
            return byRole ?? Scannable(snapshot).FirstOrDefault(item => fallback(item.Text));
        }

        private static RequiredLineStatus EvaluateRequiredLine(LocalScanSnapshot snapshot,
            LocalParagraphSnapshot paragraph,
            double minimumWidthRatio, double maximumWidthRatio)
        {
            var associated = new List<LocalLineShapeSnapshot>();
            foreach (var line in snapshot.LineShapes)
            {
                if (line.ShapeType != 9) continue;
                if (!string.Equals(line.AnchorStoryType, paragraph.StoryType, StringComparison.Ordinal)) continue;
                if (line.AnchorSectionIndex != paragraph.SectionIndex) continue;
                if (paragraph.PageNumber > 0 && line.AnchorPageNumber > 0 && line.AnchorPageNumber != paragraph.PageNumber) continue;
                if (!IsPotentiallyAssociated(line, paragraph)) continue;
                associated.Add(line);
            }

            if (associated.Count == 0) return RequiredLineStatus.Missing;

            var styled = associated.Where(HasRequiredLineStyle).ToArray();
            if (styled.Length == 0) return RequiredLineStatus.InvalidStyle;

            var positioned = styled.Where(line => IsBelow(line, paragraph)).ToArray();
            if (positioned.Length == 0) return RequiredLineStatus.InvalidPosition;

            return positioned.Any(line => IsHorizontal(line) &&
                HasExpectedWidthAndCenter(line, paragraph, minimumWidthRatio, maximumWidthRatio))
                ? RequiredLineStatus.Valid
                : RequiredLineStatus.InvalidGeometry;
        }

        private static bool IsOwnedLineForParagraph(string? name, int paragraphIndex)
        {
            if (name == null || name.Length == 0) return false;
            var marker = "P" + paragraphIndex.ToString(CultureInfo.InvariantCulture);
            return (name.StartsWith("CHUANHOA2_", StringComparison.Ordinal) ||
                    name.StartsWith("CHUANHOA_", StringComparison.Ordinal) ||
                    name.StartsWith("CH_L_", StringComparison.Ordinal)) &&
                   (name.EndsWith(marker, StringComparison.Ordinal) || name.IndexOf(marker + "_", StringComparison.Ordinal) >= 0);
        }

        private static bool HasRequiredLineStyle(LocalLineShapeSnapshot line)
        {
            if (!line.LineVisible) return false;
            if (line.DashStyle.HasValue && line.DashStyle.Value != 1) return false;
            if (line.BeginArrowheadStyle.HasValue && line.BeginArrowheadStyle.Value != 1) return false;
            return !line.EndArrowheadStyle.HasValue || line.EndArrowheadStyle.Value == 1;
        }

        private static bool IsPotentiallyAssociated(LocalLineShapeSnapshot line,
            LocalParagraphSnapshot paragraph)
        {
            if (IsOwnedLineForParagraph(line.Name, paragraph.Index))
                return true;
            if (line.AnchorParagraphIndex.HasValue && line.AnchorParagraphIndex.Value == paragraph.Index)
                return true;
            if (line.AnchorParagraphIndex.HasValue &&
                line.AnchorParagraphIndex.Value > paragraph.Index &&
                line.AnchorParagraphIndex.Value <= paragraph.Index + 2 &&
                (line.RelativeVerticalPosition == 2 || line.RelativeVerticalPosition == 3) &&
                line.TopPoints >= -6d && line.TopPoints <= 100d &&
                line.PageTopPoints.HasValue && paragraph.PageTopPoints.HasValue &&
                line.PageTopPoints.Value < paragraph.PageTopPoints.Value - 6d)
            {
                // Some Word builds report a near-zero page position for a line
                // anchored to the blank paragraph immediately after its caption.
                return true;
            }
            if (!line.PageTopPoints.HasValue || !paragraph.PageTopPoints.HasValue) return false;
            if (!line.PageLeftPoints.HasValue || !paragraph.PageLeftPoints.HasValue ||
                !paragraph.TextWidthPoints.HasValue || paragraph.TextWidthPoints.Value <= 0d)
                return IsAnchoredNear(line, paragraph);

            // A floating Word shape may stay visually underneath a component while Word
            // anchors it to an unrelated paragraph on the same page. Match the rendered
            // position first; use the anchor only when page coordinates are unavailable.
            var verticalDistance = LineVerticalCenter(line) - paragraph.PageTopPoints.Value;
            if (verticalDistance < -6d || verticalDistance > 110d) return false;

            var lineCenter = line.PageLeftPoints.Value + Math.Abs(line.WidthPoints) / 2d;
            var textCenter = paragraph.PageLeftPoints.Value + paragraph.TextWidthPoints.Value / 2d;
            return Math.Abs(lineCenter - textCenter) <= Math.Max(24d, paragraph.TextWidthPoints.Value * .35d);
        }

        private static bool IsAnchoredNear(LocalLineShapeSnapshot line, LocalParagraphSnapshot paragraph)
        {
            if (line.AnchorParagraphIndex.HasValue)
                return line.AnchorParagraphIndex.Value >= paragraph.Index - 1 &&
                    line.AnchorParagraphIndex.Value <= paragraph.Index + 2;
            return Math.Abs(line.AnchorAbsoluteStart - paragraph.AbsoluteStart) <= 500;
        }

        private static bool IsBelow(LocalLineShapeSnapshot line, LocalParagraphSnapshot paragraph)
        {
            if (IsOwnedLineForParagraph(line.Name, paragraph.Index))
                return true;
            if (line.AnchorParagraphIndex.HasValue &&
                line.AnchorParagraphIndex.Value > paragraph.Index &&
                line.AnchorParagraphIndex.Value <= paragraph.Index + 2 &&
                (line.RelativeVerticalPosition == 2 || line.RelativeVerticalPosition == 3) &&
                line.TopPoints >= -6d && line.TopPoints <= 100d)
            {
                // NĐ30 legacy templates anchor the separator to the blank paragraph
                // immediately following the caption. Some Word builds return zero for
                // that blank anchor's page position, but the paragraph-relative Top is
                // still stable and represents the rendered line correctly.
                return true;
            }

            if (line.PageTopPoints.HasValue && paragraph.PageTopPoints.HasValue)
            {
                var distance = LineVerticalCenter(line) - paragraph.PageTopPoints.Value;
                var fontSize = paragraph.FontSizePoints.GetValueOrDefault(13d);
                return distance >= 2d && distance <= Math.Max(48d, fontSize * 5.5d);
            }
            return IsAnchoredNear(line, paragraph) && line.TopPoints >= -6d && line.TopPoints <= 100d;
        }

        private static bool IsHorizontal(LocalLineShapeSnapshot line)
        {
            var length = LineLength(line);
            if (length <= 0d) return false;
            return Math.Abs(line.HeightPoints) <= Math.Max(3d, length * .04d);
        }

        private static double LineLength(LocalLineShapeSnapshot line)
        {
            var width = Math.Abs(line.WidthPoints);
            var height = Math.Abs(line.HeightPoints);
            return Math.Sqrt(width * width + height * height);
        }

        private static double LineVerticalCenter(LocalLineShapeSnapshot line)
        {
            return line.PageTopPoints.GetValueOrDefault() + Math.Abs(line.HeightPoints) / 2d;
        }

        private static bool HasExpectedWidthAndCenter(LocalLineShapeSnapshot line, LocalParagraphSnapshot paragraph,
            double minimumWidthRatio, double maximumWidthRatio)
        {
            if (IsOwnedLineForParagraph(line.Name, paragraph.Index))
                return true;
            if (paragraph.TextWidthPoints.HasValue && paragraph.TextWidthPoints.Value > 0)
            {
                var lineLength = LineLength(line);
                var ratio = lineLength / paragraph.TextWidthPoints.Value;
                if (ratio < minimumWidthRatio - .04d || ratio > maximumWidthRatio + .04d) return false;
                // Legacy DOC templates commonly store a visually centered line
                // relative to its table column (Word value 2). Word exposes that
                // coordinate in a different local origin after repagination, so the
                // absolute-center comparison is reliable only for page-relative lines.
                if (line.RelativeHorizontalPosition == 1 &&
                    line.PageLeftPoints.HasValue && paragraph.PageLeftPoints.HasValue)
                {
                    var lineCenter = line.PageLeftPoints.Value + Math.Abs(line.WidthPoints) / 2d;
                    var textCenter = paragraph.PageLeftPoints.Value + paragraph.TextWidthPoints.Value / 2d;
                    // A small offset is visually balanced in Word and is common after
                    // DOC compatibility-mode pagination. Reject only a material shift;
                    // 1-Click still re-centres every canonical line when explicitly run.
                    if (Math.Abs(lineCenter - textCenter) >
                        Math.Max(8d, paragraph.TextWidthPoints.Value * .04d))
                        return false;
                }
            }
            return true;
        }

        private enum RequiredLineStatus
        {
            Missing,
            InvalidStyle,
            InvalidPosition,
            InvalidGeometry,
            Valid
        }

        private static void CheckCodeNumber(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            foreach (var paragraph in WithRole(snapshot, roles, "codeNumber"))
            {
                var match = CodeNumber.Match(paragraph.Text);
                if (!match.Success) continue;
                if (!party && match.Groups["colon"].Length == 0)
                    findings.Add(Paragraph("ND30-PL1-M2-K3-PREFIX", paragraph, "Thiếu dấu hai chấm sau từ Số.", "Viết Số: ...", rules));
                if (party && match.Groups["colon"].Length > 0)
                    findings.Add(Paragraph("ND30-PL1-M2-K3-PREFIX", paragraph, "Văn bản Đảng không dùng dấu hai chấm sau từ Số.", "Viết Số 01-QĐ/TW.", rules));
                int value;
                var number = match.Groups["number"];
                if (int.TryParse(number.Value, out value) && value < 10 && number.Length < 2)
                    findings.Add(Span("ND30-PL1-M2-K3-PAD", paragraph, number.Index, number.Length, "Số nhỏ hơn 10 chưa đệm 0.", "Viết đủ hai chữ số.", rules));
                var separator = match.Groups["separator"];
                var notation = match.Groups["notation"];
                var separatorValid = party
                    ? separator.Value == "-" && Rx(@"^[\p{L}\d]+/[\p{L}\d-]+$").IsMatch(notation.Value)
                    : separator.Value == "/" && Rx(@"^[\p{L}\d-]+$").IsMatch(notation.Value);
                if (!separatorValid)
                    findings.Add(Paragraph("ND30-PL1-M2-K3-SEP", paragraph, "Dấu phân cách số, ký hiệu không đúng.",
                        party ? "Dùng Số 01-QĐ/TW." : "Dùng dấu / và dấu - giữa các nhóm ký hiệu.", rules));
                if (notation.Value.Any(char.IsWhiteSpace))
                    findings.Add(Span("ND30-PL1-M2-K3-SPACE", paragraph, notation.Index, notation.Length, "Ký hiệu có khoảng trắng.", "Không để khoảng trắng trong ký hiệu.", rules));
                if (notation.Value.Any(char.IsLower))
                    findings.Add(Span("ND30-PL1-M2-K3-CASE", paragraph, notation.Index, notation.Length, "Ký hiệu có chữ thường.", "Viết hoa ký hiệu, trừ viết tắt được quy định riêng.", rules));
                var type = WithRole(snapshot, roles, "typeName").FirstOrDefault();
                if (type != null)
                {
                    var entry = rules.DocumentTypeAbbreviations.FirstOrDefault(x => Eq(Collapse(type.Text).Trim('.', ':'), x.TypeName));
                    var actual = party ? notation.Value.Split('/')[0] : notation.Value.Split('-')[0];
                    if (entry != null && !Eq(actual, entry.Abbreviation))
                        findings.Add(Span("ND30-PL1-M2-K3-ABBR", paragraph, notation.Index, Math.Min(actual.Length, notation.Length),
                            "Ký hiệu tên loại không khớp tên văn bản.", "Dùng " + entry.Abbreviation + " cho " + entry.TypeName + ".", rules));
                }
            }
        }

        private static void CheckPlaceDate(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            foreach (var paragraph in WithRole(snapshot, roles, "placeAndIssuedDate"))
            {
                CheckStyle(findings, "ND30-PL1-M2-K4-STYLE", paragraph, rules, party ? 14 : 13, 14, false, true, 1, "Địa danh và ngày tháng");
                var match = PlaceDate.Match(paragraph.Text);
                if (!match.Success) continue;
                if (match.Groups["comma"].Length == 0)
                    findings.Add(Span("ND30-PL1-M2-K4-COMMA", paragraph, match.Groups["place"].Index, match.Groups["place"].Length,
                        "Thiếu dấu phẩy sau địa danh.", "Thêm dấu phẩy sau địa danh.", rules));
                var place = match.Groups["place"];
                if (!IsTitleCase(place.Value.Trim()))
                    findings.Add(Span("ND30-PL1-M2-K4-CASE", paragraph, place.Index, place.Length, "Địa danh sai chữ hoa.", "Viết hoa tên riêng địa danh.", rules));
                foreach (var groupName in new[] { "day", "month" })
                {
                    var group = match.Groups[groupName];
                    int value;
                    if (group.Length == 1 && int.TryParse(group.Value, out value) && value < 10 && (groupName == "day" || value <= 2))
                        findings.Add(Span("ND30-PL1-M2-K4-PAD", paragraph, group.Index, group.Length, "Ngày hoặc tháng thiếu số 0 phía trước.", "Đệm 0 cho ngày dưới 10 và tháng 1, 2.", rules));
                }
            }
        }

        private static void CheckTypeAndSubject(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            foreach (var paragraph in WithRole(snapshot, roles, "typeName"))
                CheckStyle(findings, "ND30-PL1-M2-K5A-TYPE", paragraph, rules, party ? 15 : 13, party ? 16 : 14, true, false, 1, "Tên loại văn bản");
            foreach (var paragraph in WithRole(snapshot, roles, "subject"))
                CheckStyle(findings, "ND30-PL1-M2-K5A-SUBJ", paragraph, rules, party ? 14 : 13, party ? 15 : 14, true, false, 1, "Trích yếu");
            foreach (var paragraph in WithRole(snapshot, roles, "subjectContinuation"))
                CheckStyle(findings, "ND30-PL1-M2-K5A-SUBJ", paragraph, rules, party ? 14 : 13, party ? 15 : 14, true, false, 1, "Trích yếu");
            foreach (var paragraph in WithRole(snapshot, roles, "officialLetterSubject"))
            {
                CheckStyle(findings, "ND30-PL1-M2-K5B-STYLE", paragraph, rules, 12, party ? 12 : 13, false, party, party ? 1 : 1, "Trích yếu công văn");
                var prefix = party ? @"^\s*(?:về việc|V/v)\s+" : @"^\s*(?:V/v|Về việc)\s+";
                if (!Rx(prefix, !party).IsMatch(paragraph.Text))
                    findings.Add(Paragraph("ND30-PL1-M2-K5B-SPACE", paragraph, "Trích yếu công văn sai tiền tố hoặc khoảng cách.",
                        party ? "Trình bày trích yếu công văn bằng chữ thường, cỡ 12, nghiêng." : "Bắt đầu bằng V/v hoặc Về việc và một khoảng trắng.", rules));
            }
        }

        private static void CheckLegalBasisAndCitations(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            var bases = WithRole(snapshot, roles, "legalBasis").ToArray();
            for (var index = 0; index < bases.Length; index++)
            {
                var paragraph = bases[index];
                CheckStyle(findings, "ND30-PL1-M2-K6A-STYLE", paragraph, rules, party ? 14 : 13, party ? 15 : 14, false, true, 3, "Căn cứ pháp lý");
                var printable = TrimParagraphTerminator(paragraph.Text).TrimEnd();
                var isLast = index == bases.Length - 1;
                var expected = isLast ? "." : ";";
                if (printable.Length > 0 && !printable.EndsWith(expected, StringComparison.Ordinal))
                {
                    var offset = Math.Max(0, printable.Length - 1);
                    var length = Math.Max(1, printable.Length - offset);
                    findings.Add(Span("ND30-PL1-M2-K6A-PUNCT", paragraph, offset, length,
                        "Căn cứ pháp lý sai dấu kết thúc.", isLast ? "Kết thúc căn cứ cuối bằng dấu chấm." : "Kết thúc từng căn cứ trước bằng dấu chấm phẩy.", rules));
                }
                if (party && !Rx(@"^\s*[-–—]\s*").IsMatch(paragraph.Text))
                    findings.Add(Paragraph("ND30-PL1-M2-K6A-STYLE", paragraph, "Căn cứ của văn bản Đảng thiếu gạch ngang đầu dòng.", "Thêm một dấu gạch ngang trước mỗi căn cứ.", rules));
            }
            var scan = Scannable(snapshot).Where(p => p.StoryType == "wdMainTextStory").OrderBy(p => p.Index).ToArray();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var paragraph in scan)
            {
                string paragraphRole;
                var isLegalBasis = roles.TryGetValue(paragraph.Index, out paragraphRole) &&
                    string.Equals(paragraphRole, "legalBasis", StringComparison.Ordinal);
                var citations = Citation.Matches(paragraph.Text).Cast<Match>().ToArray();
                for (var citationIndex = 0; citationIndex < citations.Length; citationIndex++)
                {
                    var match = citations[citationIndex];
                    var citedType = match.Groups["type"].Value;
                    var isLawOrOrdinance = Rx(@"^(luật|pháp\s+lệnh)$", true).IsMatch(citedType);
                    var so = match.Groups["so"];
                    if (!isLawOrOrdinance && so.Length == 0)
                        findings.Add(Span("ND30-PL1-M2-K6B-SO", paragraph, match.Index, match.Length, "Viện dẫn thiếu từ số.", "Thêm từ số sau tên loại văn bản.", rules));
                    var code = match.Groups["code"].Value.TrimEnd('.', ',', ';', ':');
                    if (!isLawOrOrdinance && seen.Add(code))
                    {
                        var segmentEnd = citationIndex + 1 < citations.Length
                            ? citations[citationIndex + 1].Index
                            : paragraph.Text.Length;
                        var segment = paragraph.Text.Substring(match.Index, segmentEnd - match.Index);
                        var hasFullDate = Rx(@"ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}", true).IsMatch(segment);
                        var hasAbbreviatedDate = AbbreviatedDate.IsMatch(segment);
                        var hasShortDate = LegalBasisShortDate.IsMatch(segment);
                        var hasIssuingAgency = segment.IndexOf("của ", StringComparison.OrdinalIgnoreCase) >= 0;
                        if ((!hasFullDate && !hasAbbreviatedDate && !hasShortDate) || !hasIssuingAgency)
                            findings.Add(Span("ND30-PL1-M2-K6B-CITE", paragraph, match.Index, match.Length, "Lần viện dẫn đầu thiếu ngày ban hành hoặc cơ quan ban hành.", "Ghi đủ tên loại, số, ký hiệu, ngày và cơ quan ban hành.", rules));
                    }
                }
            }
        }

        private static void CheckStructure(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var paragraphs = Scannable(snapshot).Where(p => p.StoryType == "wdMainTextStory").OrderBy(p => p.Index).ToArray();
            var articles = paragraphs.Select(p => new { Paragraph = p, Match = Article.Match(p.Text) }).Where(x => x.Match.Success).ToArray();
            int? previousArticle = null;
            foreach (var item in articles)
            {
                int number;
                int.TryParse(item.Match.Groups["number"].Value, out number);
                var problems = new List<string>();
                if (previousArticle.HasValue && number != previousArticle.Value + 1) problems.Add("số điều không liên tục");
                if (item.Paragraph.FirstLineIndentPoints.HasValue && !Between(item.Paragraph.FirstLineIndentPoints.Value, 10, 12.7)) problems.Add("thụt đầu dòng sai");
                if (item.Paragraph.Bold.HasValue && !item.Paragraph.Bold.Value) problems.Add("chưa in đậm");
                if (problems.Count > 0)
                    findings.Add(Paragraph("ND30-PL1-M2-K6D-ARTICLE", item.Paragraph, "Điều " + number + " sai thể thức: " + string.Join(", ", problems) + ".", "Sửa số thứ tự, thụt dòng và kiểu chữ của điều.", rules));
                if (string.IsNullOrWhiteSpace(item.Match.Groups["title"].Value))
                    findings.Add(Paragraph("ND30-PL1-M2-K6D-TITLE", item.Paragraph, "Điều " + number + " thiếu tiêu đề.", "Bổ sung tiêu đề sau số điều.", rules));
                previousArticle = number;
            }
            if (articles.Length > 0)
            {
                foreach (var paragraph in paragraphs.Where(p => Clause.IsMatch(p.Text)))
                    if (paragraph.Italic.HasValue && paragraph.Italic.Value)
                        findings.Add(Paragraph("ND30-PL1-M2-K6D-CLAUSE", paragraph, "Khoản đang in nghiêng.", "Khoản dùng kiểu chữ đứng.", rules));
            }
            var pointAlphabet = new[] { "a", "b", "c", "d", "đ", "e", "g", "h", "i", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "x", "y" };
            int previousPoint = -1;
            foreach (var paragraph in paragraphs)
            {
                var match = Point.Match(paragraph.Text);
                if (!match.Success)
                {
                    if (Article.IsMatch(paragraph.Text) || Clause.IsMatch(paragraph.Text)) previousPoint = -1;
                    continue;
                }
                if ((paragraph.Italic.HasValue && paragraph.Italic.Value) || (paragraph.Bold.HasValue && paragraph.Bold.Value))
                    findings.Add(Paragraph("ND30-PL1-M2-K6D-POINT", paragraph, "Điểm sai kiểu chữ.", "Điểm dùng chữ đứng, không đậm.", rules));
                var current = Array.IndexOf(pointAlphabet, match.Groups["letter"].Value.ToLower(Vietnamese));
                if (previousPoint >= 0 && current != previousPoint + 1)
                    findings.Add(Span("ND30-PL1-M2-K6D-ALPHABET", paragraph, match.Groups["letter"].Index, match.Groups["letter"].Length,
                        "Thứ tự điểm không liên tục theo bảng chữ cái tiếng Việt.", "Sắp xếp a, b, c, d, đ, e, g...", rules));
                previousPoint = current;
            }
        }

        private static void CheckSignerAndRecipients(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var party = IsParty(snapshot);
            foreach (var paragraph in WithRole(snapshot, roles, "signerAuthority"))
            {
                CheckStyle(findings, "ND30-PL1-M2-K7D-STYLE", paragraph, rules, party ? 14 : 13, 14, true, false, 1, "Quyền hạn, chức vụ người ký");
                var match = Rx(party ? @"^\s*(?<abbr>t/m|k/t|t/l|q)\.?" : @"^\s*(?<abbr>tm|kt|tl|tuq|q)\.", true).Match(paragraph.Text);
                if (match.Success && match.Groups["abbr"].Value.Any(char.IsLower))
                    findings.Add(Span("ND30-PL1-M2-K7B-AUTH", paragraph, match.Groups["abbr"].Index, match.Groups["abbr"].Length,
                        "Chữ viết tắt thẩm quyền ký viết thường.", party ? "Viết hoa T/M, K/T, T/L hoặc Q." : "Viết hoa TM., KT., TL., TUQ. hoặc Q.", rules));
            }
            foreach (var paragraph in Scannable(snapshot).Where(p => Rx(@"^\s*Kính\s+(gửi|trình)", true).IsMatch(p.Text)))
            {
                CheckStyle(findings, "ND30-PL1-M2-K9A-LAYOUT", paragraph, rules, party ? 14 : 13, party ? 15 : 14,
                    false, party, paragraph.Alignment.GetValueOrDefault(3), "Kính gửi/Kính trình");
                var role = roles.ContainsKey(paragraph.Index) ? roles[paragraph.Index] : string.Empty;
                if (paragraph.Text.IndexOf(':') < 0)
                    findings.Add(Paragraph("ND30-PL1-M2-K9A-COLON", paragraph, "Thiếu dấu hai chấm sau Kính gửi.", "Viết Kính gửi: ...", rules));
                if (role == "recipientSalutation" && !WithRole(snapshot, roles, "recipientSalutationList").Any())
                    findings.Add(Paragraph("ND30-PL1-M2-K9A-LAYOUT", paragraph, "Kính gửi tách dòng nhưng không có danh sách bên dưới.", "Một nơi thì viết cùng dòng; nhiều nơi thì xuống dòng có gạch đầu dòng.", rules));
                if (role == "recipientSalutationInline" && !paragraph.Text.Trim().EndsWith(".", StringComparison.Ordinal))
                    findings.Add(Paragraph("ND30-PL1-M2-K9A-INLINE-END", paragraph, "Kính gửi viết cùng dòng chưa kết thúc bằng dấu chấm.", "Thêm dấu chấm cuối dòng.", rules));
            }
            var salutationItems = WithRole(snapshot, roles, "recipientSalutationList").OrderBy(p => p.Index).ToArray();
            for (var i = 0; i < salutationItems.Length; i++)
            {
                var text = salutationItems[i].Text.Trim();
                var expectedEnd = i == salutationItems.Length - 1 ? "." : ";";
                if (!text.StartsWith("-", StringComparison.Ordinal) || !text.EndsWith(expectedEnd, StringComparison.Ordinal))
                    findings.Add(Paragraph("ND30-PL1-M2-K9A-PUNCT", salutationItems[i], "Danh sách Kính gửi sai gạch đầu dòng hoặc dấu kết thúc.", "Dùng gạch đầu dòng; chấm phẩy giữa các mục và chấm ở mục cuối.", rules));
            }
            foreach (var paragraph in WithRole(snapshot, roles, "recipientLabel"))
            {
                CheckStyle(findings, "ND30-PL1-M2-K9B-LABEL", paragraph, rules, party ? 14 : 12, party ? 14 : 12,
                    !party, !party, 0, "Nơi nhận");
                if (party && (!paragraph.Underline.HasValue || paragraph.Underline.Value == 0))
                    findings.Add(Paragraph("ND30-PL1-M2-K9B-LABEL", paragraph, "Nơi nhận của văn bản Đảng chưa gạch chân.", "Gạch chân cụm từ Nơi nhận.", rules));
                if (!Rx(@"^\s*Nơi\s+nhận\s*:", true).IsMatch(paragraph.Text))
                    findings.Add(Paragraph("ND30-PL1-M2-K9B-LABEL", paragraph, "Nơi nhận thiếu dấu hai chấm.", "Viết Nơi nhận:", rules));
            }
            var recipientItems = WithRole(snapshot, roles, "recipientList").OrderBy(p => p.Index).ToArray();
            for (var i = 0; i < recipientItems.Length; i++)
            {
                CheckStyle(findings, "ND30-PL1-M2-K9B-LIST", recipientItems[i], rules, party ? 12 : 11,
                    party ? 12 : 11, false, false, 0, "Danh sách Nơi nhận");
                var text = recipientItems[i].Text.Trim();
                if (!text.StartsWith("-", StringComparison.Ordinal) || (i < recipientItems.Length - 1 && !text.EndsWith(";", StringComparison.Ordinal)))
                    findings.Add(Paragraph("ND30-PL1-M2-K9B-LIST", recipientItems[i], "Danh sách Nơi nhận sai gạch đầu dòng hoặc dấu câu.", "Mỗi dòng bắt đầu bằng gạch ngang và các dòng trước kết thúc bằng chấm phẩy.", rules));
            }
            if (recipientItems.Length > 0 && !Rx(@"^\s*-\s*Lưu\s*:\s*[^;,.]+[.;]?\s*$", true).IsMatch(recipientItems[recipientItems.Length - 1].Text))
                findings.Add(Paragraph("ND30-PL1-M2-K9B-LUU", recipientItems[recipientItems.Length - 1], "Dòng Lưu chưa đúng cấu trúc.", "Dòng cuối có dạng - Lưu: VT, ...", rules));
        }

        private static void CheckAppendices(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var labels = WithRole(snapshot, roles, "appendixLabel").OrderBy(p => p.Index).ToArray();
            if (labels.Length == 0) return;
            var first = labels[0].Index;
            if (!Scannable(snapshot).Any(p => p.Index < first && p.Text.IndexOf("phụ lục", StringComparison.OrdinalIgnoreCase) >= 0))
                findings.Add(Paragraph("ND30-PL1-M3-K1A-REF", labels[0], "Văn bản có phụ lục nhưng phần nội dung chưa nhắc tới.", "Bổ sung chỉ dẫn về phụ lục kèm theo.", rules));
            if (labels.Length > 1)
            {
                foreach (var label in labels)
                    if (!Rx(@"^\s*Phụ\s+lục\s+(?:[IVXLCDM]+|\d+)\b", true).IsMatch(label.Text))
                        findings.Add(Paragraph("ND30-PL1-M3-K1A-NUM", label, "Phụ lục chưa có số thứ tự.",
                            "Dùng số Ả Rập (Phụ lục 1, 2, 3...) hoặc số La Mã (Phụ lục I, II, III...).", rules));
            }
            foreach (var label in labels)
            {
                CheckStyle(findings, "ND30-PL1-M3-K1B", label, rules, 14, 14, true, false, 1, "Nhãn phụ lục");
                var nextLabelIndex = labels.Where(item => item.Index > label.Index)
                    .Select(item => item.Index).DefaultIfEmpty(int.MaxValue).First();
                var titleLines = WithRole(snapshot, roles, "appendixTitle")
                    .Where(item => item.SectionIndex == label.SectionIndex && item.Index > label.Index &&
                        item.Index < nextLabelIndex)
                    .OrderBy(item => item.Index)
                    .ToArray();
                foreach (var title in titleLines)
                {
                    CheckStyle(findings, "ND30-PL1-M3-K1B", title, rules, 13, 14, true, false, 1, "Tên phụ lục");
                    if (!IsUppercaseText(title.Text))
                        findings.Add(Paragraph("ND30-PL1-M3-K1B", title, "Tên phụ lục chưa viết hoa.",
                            "Viết hoa toàn bộ tên Phụ lục, cỡ chữ 13–14, in đậm và căn giữa.", rules));
                }
                var reference = WithRole(snapshot, roles, "appendixReference")
                    .FirstOrDefault(item => item.SectionIndex == label.SectionIndex &&
                        item.Index > label.Index && item.Index < nextLabelIndex);
                if (reference != null)
                    CheckStyle(findings, "ND30-PL1-M3-K1C", reference, rules, 13, 14, false, true, 1,
                        "Thông tin kèm theo phụ lục");
                else
                    findings.Add(Paragraph("ND30-PL1-M3-K1C", label,
                        "Phụ lục thiếu dòng thông tin kèm theo văn bản.",
                        "Thêm dòng căn giữa, cỡ chữ 13–14, in nghiêng theo dạng: (Kèm theo Văn bản số ... ngày ... tháng ... năm ... của ...).",
                        rules));

                var digitalInfo = Scannable(snapshot)
                    .Where(p => p.SectionIndex == label.SectionIndex && p.Index < label.Index &&
                        roles.TryGetValue(p.Index, out var role) && role == "appendixDigitalSignatureInfo")
                    .OrderByDescending(p => p.Index)
                    .Take(3)
                    .FirstOrDefault();
                if (digitalInfo != null)
                    CheckStyle(findings, "ND30-PL1-M3-K1C", digitalInfo, rules, 10, 10, false, false, 2,
                        "Thông tin ký số của Phụ lục điện tử");
                var section = snapshot.Sections.FirstOrDefault(s => s.Index == label.SectionIndex);
                if (section != null && !section.RestartPageNumbering)
                    findings.Add(Section("ND30-PL1-M3-K1D", section.Index,
                        "Phụ lục chưa đánh số trang riêng.",
                        "Tạo section riêng và đánh số trang lại từ đầu cho từng Phụ lục.", rules));
            }
        }

        private static bool IsUppercaseText(string value)
        {
            var letters = (value ?? string.Empty).Where(char.IsLetter).ToArray();
            return letters.Length == 0 || new string(letters) == new string(letters).ToUpper(Vietnamese);
        }

        private static void CheckFontSizeConsistency(ICollection<AnnotationFinding> findings, LocalScanSnapshot snapshot,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            var recognized = Scannable(snapshot).Where(p => roles.ContainsKey(p.Index) && p.FontSizePoints.HasValue).ToArray();
            foreach (var paragraph in recognized)
            {
                double min;
                double max;
                ExpectedSize(roles[paragraph.Index], rules, out min, out max);
                var size = paragraph.FontSizePoints.GetValueOrDefault();
                if (min > 0 && (size < min - .1 || size > max + .1))
                {
                    findings.Add(Paragraph("ND30-PL1-MV-CT1", paragraph, "Cỡ chữ không nhất quán với vai trò được nhận diện.", "Dùng cỡ chữ theo thành phần thể thức.", rules));
                }
            }
        }

        private static void CheckHiddenCharacters(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            foreach (var hidden in rules.HiddenCharacters)
            {
                var offset = 0;
                while ((offset = paragraph.Text.IndexOf(hidden, offset)) >= 0)
                {
                    findings.Add(Span("LOCAL-TYPO-HIDDEN", paragraph, offset, 1, "Có ký tự ẩn U+" + ((int)hidden).ToString("X4", CultureInfo.InvariantCulture) + ".", "Loại bỏ ký tự ẩn.", rules));
                    offset++;
                }
            }
        }

        private static void CheckDictionary(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            foreach (var correction in rules.Corrections)
            foreach (var occurrence in WholePhraseMatches(paragraph.Text, correction.Wrong))
            {
                var actual = paragraph.Text.Substring(occurrence.Item1, occurrence.Item2);
                findings.Add(Span("LOCAL-TYPO-DICT", paragraph, occurrence.Item1, occurrence.Item2, "Cụm từ có thể sai chính tả.", "Nên dùng “" + ApplyCase(actual, correction.Replacement) + "”.", rules));
            }
        }

        private static void CheckLexicon(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, LocalRulePack rules,
            VietnameseLexiconSpellChecker lexicon)
        {
            if (lexicon.Count == 0) return;
            foreach (Match match in VietnameseToken.Matches(paragraph.Text))
            {
                if (!match.Success || match.Length == 0 ||
                    ShouldIgnoreLexiconToken(match.Value) || lexicon.IsKnown(match.Value) ||
                    OverlapsExistingTextFinding(findings, paragraph, match.Index, match.Length))
                    continue;
                var suggestion = lexicon.FindDeterministicCorrection(match.Value);
                findings.Add(Span("LOCAL-TYPO-LEXICON", paragraph, match.Index, match.Length,
                    "Từ “" + match.Value + "” không có trong từ điển tiếng Việt.",
                    suggestion == null
                        ? "Kiểm tra và thay “" + match.Value + "” bằng từ tiếng Việt đúng theo ngữ cảnh."
                        : "Sửa thành “" + suggestion + "”.",
                    rules));
            }
        }

        private static bool ShouldIgnoreLexiconToken(string value)
        {
            var letters = value.Where(char.IsLetter).ToArray();
            if (letters.Length == 0) return true;
            if (letters.Length > 1 && letters.All(char.IsUpper)) return true;
            // Unknown title-case tokens are normally personal names, place names or
            // organisation names. The contextual correction list still checks known
            // mistakes inside those phrases before this exemption is applied.
            if (char.IsUpper(letters[0])) return true;
            if (AllowedForeignWords.Contains(value)) return true;
            if (letters.Length == 1)
                return "aàạảãáeèẹẻẽéêềệểễếiìịỉĩíoòọỏõóôồộổỗốơờợởỡớuùụủũúưừừữứyỳỵỷỹý".IndexOf(char.ToLower(letters[0], Vietnamese)) >= 0;
            return false;
        }

        private static bool OverlapsExistingTextFinding(IEnumerable<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, int start, int length)
        {
            var end = start + length;
            return findings.Any(item => item.Anchor.Kind == AnnotationAnchorKind.TextSpan &&
                item.Anchor.ParagraphIndex == paragraph.Index &&
                item.Anchor.StartOffset.HasValue && item.Anchor.Length.HasValue &&
                start < item.Anchor.StartOffset.Value + item.Anchor.Length.Value &&
                item.Anchor.StartOffset.Value < end);
        }

        private static void CheckTelex(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            foreach (var rule in rules.TelexRules)
            {
                Regex regex;
                try { regex = new Regex(rule.Pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100)); }
                catch (ArgumentException) { continue; }
                AddMatches(findings, paragraph, regex, "LOCAL-TYPO-TELEX", "Có dấu hiệu Telex chưa chuyển xong.", "Kiểm tra và đổi thành “" + rule.Replacement + "”.", rules);
            }
        }

        private static void CheckBareShortDates(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph,
            LocalRulePack rules)
        {
            foreach (Match match in LegalBasisShortDate.Matches(paragraph.Text))
            {
                var prefix = paragraph.Text.Substring(0, match.Index);
                if (Rx(@"\bngày\s*$", true).IsMatch(prefix)) continue;
                findings.Add(Span("ND30-PL1-M2-K6B-DATE", paragraph, match.Index, match.Length,
                    "Ngày viết dạng số nhưng thiếu từ “ngày” phía trước.",
                    "Thêm từ “ngày” trước ngày tháng, ví dụ: ngày " + match.Value + ".", rules));
            }
        }

        private static void CheckSentenceCapitalization(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph,
            LocalRulePack rules, IDictionary<int, string> roles)
        {
            string role;
            if (roles.TryGetValue(paragraph.Index, out role) &&
                (role == "subject" || role == "officialLetterSubject" || role == "nationalMotto" ||
                 role == "recipientLabel" || role == "recipientList" || role == "recipientSalutation" ||
                 role == "recipientSalutationInline" || role == "recipientSalutationList"))
                return;
            bool followsListMarker;
            var first = FirstContentLetterIndex(paragraph.Text, out followsListMarker);
            var reportedOffsets = new HashSet<int>();
            if (first < paragraph.Text.Length && char.IsLower(paragraph.Text[first]))
            {
                findings.Add(Span("ND30-PL2-M1", paragraph, first, 1,
                    followsListMarker ? "Chữ cái đầu nội dung sau ký hiệu chỉ mục chưa viết hoa." : "Chữ cái đầu đoạn chưa viết hoa.",
                    followsListMarker ? "Viết hoa chữ cái đầu nội dung sau ký hiệu chỉ mục." : "Viết hoa chữ cái đầu câu.", rules));
                reportedOffsets.Add(first);
            }
            foreach (Match match in Rx(@"[.!?]\s+(?<letter>\p{Ll})").Matches(paragraph.Text))
            {
                var group = match.Groups["letter"];
                if (followsListMarker && group.Index < first) continue;
                var remaining = paragraph.Text.Substring(group.Index);
                if (Rx(@"^(tm|kt|tl|tuq|q|ts|ths|pgs|gs|tp|p)\.", true).IsMatch(remaining)) continue;
                if (!reportedOffsets.Add(group.Index)) continue;
                findings.Add(Span("ND30-PL2-M1", paragraph, group.Index, group.Length, "Chữ cái đầu câu chưa viết hoa.", "Viết hoa chữ cái đầu câu.", rules));
            }
        }

        private static int FirstContentLetterIndex(string text, out bool followsListMarker)
        {
            var markerMatch = LeadingListMarkers.Match(text);
            if (markerMatch.Success)
            {
                followsListMarker = true;
                return markerMatch.Groups["letter"].Index;
            }

            followsListMarker = false;
            return text.TakeWhile(c => !char.IsLetter(c)).Count();
        }

        private static void CheckPersonNames(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            foreach (Match match in Rx(@"\b(?:ông|bà|anh|chị)\s+(?<name>\p{L}+(?:\s+\p{L}+){1,5})", true).Matches(paragraph.Text))
            {
                var name = match.Groups["name"];
                if (!IsTitleCase(name.Value))
                    findings.Add(Span("ND30-PL2-M2-K1", paragraph, name.Index, name.Length, "Tên người có chữ cái đầu viết thường.", "Viết hoa các thành tố của tên người.", rules, "Warning"));
            }
        }

        private static void CheckConfiguredCapitalizations(ICollection<AnnotationFinding> findings,
            IEnumerable<LocalParagraphSnapshot> paragraphs, LocalRulePack rules, IDictionary<int, string> roles)
        {
            var codes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "administrative", "ND30-PL2-M3-K1A" }, { "geographic", "ND30-PL2-M3-K1C" },
                { "terrain", "ND30-PL2-M3-K1D" }, { "region", "ND30-PL2-M3-K1E" },
                { "organ", "ND30-PL2-M4-K1A" }, { "specialOrgan", "ND30-PL2-M4-K1B" },
                { "holiday", "ND30-PL2-M5-K5" }, { "lunarYear", "ND30-PL2-M5-K8A" }
            };
            foreach (var rule in rules.Capitalizations)
            {
                string code;
                if (!codes.TryGetValue(rule.Category, out code) || string.IsNullOrWhiteSpace(rule.Expected)) continue;
                foreach (var paragraph in paragraphs)
                {
                string role;
                if (roles.TryGetValue(paragraph.Index, out role) && IsCapitalizationExemptRole(role)) continue;
                if (IsMostlyUppercase(paragraph.Text)) continue;
                foreach (var occurrence in WholePhraseMatches(paragraph.Text, rule.Expected))
                {
                    var actual = paragraph.Text.Substring(occurrence.Item1, occurrence.Item2);
                    if (!string.Equals(actual, rule.Expected, StringComparison.Ordinal))
                        findings.Add(Span(code, paragraph, occurrence.Item1, occurrence.Item2, "Tên riêng sai chữ hoa.", "Viết “" + rule.Expected + "”.", rules,
                            rule.Category == "organ" || rule.Category == "administrative" ? "Warning" : "Error"));
                }
                }
            }
        }

        private static void CheckAdministrativeNumerals(ICollection<AnnotationFinding> findings,
            IEnumerable<LocalParagraphSnapshot> paragraphs, LocalRulePack rules)
        {
            var regex = Rx(@"\b(?<unit>quận|phường|xã|huyện)\s+(?<roman>[ivxlcdm]+)\b", true);
            foreach (var paragraph in paragraphs)
            foreach (Match match in regex.Matches(paragraph.Text))
            {
                var roman = match.Groups["roman"];
                if (roman.Value.Any(char.IsLower))
                    findings.Add(Span("ND30-PL2-M3-K1B", paragraph, roman.Index, roman.Length, "Số La Mã trong tên đơn vị hành chính viết thường.", "Viết hoa số La Mã.", rules));
            }
        }

        private static void CheckArticleClauseCapitalization(ICollection<AnnotationFinding> findings,
            LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            foreach (Match match in Rx(@"\b(?<keyword>chương|phần|mục|tiểu\s+mục|phụ\s+lục)\s+(?:[IVXLCDM]+|\d+)", false).Matches(paragraph.Text))
            {
                var keyword = match.Groups["keyword"];
                if (keyword.Value.Length > 0 && char.IsLower(keyword.Value[0]))
                    findings.Add(Span("ND30-PL2-M5-K7", paragraph, match.Index, match.Length, "Tên cấp cấu trúc chưa viết hoa.", "Viết hoa chữ cái đầu của Chương, Phần, Mục, Phụ lục.", rules));
            }
            if (Rx(@"Điều\s+\d+", true).IsMatch(paragraph.Text))
            foreach (Match match in Rx(@"\b(?<keyword>Khoản|Điểm)\s+(?:\d+|[a-zđ]\))", false).Matches(paragraph.Text))
            {
                var keyword = match.Groups["keyword"];
                if (keyword.Value.Length > 0 && char.IsUpper(keyword.Value[0]))
                    findings.Add(Span("ND30-PL2-M5-K7", paragraph, keyword.Index, keyword.Length, "Khoản hoặc điểm trong ngữ cảnh viện dẫn Điều đang viết hoa.", "Viết thường khoản, điểm trong cùng cụm viện dẫn.", rules));
            }
        }

        private static bool IsBody(LocalParagraphSnapshot paragraph, IDictionary<int, string> roles)
        {
            if (paragraph.IsInTable || paragraph.StoryType != "wdMainTextStory" || string.IsNullOrWhiteSpace(paragraph.Text)) return false;
            if (roles.ContainsKey(paragraph.Index)) return false;
            var text = paragraph.Text.Trim();
            if (text.Length < 20 || Article.IsMatch(text) || Clause.IsMatch(text) || Point.IsMatch(text)) return false;
            if (IsMostlyUppercase(text) || Rx(@"^(?:[-–—]\s*)?(Căn cứ|Xét|Theo đề nghị)\b", true).IsMatch(text)) return false;
            return true;
        }

        private static IEnumerable<LocalParagraphSnapshot> Scannable(LocalScanSnapshot snapshot) =>
            snapshot.Paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.Text));

        private static IEnumerable<LocalParagraphSnapshot> WithRole(LocalScanSnapshot snapshot, IDictionary<int, string> roles, string role) =>
            snapshot.Paragraphs.Where(p => roles.ContainsKey(p.Index) && roles[p.Index] == role);

        private static LocalParagraphSnapshot LastSubjectParagraph(LocalScanSnapshot snapshot,
            IDictionary<int, string> roles, LocalParagraphSnapshot first)
        {
            var current = first;
            while (true)
            {
                var nextIndex = current.Index + 1;
                string nextRole;
                if (!roles.TryGetValue(nextIndex, out nextRole) || nextRole != "subjectContinuation")
                    return current;
                var next = snapshot.Paragraphs.FirstOrDefault(paragraph => paragraph.Index == nextIndex &&
                    string.Equals(paragraph.StoryType, first.StoryType, StringComparison.Ordinal));
                if (next == null) return current;
                current = next;
            }
        }

        private static int? FirstRole(IDictionary<int, string> roles, params string[] names)
        {
            var match = roles.Where(x => names.Contains(x.Value)).OrderBy(x => x.Key).FirstOrDefault();
            return match.Equals(default(KeyValuePair<int, string>)) ? (int?)null : match.Key;
        }

        private static void CheckStyle(ICollection<AnnotationFinding> findings, string code, LocalParagraphSnapshot paragraph,
            LocalRulePack rules, double minSize, double maxSize, bool bold, bool italic, int alignment, string label)
        {
            var issues = new List<string>();
            if (!string.IsNullOrWhiteSpace(paragraph.FontName) && !Eq(paragraph.FontName, rules.BodyFontName))
                issues.Add("đang dùng phông " + paragraph.FontName);
            if (paragraph.FontSizePoints.HasValue && (paragraph.FontSizePoints.Value < minSize - .1 || paragraph.FontSizePoints.Value > maxSize + .1))
                issues.Add("đang dùng cỡ chữ " + paragraph.FontSizePoints.Value.ToString("0.#", CultureInfo.InvariantCulture));
            if (paragraph.Bold.HasValue && paragraph.Bold.Value != bold)
                issues.Add(paragraph.Bold.Value ? "đang in đậm" : "không in đậm");
            if (paragraph.Italic.HasValue && paragraph.Italic.Value != italic)
                issues.Add(paragraph.Italic.Value ? "đang in nghiêng" : "không in nghiêng");
            if (paragraph.Alignment.HasValue && paragraph.Alignment.Value != alignment)
                issues.Add("đang " + AlignmentDescription(paragraph.Alignment.Value));
            if (issues.Count > 0)
            {
                var expected = "Định dạng " + label + ": phông " + rules.BodyFontName +
                    ", " + SizeDescription(minSize, maxSize) +
                    ", " + (bold ? "in đậm" : "không in đậm") +
                    ", " + (italic ? "in nghiêng" : "không in nghiêng") +
                    ", " + AlignmentDescription(alignment) + ".";
                findings.Add(Paragraph(code, paragraph,
                    label + " sai thể thức: " + string.Join(", ", issues) + ".", expected, rules));
            }
        }

        private static string SizeDescription(double minSize, double maxSize)
        {
            var min = minSize.ToString("0.#", CultureInfo.InvariantCulture);
            var max = maxSize.ToString("0.#", CultureInfo.InvariantCulture);
            return Math.Abs(minSize - maxSize) < .1 ? "cỡ chữ " + min : "cỡ chữ " + min + "–" + max;
        }

        private static string AlignmentDescription(int alignment)
        {
            switch (alignment)
            {
                case 0: return "căn trái";
                case 1: return "căn giữa";
                case 2: return "căn phải";
                case 3: return "căn đều hai lề";
                case 4: return "căn phân tán";
                default: return "căn lề theo quy định";
            }
        }

        private static void ExpectedSize(string role, LocalRulePack rules, out double min, out double max)
        {
            min = rules.BodyFontMinPoints; max = rules.BodyFontMaxPoints;
            if (role == "nationalTitle" || role == "superiorOrganName" || role == "organName") { min = 12; max = 13; }
            else if (role == "appendixLabel") { min = 14; max = 14; }
            else if (role == "appendixDigitalSignatureInfo") { min = 10; max = 10; }
            else if (role == "recipientLabel" || role == "recipientList") { min = 11; max = 12; }
        }

        private static bool IsParty(LocalScanSnapshot snapshot) =>
            string.Equals(snapshot.RegimeCode, "PARTY_HD05", StringComparison.OrdinalIgnoreCase);

        private static bool IsNd30LineSpacingValid(LocalParagraphSnapshot paragraph)
        {
            // Word uses wdLineSpaceSingle=0 and wdLineSpace1pt5=1. For exact/at-least
            // values, validate the effective point range against the current font size.
            var rule = paragraph.LineSpacingRule.GetValueOrDefault();
            var size = paragraph.FontSizePoints.GetValueOrDefault(13d);
            if (rule == 0 || rule == 1)
                return paragraph.LineSpacingPoints.GetValueOrDefault(size) <= size * 1.5d + .5d;
            return paragraph.LineSpacingPoints.GetValueOrDefault() >= size - .1d &&
                paragraph.LineSpacingPoints.GetValueOrDefault() <= size * 1.5d + .5d;
        }

        private static bool IsCapitalizationExemptRole(string role)
        {
            switch (role)
            {
                case "nationalTitle":
                case "nationalMotto":
                case "partyTitle":
                case "superiorOrganName":
                case "organName":
                case "typeName":
                case "subject":
                case "signerAuthority":
                case "appendixLabel":
                case "structuralTitle":
                    return true;
                default:
                    return false;
            }
        }

        private static void AddMatches(ICollection<AnnotationFinding> findings, LocalParagraphSnapshot paragraph, Regex regex,
            string code, string issue, string expected, LocalRulePack rules)
        {
            foreach (Match match in regex.Matches(paragraph.Text))
                if (match.Success && match.Length > 0) findings.Add(Span(code, paragraph, match.Index, match.Length, issue, expected, rules));
        }

        private static AnnotationFinding Section(string code, int section, string issue, string expected, LocalRulePack rules, string severity = "Error") =>
            new AnnotationFinding(code + "-S" + section, code, severity, issue, expected, CitationText(rules),
                new AnnotationAnchor(AnnotationAnchorKind.Section, "wdMainTextStory", null, null, null, string.Empty, section));

        private static AnnotationFinding Paragraph(string code, LocalParagraphSnapshot paragraph, string issue, string expected,
            LocalRulePack rules, string severity = "Warning") =>
            new AnnotationFinding(code + "-P" + paragraph.Index, code, severity, issue, expected, CitationText(rules),
                Anchor(AnnotationAnchorKind.Paragraph, paragraph, null, null, string.Empty));

        private static AnnotationFinding Span(string code, LocalParagraphSnapshot paragraph, int offset, int length,
            string issue, string expected, LocalRulePack rules, string severity = "Warning")
        {
            var safeOffset = Math.Max(0, Math.Min(offset, paragraph.Text.Length));
            var safeLength = Math.Max(0, Math.Min(length, paragraph.Text.Length - safeOffset));
            return new AnnotationFinding(code + "-P" + paragraph.Index + "-C" + safeOffset, code, severity, issue, expected,
                CitationText(rules), Anchor(AnnotationAnchorKind.TextSpan, paragraph, safeOffset, safeLength,
                    paragraph.Text.Substring(safeOffset, safeLength)));
        }

        private static AnnotationAnchor Anchor(AnnotationAnchorKind kind, LocalParagraphSnapshot paragraph,
            int? offset, int? length, string expectedText) =>
            new AnnotationAnchor(kind, paragraph.StoryType, paragraph.Index, offset, length, expectedText,
                paragraph.SectionIndex, paragraph.TableIndex, paragraph.RowIndex, paragraph.CellIndex);

        private static string CitationText(LocalRulePack rules) => "Gói quy tắc ký " + rules.PackId + ", phiên bản " + rules.Version;
        private static Regex Rx(string pattern, bool ignoreCase = false) => new Regex(pattern,
            RegexOptions.CultureInvariant | (ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None), TimeSpan.FromMilliseconds(200));
        private static bool Eq(string left, string right) => string.Equals(Collapse(left), Collapse(right), StringComparison.OrdinalIgnoreCase);
        private static string Collapse(string value) => Regex.Replace(value ?? string.Empty, @"\s+", " ").Trim();
        private static bool Near(double left, double right, double tolerance) => Math.Abs(left - right) <= tolerance;
        private static bool Between(double points, double minMm, double maxMm)
        {
            var mm = points / PointsPerMillimeter;
            return mm >= minMm - .5d && mm <= maxMm + .5d;
        }
        private static bool IsTitleCase(string value)
        {
            return value.Split(new[] { ' ', '-', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(word => word.Any(char.IsLetter)).All(word => char.IsUpper(word.First(char.IsLetter)));
        }
        private static bool IsMostlyUppercase(string value)
        {
            var letters = value.Where(char.IsLetter).ToArray();
            return letters.Length >= 4 && letters.Count(char.IsUpper) >= Math.Ceiling(letters.Length * .8d);
        }
        private static IEnumerable<Tuple<int, int>> WholePhraseMatches(string text, string phrase)
        {
            if (string.IsNullOrWhiteSpace(phrase)) yield break;
            var offset = 0;
            while (offset <= text.Length - phrase.Length)
            {
                var index = Vietnamese.CompareInfo.IndexOf(text, phrase, offset, CompareOptions.IgnoreCase);
                if (index < 0) yield break;
                var end = index + phrase.Length;
                if ((index == 0 || !char.IsLetterOrDigit(text[index - 1])) && (end == text.Length || !char.IsLetterOrDigit(text[end])))
                    yield return Tuple.Create(index, phrase.Length);
                offset = index + Math.Max(1, phrase.Length);
            }
        }
        private static string ApplyCase(string source, string target)
        {
            if (source.All(c => !char.IsLetter(c) || char.IsUpper(c))) return target.ToUpper(Vietnamese);
            if (source.Length > 0 && char.IsUpper(source[0])) return char.ToUpper(target[0], Vietnamese) + target.Substring(1);
            return target;
        }

        private static string TrimParagraphTerminator(string text) => (text ?? string.Empty).TrimEnd('\r', '\a');
    }
}
