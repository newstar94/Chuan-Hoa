using System;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using Xunit;

namespace ChuanHoa.Client.Core.Tests.Scanning
{
    public sealed class LatexTypographicScannerTests
    {
        private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");

        private static LocalRulePack Rules(bool enabled = true,
            string[]? enabledRules = null, string[]? autoFixRules = null)
        {
            var profile = new XElement("advisoryProfiles", new XElement("profile",
                new XAttribute("code", AcademicTypographyAdvisoryProfile.ProfileCode),
                new XAttribute("enabled", enabled.ToString().ToLowerInvariant()),
                new XAttribute("detectorPolicyVersion", "1"),
                new XElement("enabledRules", (enabledRules ?? AcademicTypographyRuleCodes.All.ToArray())
                    .Select(code => new XElement("rule", new XAttribute("code", code)))),
                new XElement("autoFixRules", (autoFixRules ?? new[]
                {
                    AcademicTypographyRuleCodes.PaginationKeep,
                    AcademicTypographyRuleCodes.PaginationWidow
                }).Select(code => new XElement("rule", new XAttribute("code", code)))),
                new XElement("thresholds", new XAttribute("headingConfidenceMinimum", ".90"),
                    new XAttribute("bodyConfidenceMinimum", ".95"),
                    new XAttribute("captionMaxBlankParagraphs", "1"),
                    new XAttribute("mathMinimumSignalCount", "1"))));
            var root = new XElement("rulePack", new XAttribute("schema", LocalRulePackParser.Schema),
                new XAttribute("packId", "TEST"), new XAttribute("version", "1.0.0"),
                new XAttribute("notBeforeUtc", Now.AddDays(-1).ToString("O")),
                new XAttribute("expiresAtUtc", Now.AddDays(30).ToString("O")),
                new XAttribute("minimumClientReleaseId", "1.0.0.0"),
                new XElement("format", new XAttribute("a4WidthMm", "210"), new XAttribute("a4HeightMm", "297"),
                    new XAttribute("topMinMm", "20"), new XAttribute("topMaxMm", "25"),
                    new XAttribute("bottomMinMm", "20"), new XAttribute("bottomMaxMm", "25"),
                    new XAttribute("leftMinMm", "30"), new XAttribute("leftMaxMm", "35"),
                    new XAttribute("rightMinMm", "15"), new XAttribute("rightMaxMm", "20"),
                    new XAttribute("bodyFontName", "Times New Roman")),
                new XElement("corrections"), new XElement("lexicon"), new XElement("telex"),
                new XElement("hiddenCharacters"), new XElement("capitalizations"),
                new XElement("documentTypeAbbreviations"), profile);
            return LocalRulePackParser.Parse(Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting)),
                Now, "1.0.0.0");
        }

        private static LocalSectionSnapshot ValidSection()
        {
            const double pt = 72d / 25.4d;
            return new LocalSectionSnapshot(1, 210 * pt, 297 * pt, 20 * pt, 20 * pt, 30 * pt, 15 * pt, false);
        }

        private static LocalParagraphSnapshot P(int index, string text, bool bold = true,
            int? outlineLevel = null, bool? keepWithNext = null, bool? widowControl = null,
            string? styleName = null, bool inTable = false, int? tableIndex = null,
            int? builtInStyleId = null, bool semanticFactsKnown = true, string role = "Unknown",
            string captionKind = "") =>
            new LocalParagraphSnapshot(index, text, "wdMainTextStory", 1, index * 50,
                "Times New Roman", tableIndex: tableIndex, bold: bold, outlineLevel: outlineLevel,
                keepWithNext: keepWithNext, widowControl: widowControl, styleName: styleName,
                isInTable: inTable, role: role, absoluteEnd: index * 50 + text.Length,
                builtInStyleId: builtInStyleId,
                hasField: semanticFactsKnown ? false : (bool?)null,
                hasMathObject: semanticFactsKnown ? false : (bool?)null,
                hasHyperlink: semanticFactsKnown ? false : (bool?)null,
                hasContentControl: semanticFactsKnown ? false : (bool?)null,
                captionKind: captionKind);

        private static LocalScanSnapshot Snapshot(LocalParagraphSnapshot[] paragraphs,
            LocalTableSnapshot[]? tables = null, LocalGraphicObjectSnapshot[]? graphics = null,
            AnnotationProtectedSpan[]? protectedSpans = null) =>
            new LocalScanSnapshot("sha256:test", 1, new[] { ValidSection() }, paragraphs,
                protectedSpans ?? Array.Empty<AnnotationProtectedSpan>(), tables: tables,
                graphicObjects: graphics);

        [Fact]
        public void Signed_profile_absent_or_disabled_produces_no_advisory()
        {
            var paragraph = P(1, "1.1. Mục tiêu nghiên cứu", outlineLevel: 10,
                keepWithNext: false, styleName: "Normal");
            var scanner = new LatexTypographicScanner();
            Assert.Empty(scanner.Scan(Snapshot(new[] { paragraph }), Rules(false,
                Array.Empty<string>(), Array.Empty<string>())));
            var legacy = new LocalRulePack("TEST", "1", Now.AddDays(-1), Now.AddDays(1), "1.0.0.0",
                210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
                Array.Empty<TextCorrectionRule>(), Array.Empty<TelexRule>(), Array.Empty<char>());
            Assert.Empty(scanner.Scan(Snapshot(new[] { paragraph }), legacy));
        }

        [Fact]
        public void Flags_high_confidence_heading_style_and_keep()
        {
            var paragraphs = new[]
            {
                P(1, "1.1. Mục tiêu nghiên cứu", outlineLevel: 10, keepWithNext: false, styleName: "Normal"),
                P(2, "Đây là nội dung giải thích đầy đủ cho đề mục nghiên cứu này.", bold: false,
                    widowControl: true)
            };
            var findings = new LatexTypographicScanner().Scan(Snapshot(paragraphs), Rules());
            Assert.Contains(findings, f => f.RuleCode == AcademicTypographyRuleCodes.SectionStyle);
            Assert.Contains(findings, f => f.RuleCode == AcademicTypographyRuleCodes.PaginationKeep);
            Assert.All(findings, f => Assert.Equal("Suggestion", f.Severity));
        }

        [Fact]
        public void Detects_continuity_within_parent_and_resets_between_blocks()
        {
            var paragraphs = new[]
            {
                P(1, "QUYẾT ĐỊNH", role: "typeName"),
                P(2, "1.1. Mục tiêu A"),
                P(3, "1.3. Mục tiêu B"),
                P(4, "QUYẾT ĐỊNH", role: "typeName"),
                P(5, "1.1. Mục tiêu C")
            };
            var findings = new LatexTypographicScanner().Scan(Snapshot(paragraphs), Rules());
            var continuity = Assert.Single(findings,
                f => f.RuleCode == AcademicTypographyRuleCodes.SectionContinuity);
            Assert.Equal(3, continuity.Anchor.ParagraphIndex);
        }

        [Theory]
        [InlineData("Điều 1. Phạm vi điều chỉnh")]
        [InlineData("1. Nội dung khoản")]
        [InlineData("a) Nội dung điểm;")]
        [InlineData("Căn cứ Nghị định số 30/2020/NĐ-CP;")]
        [InlineData("Số: 01/QĐ-ABC")]
        [InlineData("Hà Nội, ngày 01 tháng 9 năm 2026")]
        public void Does_not_treat_legal_or_list_text_as_academic_heading(string text)
        {
            var paragraph = P(1, text, outlineLevel: 10, keepWithNext: false, styleName: "Normal");
            var findings = new LatexTypographicScanner().Scan(Snapshot(new[] { paragraph }), Rules());
            Assert.DoesNotContain(findings, f => f.RuleCode == AcademicTypographyRuleCodes.SectionStyle ||
                f.RuleCode == AcademicTypographyRuleCodes.PaginationKeep);
        }

        [Fact]
        public void Flags_confident_body_widow_without_length_proxy()
        {
            var paragraph = P(1,
                "Đây là đoạn thân bài có đủ từ và kết thúc đúng dấu câu để kiểm tra ngắt trang.",
                bold: false, widowControl: false);
            var findings = new LatexTypographicScanner().Scan(Snapshot(new[] { paragraph }), Rules());
            Assert.Contains(findings, f => f.RuleCode == AcademicTypographyRuleCodes.PaginationWidow);
        }

        [Fact]
        public void Booktabs_requires_complete_simple_data_table_facts()
        {
            var paragraphs = new[] { P(1, "Chỉ tiêu", inTable: true, tableIndex: 1),
                P(2, "Giá trị", inTable: true, tableIndex: 1) };
            var present = new LocalBorderSnapshot(LocalSnapshotValueState.Present, 1, 1.25);
            var thin = new LocalBorderSnapshot(LocalSnapshotValueState.Present, 1, .5);
            var table = new LocalTableSnapshot(1, 2, 2, headerRowIndexes: new[] { 1 },
                hasMergedCells: false, hasMergedCellsState: false, absoluteStart: 50,
                absoluteEnd: 150, topBorder: present, bottomBorder: present,
                leftBorder: LocalBorderSnapshot.None, rightBorder: LocalBorderSnapshot.None,
                insideVerticalBorder: new LocalBorderSnapshot(LocalSnapshotValueState.Present, 1, .5),
                headerSeparatorBorder: thin);
            var findings = new LatexTypographicScanner().Scan(Snapshot(paragraphs, new[] { table }), Rules());
            Assert.Contains(findings, f => f.RuleCode == AcademicTypographyRuleCodes.TableBooktabs);

            var unknown = new LocalTableSnapshot(1, 2, 2, headerRowIndexes: new[] { 1 },
                hasMergedCells: false, hasMergedCellsState: false, absoluteStart: 50,
                absoluteEnd: 150, topBorder: LocalBorderSnapshot.Unknown,
                bottomBorder: present, leftBorder: LocalBorderSnapshot.None,
                rightBorder: LocalBorderSnapshot.None, insideVerticalBorder: LocalBorderSnapshot.None,
                headerSeparatorBorder: thin);
            var outcome = new CanonicalRuleScanner().ScanFormatDetailed(
                Snapshot(paragraphs, new[] { unknown }), Rules());
            Assert.Contains(AcademicTypographyRuleCodes.TableBooktabs,
                outcome.AcademicTypography.NotEvaluatedRuleCodes);
        }

        [Fact]
        public void Caption_position_uses_table_and_graphic_association()
        {
            var paragraphs = new[]
            {
                P(1, "Dữ liệu", inTable: true, tableIndex: 1),
                P(2, "Bảng 1: Thống kê số liệu"),
                P(3, "Hình 1: Sơ đồ quy trình", captionKind: "Figure"),
                P(4, "Đối tượng hình")
            };
            var table = new LocalTableSnapshot(1, 2, 2, absoluteStart: 50, absoluteEnd: 90);
            var graphic = new LocalGraphicObjectSnapshot(1, "InlineShape:Picture", true,
                "wdMainTextStory", 1, 200, 201, 4);
            var findings = new LatexTypographicScanner().Scan(
                Snapshot(paragraphs, new[] { table }, new[] { graphic }), Rules());
            Assert.Equal(2, findings.Count(f => f.RuleCode == AcademicTypographyRuleCodes.CaptionPosition));
        }

        [Fact]
        public void Math_requires_balanced_signal_and_excludes_currency_and_protected_content()
        {
            var paragraphs = new[]
            {
                P(1, "Ta có $E=mc^2$ và $$\\int_0^1 f(x)dx$$."),
                P(2, "Chi phí là US$100 và tổng $200."),
                P(3, "Ký hiệu thoát \\$không phải toán."),
                P(4, "Công thức $x+y$ trong vùng bảo vệ.")
            };
            var protectedSpan = new AnnotationProtectedSpan("wdMainTextStory", 200,
                paragraphs[3].Text.Length);
            var findings = new LatexTypographicScanner().Scan(
                Snapshot(paragraphs, protectedSpans: new[] { protectedSpan }), Rules());
            Assert.Equal(2, findings.Count(f => f.RuleCode == AcademicTypographyRuleCodes.MathSyntax));
            Assert.All(findings.Where(f => f.RuleCode == AcademicTypographyRuleCodes.MathSyntax),
                f => Assert.Equal(1, f.Anchor.ParagraphIndex));
        }

        [Fact]
        public void Unknown_math_state_is_not_evaluated_not_a_false_pass()
        {
            var paragraph = P(1, "Ta có $E=mc^2$.", semanticFactsKnown: false);
            var outcome = new CanonicalRuleScanner().ScanFormatDetailed(Snapshot(new[] { paragraph }), Rules());
            Assert.DoesNotContain(outcome.Findings, f => f.RuleCode == AcademicTypographyRuleCodes.MathSyntax);
            Assert.Contains(AcademicTypographyRuleCodes.MathSyntax,
                outcome.AcademicTypography.NotEvaluatedRuleCodes);
        }
    }
}
