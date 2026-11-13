using System;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;

namespace ChuanHoa.Client.Core.Tests;

/// <summary>
/// Phụ lục III NĐ 30/2020/NĐ-CP type abbreviations (including the official
/// mixed-case TTr, QyĐ, CTr, TrS), Latin Eth look-alikes of Đ/đ and leftover
/// drafting placeholders.
/// </summary>
public sealed class CodeNotationAndDraftTextTests
{
    private static readonly LocalRulePack Bundled = BundledRulePack.Load("1.0.0.197");

    [Fact]
    public void Bundled_pack_lists_every_nd30_appendix_iii_abbreviation_and_no_cong_van_code()
    {
        var expected = new (string TypeName, string Abbreviation)[]
        {
            ("Nghị quyết", "NQ"), ("Quyết định", "QĐ"), ("Chỉ thị", "CT"), ("Quy chế", "QC"),
            ("Quy định", "QyĐ"), ("Thông cáo", "TC"), ("Thông báo", "TB"), ("Hướng dẫn", "HD"),
            ("Chương trình", "CTr"), ("Kế hoạch", "KH"), ("Phương án", "PA"), ("Đề án", "ĐA"),
            ("Dự án", "DA"), ("Báo cáo", "BC"), ("Biên bản", "BB"), ("Tờ trình", "TTr"),
            ("Hợp đồng", "HĐ"), ("Công điện", "CĐ"), ("Bản ghi nhớ", "BGN"), ("Bản thỏa thuận", "BTT"),
            ("Giấy ủy quyền", "GUQ"), ("Giấy mời", "GM"), ("Giấy giới thiệu", "GGT"),
            ("Giấy nghỉ phép", "GNP"), ("Phiếu gửi", "PG"), ("Phiếu chuyển", "PC"), ("Phiếu báo", "PB"),
            ("Bản sao y", "SY"), ("Bản trích sao", "TrS"), ("Bản sao lục", "SL")
        };

        foreach (var item in expected)
            Assert.Contains(Bundled.DocumentTypeAbbreviations, rule =>
                rule.TypeName == item.TypeName && rule.Abbreviation == item.Abbreviation);
        Assert.DoesNotContain(Bundled.DocumentTypeAbbreviations, rule => rule.TypeName == "Công văn");
    }

    [Theory]
    [InlineData("TỜ TRÌNH", "Số: 15/TTr-UBND")]
    [InlineData("QUY ĐỊNH", "Số: 15/QyĐ-UBND")]
    [InlineData("CHƯƠNG TRÌNH", "Số: 15/CTr-UBND")]
    [InlineData("GIẤY MỜI", "Số: 15/GM-UBND")]
    public void Official_type_abbreviation_is_accepted_in_its_official_case(string typeName, string codeNumber)
    {
        var findings = ScanCode(typeName, codeNumber);

        Assert.DoesNotContain(findings, item => item.RuleCode is "ND30-PL1-M2-K3-CASE" or "ND30-PL1-M2-K3-ABBR");
    }

    [Fact]
    public void Upper_cased_mixed_case_abbreviation_is_reported_once_with_the_official_form()
    {
        var findings = ScanCode("TỜ TRÌNH", "Số: 15/TTR-UBND");

        var abbreviation = Assert.Single(findings, item => item.RuleCode == "ND30-PL1-M2-K3-ABBR");
        Assert.Equal("TTR", abbreviation.Anchor.ExpectedText);
        Assert.Contains("TTr", abbreviation.Expected, StringComparison.Ordinal);
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K3-CASE");
    }

    [Fact]
    public void Lower_case_issuer_group_is_reported_without_touching_the_official_type_group()
    {
        var findings = ScanCode("TỜ TRÌNH", "Số: 15/TTr-ubnd");

        var upperCase = Assert.Single(findings, item => item.RuleCode == "ND30-PL1-M2-K3-CASE");
        Assert.Equal("ubnd", upperCase.Anchor.ExpectedText);
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K3-ABBR");
    }

    [Fact]
    public void Newly_listed_type_reports_a_mismatched_abbreviation()
    {
        var findings = ScanCode("GIẤY MỜI", "Số: 15/TB-UBND");

        var abbreviation = Assert.Single(findings, item => item.RuleCode == "ND30-PL1-M2-K3-ABBR");
        Assert.Contains("GM", abbreviation.Expected, StringComparison.Ordinal);
    }

    [Fact]
    public void Official_mixed_case_group_is_accepted_without_a_detected_type_name()
    {
        var paragraph = P(1, "Số: 15/TTr-UBND", "codeNumber");
        var findings = Format(new[] { paragraph });

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K3-CASE");
    }

    [Fact]
    public void Party_regulation_keeps_the_hd05_qd_abbreviation()
    {
        Assert.Equal("QĐ", DocumentTypeAbbreviationPolicy.Resolve(Bundled, "QUY ĐỊNH", party: true)!.Abbreviation);
        Assert.Equal("QyĐ", DocumentTypeAbbreviationPolicy.Resolve(Bundled, "QUY ĐỊNH", party: false)!.Abbreviation);
        Assert.Equal("TTr", DocumentTypeAbbreviationPolicy.Resolve(Bundled, " Tờ  trình: ", party: true)!.Abbreviation);
        Assert.Null(DocumentTypeAbbreviationPolicy.Resolve(Bundled, "CÔNG VĂN", party: false));
    }

    [Theory]
    [InlineData("Số: 5/ttr-ubnd", false, "TTr", "Số: 05/TTr-UBND")]
    [InlineData("Số: 15/TTR-UBND", false, "TTr", "Số: 15/TTr-UBND")]
    [InlineData("Số: 15/TTr-ubnd", false, "", "Số: 15/TTr-UBND")]
    [InlineData("Số: 7/2026/qyđ-ubnd", false, "QyĐ", "Số: 07/QyĐ-UBND")]
    [InlineData("Số 5-ttr/tu", true, "TTr", "Số 05-TTr/TU")]
    [InlineData("Số: 129/2026/QĐ-TTĐ.BMC", false, "QĐ", "Số: 129/QĐ-TTĐ-BMC")]
    public void Code_number_normalizer_writes_the_official_case_of_the_type_group(
        string source, bool party, string abbreviation, string expected)
    {
        Assert.Equal(expected, LocalAdministrativeTextNormalizer.NormalizeCodeNumber(source, party, abbreviation));
    }

    [Fact]
    public void Latin_eth_lookalike_is_reported_in_names_and_headings_with_an_exact_replacement()
    {
        var paragraph = P(1, "ÐIỀU 1. Ðối tượng áp dụng tại quận Hà Ðông.");
        var findings = new LocalDocumentScanner().ScanSpelling(Snapshot(paragraph), Bundled).Findings;

        var lookalikes = findings.Where(item => item.RuleCode == "LOCAL-TYPO-LEXICON").ToArray();
        Assert.Contains(lookalikes, item => item.Anchor.ExpectedText == "ÐIỀU" && item.Expected.Contains("“ĐIỀU”"));
        Assert.Contains(lookalikes, item => item.Anchor.ExpectedText == "Ðối" && item.Expected.Contains("“Đối”"));
        Assert.Contains(lookalikes, item => item.Anchor.ExpectedText == "Ðông" && item.Expected.Contains("“Đông”"));
    }

    [Fact]
    public void Latin_eth_lookalike_correction_is_deterministic_and_keeps_case()
    {
        var lexicon = new VietnameseLexiconSpellChecker(Bundled.Lexicon);

        Assert.Equal("ĐẢNG", lexicon.FindDeterministicCorrection("ÐẢNG"));
        Assert.Equal("đồng", lexicon.FindDeterministicCorrection("ðồng"));
        Assert.Equal("ĐHQG", lexicon.FindDeterministicCorrection("ÐHQG"));
        Assert.Null(lexicon.FindLookalikeCorrection("Đảng"));
        Assert.True(VietnameseLookalikeLetters.ContainsEth("Ðà Nẵng"));
        Assert.Equal("Đà Nẵng", VietnameseLookalikeLetters.FoldEth("Ðà Nẵng"));
    }

    [Theory]
    [InlineData("Kinh phí thực hiện: [CẦN BỔ SUNG: số tiền].", "[CẦN BỔ SUNG: số tiền]")]
    [InlineData("Giao [Tên cơ quan] chủ trì thực hiện.", "[Tên cơ quan]")]
    [InlineData("Người ký: <họ và tên>.", "<họ và tên>")]
    [InlineData("Số lượng: ??? người.", "???")]
    [InlineData("Kính gửi: «Ho_ten».", "«Ho_ten»")]
    [InlineData("Kính gửi: <<Đơn vị nhận>>.", "<<Đơn vị nhận>>")]
    public void Leftover_drafting_placeholder_is_reported(string text, string placeholder)
    {
        var findings = new LocalDocumentScanner().ScanSpelling(Snapshot(P(1, text)), Bundled).Findings;

        var finding = Assert.Single(findings, item => item.RuleCode == "LOCAL-TYPO-PLACEHOLDER");
        Assert.Equal(placeholder, finding.Anchor.ExpectedText);
        Assert.Equal("Warning", finding.Severity);
    }

    [Theory]
    [InlineData("Số: .../QĐ-UBND ngày ... tháng ... năm ...")]
    [InlineData("Trích dẫn: “Nhà nước [...] bảo đảm quyền con người”.")]
    [InlineData("Theo tài liệu [1] và [2].")]
    [InlineData("Chỉ tiêu a < b và c > d.")]
    [InlineData("Dòng mã <br> trong nội dung.")]
    [InlineData("Có thật không? Có.")]
    public void Authored_fill_in_areas_and_ordinary_brackets_are_not_placeholders(string text)
    {
        var findings = new LocalDocumentScanner().ScanSpelling(Snapshot(P(1, text)), Bundled).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "LOCAL-TYPO-PLACEHOLDER");
    }

    private static AnnotationFinding[] ScanCode(string typeName, string codeNumber) =>
        Format(new[] { P(1, typeName, "typeName"), P(2, codeNumber, "codeNumber") });

    private static AnnotationFinding[] Format(LocalParagraphSnapshot[] paragraphs) =>
        new LocalDocumentScanner().ScanFormat(new LocalScanSnapshot("sha256:code-notation", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>()), Bundled).Findings.ToArray();

    private static LocalScanSnapshot Snapshot(LocalParagraphSnapshot paragraph) =>
        new("sha256:draft-text", 1, new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

    private static LocalParagraphSnapshot P(int index, string text, string role = "Unknown") =>
        new(index, text, "wdMainTextStory", 1, index * 100, "Times New Roman", fontSizePoints: 13,
            bold: false, italic: false, alignment: 3, role: role);

    private static LocalSectionSnapshot ValidSection()
    {
        const double point = 72d / 25.4d;
        return new LocalSectionSnapshot(1, 210 * point, 297 * point, 20 * point, 20 * point, 30 * point,
            15 * point, false, true);
    }
}
