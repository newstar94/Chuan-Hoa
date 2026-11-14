using System;
using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

/// <summary>
/// VBQPPL regime: Phụ lục I and Mẫu số 17-20 of Phụ lục III, Nghị định 187/2025/NĐ-CP
/// (which replaces the appendices of Nghị định 78/2025/NĐ-CP). Source text:
/// shared/docs/extracted_text/PL1_ND187.txt and PL3_ND187.txt.
/// </summary>
public sealed class VbqpplRegimeTests
{
    private const double Pt = 72d / 25.4d;
    private static readonly LocalRulePack Bundled = BundledRulePack.Load("1.0.0.197");
    private static readonly string[] Nd30CodeNumberCodes =
    {
        "ND30-PL1-M2-K3-PREFIX", "ND30-PL1-M2-K3-SEP", "ND30-PL1-M2-K3-ABBR", "ND30-PL1-M2-K3-SPACE",
        "ND30-PL1-M2-K3-CASE", "ND30-PL1-M2-K3-PAD"
    };

    // ---- Detection -------------------------------------------------------------------

    [Theory]
    [InlineData("Số: 05/2025/QĐ-UBND", VbqpplDocumentKind.UbndDecision)]
    [InlineData("Số: 12/2025/NQ-HĐND", VbqpplDocumentKind.HdndResolution)]
    [InlineData("Số 12 / 2025 / NQ - HĐND", VbqpplDocumentKind.HdndResolution)]
    [InlineData("Số: 05/2025/qđ-ubnd", VbqpplDocumentKind.UbndDecision)]
    [InlineData("Số: 05/QĐ-UBND", VbqpplDocumentKind.Other)]
    [InlineData("Số: 12/NQ-HĐND", VbqpplDocumentKind.Other)]
    [InlineData("Số: 05/25/QĐ-UBND", VbqpplDocumentKind.Other)]
    [InlineData("Số: 15/2025/TT-BTC", VbqpplDocumentKind.Other)]
    [InlineData("Số: 05/2025/QĐ-BTC", VbqpplDocumentKind.Other)]
    [InlineData("", VbqpplDocumentKind.Other)]
    public void Kind_is_read_only_from_a_year_bearing_hdnd_or_ubnd_notation(string code, VbqpplDocumentKind expected)
    {
        Assert.Equal(expected, VbqpplRegimePolicy.KindFromCodeNumber(code));
    }

    [Fact]
    public void Year_bearing_ubnd_decision_is_detected_as_vbqppl_when_regime_was_not_chosen()
    {
        var snapshot = UbndDecision(regime: LocalRegimeCodes.Nd30, manual: false);

        Assert.Equal(LocalRegimeCodes.Vbqppl, new DocumentRoleDetector().ResolveRegime(snapshot));
        Assert.Equal(VbqpplDocumentKind.UbndDecision, new DocumentRoleDetector().DetectVbqpplKind(snapshot));
    }

    [Fact]
    public void Detection_works_from_unlabelled_header_text()
    {
        var paragraphs = new[]
        {
            P(1, "HỘI ĐỒNG NHÂN DÂN TỈNH AN GIANG", bold: true, alignment: 1),
            P(2, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold: true, alignment: 1),
            P(3, "Độc lập - Tự do - Hạnh phúc", bold: true, alignment: 1),
            P(4, "Số: 12/2025/NQ-HĐND", alignment: 1),
            P(5, "NGHỊ QUYẾT", size: 14, bold: true, alignment: 1),
            P(6, "Quy định mức chi hỗ trợ", size: 14, bold: true, alignment: 1),
            P(7, "Căn cứ Luật Tổ chức chính quyền địa phương;", size: 14, italic: true),
            P(8, "Hội đồng nhân dân ban hành Nghị quyết quy định mức chi hỗ trợ.", size: 14, italic: true),
            P(9, "Điều 1. Phạm vi điều chỉnh", bold: true)
        };
        var snapshot = Snapshot(paragraphs, LocalRegimeCodes.Nd30, false);

        Assert.Equal(LocalRegimeCodes.Vbqppl, new DocumentRoleDetector().ResolveRegime(snapshot));
    }

    [Fact]
    public void Administrative_ubnd_decision_without_year_stays_nd30()
    {
        var snapshot = UbndDecision(code: "Số: 05/QĐ-UBND", regime: LocalRegimeCodes.Nd30, manual: false);

        Assert.Equal(LocalRegimeCodes.Nd30, new DocumentRoleDetector().ResolveRegime(snapshot));
    }

    [Fact]
    public void Manual_selection_wins_over_detection_in_both_directions()
    {
        var detector = new DocumentRoleDetector();

        Assert.Equal(LocalRegimeCodes.Nd30, detector.ResolveRegime(UbndDecision(regime: LocalRegimeCodes.Nd30, manual: true)));
        Assert.Equal(LocalRegimeCodes.PartyHd05, detector.ResolveRegime(UbndDecision(regime: LocalRegimeCodes.PartyHd05, manual: true)));
        Assert.Equal(LocalRegimeCodes.Vbqppl, detector.ResolveRegime(
            UbndDecision(code: "Số: 05/QĐ-UBND", regime: LocalRegimeCodes.Vbqppl, manual: true)));
    }

    [Fact]
    public void Earlier_automatic_vbqppl_reverts_to_nd30_when_the_year_is_gone()
    {
        var snapshot = UbndDecision(code: "Số: 05/QĐ-UBND", regime: LocalRegimeCodes.Vbqppl, manual: false);

        Assert.Equal(LocalRegimeCodes.Nd30, new DocumentRoleDetector().ResolveRegime(snapshot));
    }

    // ---- Scanner: valid documents --------------------------------------------------------

    [Fact]
    public void Valid_ubnd_decision_has_no_vbqppl_findings_and_is_not_checked_as_nd30()
    {
        var findings = Scan(UbndDecision(regime: LocalRegimeCodes.Nd30, manual: false));

        Assert.DoesNotContain(findings, f => f.RuleCode.StartsWith("ND187-", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, f => Nd30CodeNumberCodes.Contains(f.RuleCode));
        Assert.DoesNotContain(findings, f => f.RuleCode is "ND30-PL1-M2-K6A-PUNCT" or "ND30-PL1-M2-K6A-STYLE" or
            "ND30-PL1-M2-K6E-DOTSLASH" or "ND30-PL1-M1-K3" or "ND30-PL1-M2-K5A-TYPE" or "ND30-PL1-M2-K7D-STYLE");
    }

    [Fact]
    public void Valid_hdnd_resolution_has_no_vbqppl_findings_without_a_place_date_line()
    {
        var findings = Scan(HdndResolution());

        Assert.DoesNotContain(findings, f => f.RuleCode.StartsWith("ND187-", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, f => Nd30CodeNumberCodes.Contains(f.RuleCode));
    }

    [Fact]
    public void Manual_nd30_keeps_the_nd30_code_number_check_for_the_same_text()
    {
        var findings = Scan(UbndDecision(regime: LocalRegimeCodes.Nd30, manual: true));

        Assert.Contains(findings, f => Nd30CodeNumberCodes.Contains(f.RuleCode));
        Assert.DoesNotContain(findings, f => f.RuleCode.StartsWith("ND187-", StringComparison.Ordinal));
    }

    [Fact]
    public void Administrative_decision_without_year_keeps_nd30_behavior()
    {
        var findings = Scan(UbndDecision(code: "Số: 05/QĐ-UBND", regime: LocalRegimeCodes.Nd30, manual: false));

        Assert.DoesNotContain(findings, f => f.RuleCode.StartsWith("ND187-", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, f => Nd30CodeNumberCodes.Contains(f.RuleCode));
    }

    // ---- Scanner: defects --------------------------------------------------------------

    [Fact]
    public void Code_number_defects_are_reported_under_the_vbqppl_rule()
    {
        var findings = Scan(UbndDecision(code: "Số 5/2025/qđ-ubnd"));

        var qualifiers = findings.Where(f => f.RuleCode == CanonicalRuleScanner.VbqpplCodeNumberCode)
            .Select(f => f.FindingId.Substring(f.FindingId.LastIndexOf("-Q", StringComparison.Ordinal) + 2))
            .ToArray();
        Assert.Contains("colon", qualifiers);
        Assert.Contains("pad", qualifiers);
        Assert.Contains("case", qualifiers);
        Assert.DoesNotContain("year", qualifiers);
        Assert.DoesNotContain(findings, f => Nd30CodeNumberCodes.Contains(f.RuleCode));
    }

    [Fact]
    public void Manual_vbqppl_without_year_reports_the_missing_year()
    {
        var findings = Scan(UbndDecision(code: "Số: 05/QĐ-UBND", regime: LocalRegimeCodes.Vbqppl, manual: true));

        Assert.Contains(findings, f => f.RuleCode == CanonicalRuleScanner.VbqpplCodeNumberCode &&
            f.FindingId.EndsWith("-Qyear", StringComparison.Ordinal));
    }

    [Fact]
    public void Wrong_type_abbreviation_is_reported()
    {
        var findings = Scan(UbndDecision(code: "Số: 05/2025/NQ-UBND", regime: LocalRegimeCodes.Vbqppl, manual: true));

        Assert.Contains(findings, f => f.RuleCode == CanonicalRuleScanner.VbqpplCodeNumberCode &&
            f.FindingId.EndsWith("-Qabbreviation", StringComparison.Ordinal) && f.Expected.Contains("QĐ"));
    }

    [Fact]
    public void Nd30_operative_heading_instead_of_the_enacting_formula_is_reported()
    {
        var findings = Scan(UbndDecision(formula: "QUYẾT ĐỊNH:"));

        var finding = Assert.Single(findings, f => f.RuleCode == CanonicalRuleScanner.VbqpplLegalBasisCode &&
            f.FindingId.EndsWith("-Qformula", StringComparison.Ordinal));
        Assert.Contains("QUYẾT ĐỊNH:", finding.CurrentIssue, StringComparison.Ordinal);
        Assert.Contains("Ủy ban nhân dân ban hành Quyết định", finding.Expected, StringComparison.Ordinal);
    }

    [Fact]
    public void Last_legal_basis_ends_with_semicolon_and_formula_with_full_stop()
    {
        var findings = Scan(UbndDecision(lastBasis: "Theo đề nghị của Giám đốc Sở Tư pháp.",
            formula: "Ủy ban nhân dân ban hành Quyết định quy định quản lý chợ"));

        Assert.Contains(findings, f => f.FindingId == CanonicalRuleScanner.VbqpplLegalBasisCode + "-P9-Qpunct");
        Assert.Contains(findings, f => f.FindingId == CanonicalRuleScanner.VbqpplLegalBasisCode + "-P10-Qformula-punct");
    }

    [Fact]
    public void Legal_basis_must_be_size_14_italic()
    {
        var findings = Scan(UbndDecision(basisSize: 13));

        Assert.Contains(findings, f => f.FindingId == CanonicalRuleScanner.VbqpplLegalBasisCode + "-P8");
    }

    [Fact]
    public void Ubnd_decision_signer_without_tm_line_is_reported()
    {
        var findings = Scan(UbndDecision(includeTm: false));

        Assert.Contains(findings, f => f.RuleCode == CanonicalRuleScanner.VbqpplSignerCode &&
            f.FindingId.EndsWith("-Qtm", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(13, false)]
    [InlineData(14, true)]
    public void Signer_authority_size_is_13(double size, bool reported)
    {
        var findings = Scan(UbndDecision(signerSize: size));

        Assert.Equal(reported, findings.Any(f => f.FindingId == CanonicalRuleScanner.VbqpplSignerCode + "-P17"));
    }

    [Theory]
    [InlineData(14, false)]
    [InlineData(13, true)]
    public void Type_name_size_is_14(double size, bool reported)
    {
        var findings = Scan(UbndDecision(typeSize: size));

        Assert.Equal(reported, findings.Any(f => f.RuleCode == CanonicalRuleScanner.VbqpplTypeNameCode));
    }

    [Theory]
    [InlineData(30, 15, 15, 15, false)] // top/bottom 15 mm: PL I khoản 3 value, not enforced
    [InlineData(30, 15, 25, 25, false)] // top/bottom 25 mm: sơ đồ value, not enforced
    [InlineData(35, 20, 20, 20, false)]
    [InlineData(29, 15, 20, 20, true)]
    [InlineData(36, 15, 20, 20, true)]
    [InlineData(30, 14, 20, 20, true)]
    [InlineData(30, 21, 20, 20, true)]
    public void Only_left_and_right_margins_are_enforced(double left, double right, double top, double bottom, bool reported)
    {
        var section = new LocalSectionSnapshot(1, 210 * Pt, 297 * Pt, top * Pt, bottom * Pt, left * Pt, right * Pt, false, true);
        var findings = Scan(UbndDecision(section: section));

        Assert.Equal(reported, findings.Any(f => f.RuleCode == CanonicalRuleScanner.VbqpplMarginCode));
        Assert.DoesNotContain(findings, f => f.RuleCode == "ND30-PL1-M1-K3");
    }

    [Fact]
    public void Hdnd_resolution_reports_missing_passage_line()
    {
        var findings = Scan(HdndResolution(passage: null));

        Assert.Contains(findings, f => f.RuleCode == CanonicalRuleScanner.VbqpplPassageCode &&
            f.FindingId.EndsWith("-Qmissing", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("Nghị quyết này đã được Hội đồng nhân dân tỉnh An Giang khóa X, Kỳ họp thứ 5 thông qua ngày 10 tháng 7 năm 2025.", false)]
    [InlineData("Nghị quyết này đã được Hội đồng nhân dân tỉnh An Giang khóa X, Kỳ họp thứ nhất thông qua ngày 10 tháng 7 năm 2025.", false)]
    [InlineData("Nghị quyết này đã được Hội đồng nhân dân tỉnh An Giang khóa 10, Kỳ họp thứ 5 thông qua ngày 10 tháng 7 năm 2025.", true)]
    [InlineData("Nghị quyết này đã được Hội đồng nhân dân tỉnh An Giang khóa X, Kỳ họp thứ V thông qua ngày 10 tháng 7 năm 2025.", true)]
    public void Passage_line_uses_roman_term_and_arabic_session(string passage, bool reported)
    {
        var findings = Scan(HdndResolution(passage: passage));

        Assert.Equal(reported, findings.Any(f => f.RuleCode == CanonicalRuleScanner.VbqpplPassageCode &&
            f.FindingId.EndsWith("-Qnumbering", StringComparison.Ordinal)));
    }

    [Fact]
    public void Passage_line_must_be_italic()
    {
        var findings = Scan(HdndResolution(passageItalic: false));

        Assert.Contains(findings, f => f.FindingId == CanonicalRuleScanner.VbqpplPassageCode + "-P16");
    }

    [Fact]
    public void Hdnd_resolution_signed_tm_is_reported()
    {
        var findings = Scan(HdndResolution(tm: true));

        Assert.Contains(findings, f => f.RuleCode == CanonicalRuleScanner.VbqpplSignerCode &&
            f.FindingId.EndsWith("-Qtm", StringComparison.Ordinal));
    }

    [Fact]
    public void Vbqppl_findings_belong_to_their_own_source_family()
    {
        Assert.Equal(AnnotationSourceFamily.Vbqppl,
            AnnotationSourceFamilyResolver.FromRuleCode(CanonicalRuleScanner.VbqpplCodeNumberCode));
        Assert.All(CanonicalRuleScanner.RegisteredRuleCodes.Where(c => c.StartsWith("ND187-", StringComparison.Ordinal)),
            c => Assert.Equal(AnnotationSourceFamily.Vbqppl, AnnotationSourceFamilyResolver.FromRuleCode(c)));
    }

    // ---- 1-Click normalizer ------------------------------------------------------------

    [Theory]
    [InlineData("Số: 05/2025/QĐ-UBND", "QĐ", "Số: 05/2025/QĐ-UBND")]
    [InlineData("Số 5/2025/qđ-ubnd", "QĐ", "Số: 05/2025/QĐ-UBND")]
    [InlineData("Số: 5 / 2025 / NQ - HĐND", "NQ", "Số: 05/2025/NQ-HĐND")]
    [InlineData("Số: 12/2025/NQ-UBND", "QĐ", "Số: 12/2025/QĐ-UBND")]
    [InlineData("Số: 12/2025/UBND", "QĐ", "Số: 12/2025/QĐ-UBND")]
    [InlineData("Số: 15/2025/tt-btc", "", "Số: 15/2025/TT-BTC")]
    public void Vbqppl_normalizer_keeps_the_year(string input, string abbreviation, string expected)
    {
        Assert.Equal(expected, LocalAdministrativeTextNormalizer.NormalizeVbqpplCodeNumber(input, abbreviation));
    }

    [Theory]
    [InlineData("Số: 05/QĐ-UBND")]
    [InlineData("Số: 05/25/QĐ-UBND")]
    [InlineData("Số: .../2025/QĐ-UBND")]
    public void Vbqppl_normalizer_never_invents_a_year(string input)
    {
        Assert.Equal(input, LocalAdministrativeTextNormalizer.NormalizeVbqpplCodeNumber(input, "QĐ"));
    }

    [Fact]
    public void Nd30_normalizer_behavior_is_unchanged_for_the_legacy_year_form()
    {
        Assert.Equal("Số: 129/QĐ-UBND",
            LocalAdministrativeTextNormalizer.NormalizeCodeNumber("Số: 129/2026/QĐ-UBND", false, "QĐ"));
    }

    // ---- Fixtures ---------------------------------------------------------------------

    internal static IReadOnlyList<AnnotationFinding> Scan(LocalScanSnapshot snapshot) =>
        new LocalDocumentScanner().ScanFormat(snapshot, Bundled).Findings;

    /// <summary>A defective UBND decision under the VBQPPL regime: exercises the margin,
    /// code-number, type-name, legal-basis and signer rules.</summary>
    internal static LocalScanSnapshot BadUbndDecision() => UbndDecision(code: "Số 5/2025/qđ-ubnd",
        typeSize: 13, basisSize: 13, formula: "QUYẾT ĐỊNH:", includeTm: false, signerSize: 14,
        section: new LocalSectionSnapshot(1, 210 * Pt, 297 * Pt, 20 * Pt, 20 * Pt, 25 * Pt, 15 * Pt, false, true));

    /// <summary>A defective HĐND resolution under the VBQPPL regime: exercises the passage rule.</summary>
    internal static LocalScanSnapshot BadHdndResolution() =>
        HdndResolution(passage: "Nghị quyết này đã được Hội đồng nhân dân tỉnh An Giang khóa 10, Kỳ họp thứ 5 thông qua ngày 10 tháng 7 năm 2025.",
            passageItalic: false, tm: true);

    private static LocalScanSnapshot UbndDecision(string code = "Số: 05/2025/QĐ-UBND",
        string regime = LocalRegimeCodes.Nd30, bool manual = false, double typeSize = 14,
        double basisSize = 14, string lastBasis = "Theo đề nghị của Giám đốc Sở Tư pháp;",
        string formula = "Ủy ban nhân dân ban hành Quyết định quy định quản lý chợ.",
        bool includeTm = true, double signerSize = 13, LocalSectionSnapshot? section = null)
    {
        var paragraphs = new List<LocalParagraphSnapshot>
        {
            P(1, "ỦY BAN NHÂN DÂN TỈNH AN GIANG", "organName", bold: true, alignment: 1),
            P(2, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "nationalTitle", bold: true, alignment: 1),
            P(3, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", bold: true, alignment: 1),
            P(4, code, "codeNumber", alignment: 1),
            P(5, "An Giang, ngày 10 tháng 7 năm 2025", "placeAndIssuedDate", italic: true, alignment: 1),
            P(6, "QUYẾT ĐỊNH", "typeName", size: typeSize, bold: true, alignment: 1),
            P(7, "Quy định quản lý chợ trên địa bàn tỉnh", "subject", size: 14, bold: true, alignment: 1),
            P(8, "Căn cứ Luật Tổ chức chính quyền địa phương số 72/2025/QH15 ngày 16 tháng 6 năm 2025 của Quốc hội;",
                "legalBasis", size: basisSize, italic: true),
            P(9, lastBasis, "legalBasis", size: 14, italic: true),
            P(10, formula, size: 14, italic: !formula.StartsWith("QUYẾT", StringComparison.Ordinal),
                bold: formula.StartsWith("QUYẾT", StringComparison.Ordinal),
                alignment: formula.StartsWith("QUYẾT", StringComparison.Ordinal) ? 1 : 3),
            P(11, "Điều 1. Phạm vi điều chỉnh", bold: true),
            P(12, "Quyết định này quy định việc quản lý chợ trên địa bàn tỉnh."),
            P(13, "Điều 2. Hiệu lực thi hành", bold: true),
            P(14, "Quyết định này có hiệu lực từ ngày 20 tháng 7 năm 2025."),
            P(15, "Nơi nhận:", "recipientLabel", size: 12, bold: true, italic: true, alignment: 0, indent: 0),
            P(16, "- Như Điều 2;", "recipientList", size: 11, alignment: 0, indent: 0)
        };
        if (includeTm)
            paragraphs.Add(P(17, "TM. ỦY BAN NHÂN DÂN", "signerAuthority", size: signerSize, bold: true, alignment: 1));
        paragraphs.Add(P(includeTm ? 18 : 17, "CHỦ TỊCH", "signerAuthority", size: signerSize, bold: true, alignment: 1));
        paragraphs.Add(P(19, "Nguyễn Văn An", "signerName", size: 14, bold: true, alignment: 1));
        return Snapshot(paragraphs, regime, manual, section);
    }

    private static LocalScanSnapshot HdndResolution(
        string? passage = "Nghị quyết này đã được Hội đồng nhân dân tỉnh An Giang khóa X, Kỳ họp thứ 5 thông qua ngày 10 tháng 7 năm 2025.",
        bool passageItalic = true, bool tm = false)
    {
        var paragraphs = new List<LocalParagraphSnapshot>
        {
            P(1, "HỘI ĐỒNG NHÂN DÂN TỈNH AN GIANG", "organName", bold: true, alignment: 1),
            P(2, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "nationalTitle", bold: true, alignment: 1),
            P(3, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", bold: true, alignment: 1),
            P(4, "Số: 12/2025/NQ-HĐND", "codeNumber", alignment: 1),
            P(6, "NGHỊ QUYẾT", "typeName", size: 14, bold: true, alignment: 1),
            P(7, "Quy định mức chi hỗ trợ", "subject", size: 14, bold: true, alignment: 1),
            P(8, "Căn cứ Luật Tổ chức chính quyền địa phương số 72/2025/QH15 ngày 16 tháng 6 năm 2025 của Quốc hội;",
                "legalBasis", size: 14, italic: true),
            P(9, "Xét Tờ trình số 20/TTr-UBND ngày 01 tháng 7 năm 2025 của Ủy ban nhân dân tỉnh; Báo cáo thẩm tra của Ban Văn hóa - Xã hội; ý kiến thảo luận của đại biểu Hội đồng nhân dân tại kỳ họp;",
                "legalBasis", size: 14, italic: true),
            P(10, "Hội đồng nhân dân ban hành Nghị quyết quy định mức chi hỗ trợ.", size: 14, italic: true),
            P(11, "Điều 1. Phạm vi điều chỉnh", bold: true),
            P(12, "Nghị quyết này quy định mức chi hỗ trợ."),
            P(13, "Điều 2. Hiệu lực thi hành", bold: true),
            P(14, "Nghị quyết này có hiệu lực từ ngày 20 tháng 7 năm 2025.")
        };
        if (passage != null)
            paragraphs.Add(P(16, passage, size: 14, italic: passageItalic));
        paragraphs.Add(P(17, "Nơi nhận:", "recipientLabel", size: 12, bold: true, italic: true, alignment: 0, indent: 0));
        paragraphs.Add(P(18, "- Như Điều 2;", "recipientList", size: 11, alignment: 0, indent: 0));
        if (tm)
            paragraphs.Add(P(19, "TM. HỘI ĐỒNG NHÂN DÂN", "signerAuthority", size: 13, bold: true, alignment: 1));
        paragraphs.Add(P(20, "CHỦ TỊCH", "signerAuthority", size: 13, bold: true, alignment: 1));
        paragraphs.Add(P(21, "Nguyễn Văn Bình", "signerName", size: 14, bold: true, alignment: 1));
        return Snapshot(paragraphs, LocalRegimeCodes.Nd30, false);
    }

    private static LocalScanSnapshot Snapshot(IReadOnlyList<LocalParagraphSnapshot> paragraphs, string regime,
        bool manual, LocalSectionSnapshot? section = null) =>
        new LocalScanSnapshot("sha256:vbqppl-" + Guid.NewGuid().ToString("N"), 1,
            new[] { section ?? new LocalSectionSnapshot(1, 210 * Pt, 297 * Pt, 20 * Pt, 20 * Pt, 30 * Pt, 15 * Pt, false, true) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>(), regimeCode: regime,
            regimeWasSelectedManually: manual);

    private static LocalParagraphSnapshot P(int index, string text, string role = "Unknown", double size = 14,
        bool bold = false, bool italic = false, int alignment = 3, double indent = 30) =>
        new LocalParagraphSnapshot(index, text, "wdMainTextStory", 1, index * 100, "Times New Roman",
            fontSizePoints: size, bold: bold, italic: italic, alignment: alignment,
            firstLineIndentPoints: indent, spaceBeforePoints: 0, spaceAfterPoints: 6,
            role: role, fontColor: 0, lineSpacingPoints: 12, lineSpacingRule: 0);
}
