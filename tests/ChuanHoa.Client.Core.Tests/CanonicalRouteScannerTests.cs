using System;
using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class CanonicalRouteScannerTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");
    [Theory]
    [InlineData("nationalTitle", "ND30-PL1-M2-K1-C")]
    [InlineData("nationalMotto", "ND30-PL1-M2-K1-C")]
    [InlineData("organName", "ND30-PL1-M2-K2-ORG")]
    [InlineData("superiorOrganName", "ND30-PL1-M2-K2-SUP")]
    public void Header_components_require_single_spacing(string role, string code)
    {
        var p = new LocalParagraphSnapshot(1, "TÊN THÀNH PHẦN", "wdMainTextStory", 1, 0,
            "Times New Roman", role: role, lineSpacingRule: 1, lineSpacingPoints: 18);
        var snapshot = new LocalScanSnapshot("sha256:header-spacing", 1,
            new[] { ValidSection() }, new[] { p }, Array.Empty<AnnotationProtectedSpan>());
        Assert.Contains(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            f => f.RuleCode == code && f.CurrentIssue.Contains("giãn dòng đơn"));
    }

    [Fact]
    public void Registry_contains_product_routes_and_shape_rules_without_removed_tone_or_iy()
    {
        Assert.Equal(76, CanonicalRuleScanner.RegisteredRuleCodes.Count);
        Assert.Equal(76, CanonicalRuleScanner.RegisteredRuleCodes.Distinct(StringComparer.Ordinal).Count());
        Assert.DoesNotContain("ND30-PL1-M2-K9B-LUU", CanonicalRuleScanner.RegisteredRuleCodes);
        Assert.DoesNotContain("ND30-PL1-M1-K2", CanonicalRuleScanner.RegisteredRuleCodes);
        Assert.DoesNotContain(CanonicalRuleScanner.RegisteredRuleCodes,
            code => code.Contains("TONE", StringComparison.OrdinalIgnoreCase) || code.Contains("IY", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Positive_corpus_exercises_every_registered_route_with_an_exact_anchor()
    {
        var scanner = new LocalDocumentScanner();
        var rules = Rules();
        var scans = new[]
        {
            scanner.ScanFormat(BadFormatSnapshot(), rules),
            scanner.ScanFormat(LayoutOnlySnapshot(), rules),
            scanner.ScanSpelling(BadSpellingSnapshot(), rules),
            scanner.ScanSpelling(
                new LocalScanSnapshot("sha256:bad-lexicon", 1, new[] { ValidSection() },
                    new[] { P(1, "từ nhậnn", font: "Times New Roman", size: 13) },
                    Array.Empty<AnnotationProtectedSpan>()),
                Rules(new[] { "từ", "nhận" }))
        };
        var findings = scans.SelectMany(scan => scan.Findings).ToArray();

        var missing = CanonicalRuleScanner.RegisteredRuleCodes
            .Where(code => findings.All(item => item.RuleCode != code))
            .ToArray();

        Assert.True(missing.Length == 0, "Routes not exercised: " + string.Join(", ", missing));
        var duplicateIds = scans.SelectMany(scan => scan.Findings
                .GroupBy(item => item.FindingId, StringComparer.Ordinal)
                .Where(group => group.Count() > 1)
                .Select(group => scan.ScanId + ":" + group.Key))
            .ToArray();
        Assert.True(duplicateIds.Length == 0,
            "Scanner emitted ambiguous duplicate finding ids: " + string.Join(", ", duplicateIds));
        Assert.All(findings, finding =>
        {
            Assert.False(string.IsNullOrWhiteSpace(finding.FindingId));
            Assert.False(string.IsNullOrWhiteSpace(finding.Citation));
            Assert.True(finding.Anchor.Kind == AnnotationAnchorKind.Section || finding.Anchor.ParagraphIndex.HasValue);
            if (finding.Anchor.Kind == AnnotationAnchorKind.TextSpan)
            {
                Assert.True(finding.Anchor.StartOffset.HasValue);
                Assert.True(finding.Anchor.Length.HasValue && finding.Anchor.Length.Value > 0);
                Assert.False(string.IsNullOrEmpty(finding.Anchor.ExpectedText));
            }
        });
    }

    [Fact]
    public void Negative_corpus_with_no_applicable_components_produces_no_findings()
    {
        var snapshot = new LocalScanSnapshot("sha256:negative", 1,
            new[] { ValidSection() }, Array.Empty<LocalParagraphSnapshot>(), Array.Empty<AnnotationProtectedSpan>());
        var scanner = new LocalDocumentScanner();

        Assert.Empty(scanner.ScanFormat(snapshot, Rules()).Findings);
        Assert.Empty(scanner.ScanSpelling(snapshot, Rules()).Findings);
    }

    [Fact]
    public void Boundary_corpus_accepts_exact_margin_and_body_typography_limits()
    {
        const double pt = 72d / 25.4d;
        var paragraphs = new[]
        {
            P(1, "Đây là đoạn nội dung hợp lệ đủ dài để kiểm tra chính xác các giá trị biên thể thức.",
                font: "Times New Roman", size: 13, alignment: 3, indent: 10 * pt, after: 6,
                lineSpacing: 12, color: 0)
        };
        var snapshot = new LocalScanSnapshot("sha256:boundary", 1,
            new[] { new LocalSectionSnapshot(1, 210 * pt, 297 * pt, 20 * pt, 25 * pt, 30 * pt, 20 * pt, false, true) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode is "ND30-PL1-M1-K1" or "ND30-PL1-M1-K3" or
            "ND30-PL1-M1-K4-FONT" or "ND30-PL1-M1-K4-COLOR" or "ND30-PL1-M2-K6E-ALIGN" or
            "ND30-PL1-M2-K6E-INDENT" or "ND30-PL1-M2-K6E-SPACEAFTER" or "ND30-PL1-M2-K6E-LINESPACING");
    }

    [Fact]
    public void Line_shape_rules_accept_real_solid_centered_lines_with_legal_widths()
    {
        var paragraphs = new[]
        {
            P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13, bold: true,
                alignment: 1, page: 1, left: 320, top: 80, width: 180),
            P(2, "CƠ QUAN BAN HÀNH", "organName", size: 13, bold: true,
                alignment: 1, page: 1, left: 80, top: 80, width: 180),
            P(3, "Về công tác phòng, chống lụt bão", "subject", size: 14, bold: true,
                alignment: 1, page: 1, left: 200, top: 180, width: 240)
        };
        var lines = new[]
        {
            Line(1, paragraphs[0], 320, 98, 180),
            Line(2, paragraphs[1], 134, 98, 72),
            Line(3, paragraphs[2], 272, 198, 96)
        };
        var snapshot = new LocalScanSnapshot("sha256:lines-ok", 1, new[] { ValidSection() }, paragraphs,
            Array.Empty<AnnotationProtectedSpan>(), lines, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode.EndsWith("-LINE", StringComparison.Ordinal));
    }

    [Fact]
    public void Motto_line_offset_more_than_eight_points_is_rejected()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var offsetLine = Line(1, paragraph, 329, 98, 180);
        var snapshot = new LocalScanSnapshot("sha256:motto-line-offset", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), new[] { offsetLine }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Theory]
    [InlineData(4)]
    [InlineData(5)]
    public void Motto_line_with_dashed_style_is_reported_as_invalid_style(int dashStyle)
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var dashedLine = Line(1, paragraph, 320, 98, 180, dashStyle);
        var snapshot = new LocalScanSnapshot("sha256:motto-line-dashed-" + dashStyle, 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { dashedLine }, "STATE_ND30");

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");

        Assert.Contains("không phải nét liền", finding.CurrentIssue, StringComparison.Ordinal);
        Assert.Contains("Line Shape nét liền", finding.Expected, StringComparison.Ordinal);
    }

    [Fact]
    public void Motto_line_with_unknown_dash_style_fails_closed()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var unknownStyleLine = new LocalLineShapeSnapshot(1, "Line unknown style", 9,
            paragraph.StoryType, paragraph.SectionIndex, paragraph.AbsoluteStart, paragraph.Index,
            paragraph.PageNumber, 320, 98, 180, 0, 320, 98, 1, 1, true, null, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:motto-line-unknown-style", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { unknownStyleLine }, "STATE_ND30");

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");

        Assert.Contains("không phải nét liền", finding.CurrentIssue, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_dashes_below_motto_do_not_replace_a_solid_line_shape()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var textDashes = P(2, "--------------", size: 13, alignment: 1, page: 1,
            left: 320, top: 98, width: 180);
        var validShape = Line(1, paragraph, 320, 98, 180);
        var snapshot = new LocalScanSnapshot("sha256:motto-text-dashes", 1,
            new[] { ValidSection() }, new[] { paragraph, textDashes },
            Array.Empty<AnnotationProtectedSpan>(), new[] { validShape }, "STATE_ND30");

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");

        Assert.Contains("chuỗi dấu gạch/chấm", finding.CurrentIssue, StringComparison.Ordinal);
        Assert.Contains("Line Shape nét liền", finding.Expected, StringComparison.Ordinal);
    }

    [Fact]
    public void Owned_line_for_paragraph_is_accepted_without_false_findings()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var ownedLine = new LocalLineShapeSnapshot(1, "CHUANHOA2_TN_LINE_P1", 9, paragraph.StoryType,
            paragraph.SectionIndex, paragraph.AbsoluteStart, paragraph.Index, paragraph.PageNumber,
            320, 98, 180, 0, 320, 98, 1, 1, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:motto-line-owned", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), new[] { ownedLine }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Fact]
    public void Current_owned_line_with_wrong_geometry_is_still_reported()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var shiftedLine = new LocalLineShapeSnapshot(1, "CHUANHOA2_MOTTO_P1", 9,
            paragraph.StoryType, paragraph.SectionIndex, paragraph.AbsoluteStart, paragraph.Index,
            paragraph.PageNumber, 100, 150, 90, 0, 100, 150, 1, 1, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:motto-line-owned-shifted", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { shiftedLine }, "STATE_ND30");

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");

        Assert.Contains("Đã có Line Shape", finding.CurrentIssue, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("CHUANHOA2_ORG_P4", true, false)]
    [InlineData("CHUANHOA_ORG_P4", false, true)]
    [InlineData("CHUANHOA2_USER_P4", false, false)]
    [InlineData("CHUANHOA_USER_P4", false, false)]
    [InlineData("CH_L_ORG_P4", false, false)]
    [InlineData("Line 4", false, false)]
    public void Line_shape_ownership_requires_an_exact_released_prefix(
        string name, bool current, bool legacy)
    {
        Assert.Equal(current, LineShapeOwnership.IsCurrentOwned(name));
        Assert.Equal(legacy, LineShapeOwnership.IsLegacyOwned(name));
        Assert.Equal(current || legacy, LineShapeOwnership.IsOwnedForParagraph(name, 4));
    }

    [Fact]
    public void Motto_line_must_match_the_rendered_text_width_closely()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var shortLine = Line(1, paragraph, 329, 98, 162);
        var snapshot = new LocalScanSnapshot("sha256:motto-line-short", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), new[] { shortLine }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Fact]
    public void Underline_or_paragraph_border_does_not_satisfy_line_shape_rules()
    {
        var paragraph = new LocalParagraphSnapshot(1, "Độc lập - Tự do - Hạnh phúc", "wdMainTextStory",
            1, 100, "Times New Roman", fontSizePoints: 13, bold: true, alignment: 1,
            role: "nationalMotto", underline: 1, hasBottomBorder: true,
            pageNumber: 1, pageLeftPoints: 320, pageTopPoints: 80, textWidthPoints: 180);
        var snapshot = new LocalScanSnapshot("sha256:not-a-shape", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), Array.Empty<LocalLineShapeSnapshot>(), "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Fact]
    public void Misanchored_dashed_or_wrong_length_lines_do_not_satisfy_shape_rules()
    {
        var paragraph = P(1, "CƠ QUAN BAN HÀNH", "organName", size: 13, bold: true,
            alignment: 1, page: 1, left: 80, top: 80, width: 180);
        var lines = new[]
        {
            Line(1, paragraph, 80, 98, 180, dashStyle: 1),
            Line(2, paragraph, 140, 98, 72, dashStyle: 4),
            new LocalLineShapeSnapshot(3, "Line wrong page", 9, "wdMainTextStory", 1, 100, 1, 2,
                140, 98, 72, 0, 140, 98, 1, 1, true, 1, .75, 0, 1, 1)
        };
        var snapshot = new LocalScanSnapshot("sha256:lines-bad", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), lines, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K2-ORG-LINE");
    }

    [Fact]
    public void Rendered_line_on_same_page_is_accepted_when_word_anchors_it_far_from_component()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var line = new LocalLineShapeSnapshot(1, "Floating line", 9, paragraph.StoryType,
            paragraph.SectionIndex, 5000, 50, 1, 320, 98, 180, 0, 320, 98,
            1, 1, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:floating-line", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), new[] { line }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Fact]
    public void Line_on_following_blank_paragraph_uses_relative_top_when_word_returns_zero_anchor_position()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 184, top: 97, width: 165);
        var line = new LocalLineShapeSnapshot(1, "Blank anchor line", 9, paragraph.StoryType,
            paragraph.SectionIndex, paragraph.AbsoluteStart + paragraph.Text.Length + 1, paragraph.Index + 1,
            1, 184, 18, 165, 0, 184, 18, 1, 2, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:blank-anchor-line", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { line }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Fact]
    public void Slightly_sloped_real_line_is_treated_as_horizontal()
    {
        var paragraph = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 13,
            bold: true, alignment: 1, page: 1, left: 320, top: 80, width: 180);
        var line = new LocalLineShapeSnapshot(1, "Nearly horizontal line", 9, paragraph.StoryType,
            paragraph.SectionIndex, paragraph.AbsoluteStart, paragraph.Index, 1,
            320, 98, 179.96, 3.5, 320, 98, 1, 1, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:sloped-line", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), new[] { line }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Fact]
    public void Existing_line_with_wrong_geometry_is_not_reported_as_missing()
    {
        var paragraph = P(1, "CƠ QUAN BAN HÀNH", "organName", size: 13, bold: true,
            alignment: 1, page: 1, left: 80, top: 80, width: 180);
        var line = Line(1, paragraph, 80, 98, 180);
        var snapshot = new LocalScanSnapshot("sha256:wrong-line-geometry", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { line }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        var finding = Assert.Single(findings, item => item.RuleCode == "ND30-PL1-M2-K2-ORG-LINE");

        Assert.Contains("Đã có Line Shape", finding.CurrentIssue, StringComparison.Ordinal);
        Assert.DoesNotContain("Thiếu", finding.CurrentIssue, StringComparison.Ordinal);
    }

    [Fact]
    public void Visually_balanced_small_line_offset_is_accepted()
    {
        var paragraph = P(1, "CƠ QUAN BAN HÀNH", "organName", size: 13, bold: true,
            alignment: 1, page: 1, left: 80, top: 80, width: 180);
        var line = Line(1, paragraph, 140, 98, 72);
        // Expected centre is 170 pt; the line centre is 176 pt. Compatibility-mode
        // DOC files commonly differ by this amount while remaining visually balanced.
        var snapshot = new LocalScanSnapshot("sha256:balanced-small-offset", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { line }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K2-ORG-LINE");
    }

    [Fact]
    public void Unrelated_line_in_another_column_does_not_satisfy_component_rule()
    {
        var paragraph = P(1, "CƠ QUAN BAN HÀNH", "organName", size: 13, bold: true,
            alignment: 1, page: 1, left: 80, top: 80, width: 180);
        var line = new LocalLineShapeSnapshot(1, "Other column line", 9, paragraph.StoryType,
            paragraph.SectionIndex, 5000, 50, 1, 320, 98, 72, 0, 320, 98,
            1, 1, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:other-column", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(), new[] { line }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K2-ORG-LINE" &&
            item.CurrentIssue.StartsWith("Thiếu", StringComparison.Ordinal));
    }

    [Fact]
    public void Nearby_anchor_in_another_header_column_is_not_associated_with_the_component()
    {
        var paragraph = P(2, "CƠ QUAN BAN HÀNH", "organName", size: 13, bold: true,
            alignment: 1, page: 1, left: 80, top: 80, width: 180);
        var otherColumnLine = new LocalLineShapeSnapshot(1, "Motto line", 9, paragraph.StoryType,
            paragraph.SectionIndex, paragraph.AbsoluteStart + 100, paragraph.Index + 2, 1,
            390, 98, 180, 0, 390, 98, 1, 2, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("sha256:near-anchor-other-column", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { otherColumnLine }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K2-ORG-LINE" &&
            item.CurrentIssue.StartsWith("Thiếu", StringComparison.Ordinal));
    }

    [Fact]
    public void Decision_subject_starting_with_ve_viec_is_named_document_subject_not_official_letter_subject()
    {
        var paragraphs = new[]
        {
            P(1, "QUYẾT ĐỊNH", size: 14, bold: true, alignment: 1),
            P(2, "Về việc phê duyệt Báo cáo kinh tế - kỹ thuật", size: 14, bold: true, alignment: 1)
        };
        var snapshot = new LocalScanSnapshot("sha256:decision-subject", 1, new[] { ValidSection() }, paragraphs,
            Array.Empty<AnnotationProtectedSpan>(), documentTypeCode: LocalDocumentTypeCodes.Decision,
            documentTypeWasSelectedManually: true);

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K5B-STYLE");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K5A-SUBJ");
    }

    [Fact]
    public void Official_letter_subject_rule_requires_actual_official_letter_context()
    {
        var paragraph = P(1, "V/v nâng bậc lương năm 2026", size: 12, bold: false, alignment: 1);
        var decision = new LocalScanSnapshot("sha256:not-letter", 1, new[] { ValidSection() }, new[] { paragraph },
            Array.Empty<AnnotationProtectedSpan>(), documentTypeCode: LocalDocumentTypeCodes.Decision,
            documentTypeWasSelectedManually: true);
        var officialLetter = new LocalScanSnapshot("sha256:letter", 1, new[] { ValidSection() }, new[] { paragraph },
            Array.Empty<AnnotationProtectedSpan>(), documentTypeCode: LocalDocumentTypeCodes.OfficialLetter,
            documentTypeWasSelectedManually: true);

        Assert.DoesNotContain(new LocalDocumentScanner().ScanFormat(decision, Rules()).Findings,
            item => item.RuleCode.StartsWith("ND30-PL1-M2-K5B", StringComparison.Ordinal));
        Assert.DoesNotContain(new LocalDocumentScanner().ScanFormat(officialLetter, Rules()).Findings,
            item => item.RuleCode.StartsWith("ND30-PL1-M2-K5B", StringComparison.Ordinal));
    }

    [Fact]
    public void National_title_accepts_hoa_variant_and_does_not_title_case_viet_nam()
    {
        var paragraph = P(1, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", size: 13, bold: true, alignment: 1);
        var snapshot = new LocalScanSnapshot("sha256:national-title", 1, new[] { ValidSection() }, new[] { paragraph },
            Array.Empty<AnnotationProtectedSpan>());

        var format = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        var spelling = new LocalDocumentScanner().ScanSpelling(snapshot, Rules()).Findings;

        Assert.DoesNotContain(format, item => item.RuleCode == "ND30-PL1-M2-K1-QH");
        Assert.DoesNotContain(spelling, item => item.RuleCode == "ND30-PL2-M3-K1C");
        Assert.DoesNotContain(spelling, item => item.Expected.Contains("HÒA", StringComparison.Ordinal));
    }

    [Fact]
    public void Legal_basis_with_ngay_is_not_classified_as_place_date()
    {
        var paragraph = P(1, "Căn cứ Luật Xây dựng ngày 10 tháng 12 năm 2025;", size: 14, italic: true,
            alignment: 3, indent: 30);
        var snapshot = new LocalScanSnapshot("sha256:legal-basis", 1, new[] { ValidSection() }, new[] { paragraph },
            Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode.StartsWith("ND30-PL1-M2-K4", StringComparison.Ordinal));
    }

    [Fact]
    public void A_single_legal_basis_block_must_end_with_a_period()
    {
        var paragraph = P(1, "Căn cứ Hợp đồng số 129/2026/HĐTV/CĐCS-BMC;", "legalBasis",
            size: 14, italic: true, alignment: 3);
        var snapshot = new LocalScanSnapshot("sha256:single-legal-basis", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K6A-PUNCT");

        Assert.Equal(";", finding.Anchor.ExpectedText);
        Assert.Contains("dấu chấm (.)", finding.Expected, StringComparison.Ordinal);
    }

    [Fact]
    public void Legal_basis_punctuation_is_resolved_per_consecutive_block()
    {
        var paragraphs = new[]
        {
            P(1, "Căn cứ Luật Đấu thầu.", "legalBasis", size: 14, italic: true, alignment: 3),
            P(2, "Căn cứ Nghị định số 30/2020/NĐ-CP;", "legalBasis", size: 14, italic: true, alignment: 3),
            P(3, "QUYẾT ĐỊNH", "structuralTitle", size: 14, bold: true, alignment: 1),
            P(4, "Nội dung giải trình của tổ thẩm định.", size: 14),
            P(5, "Căn cứ Hợp đồng số 129/2026/HĐTV/CĐCS-BMC;", "legalBasis",
                size: 14, italic: true, alignment: 3),
            P(6, "b) Thành phần đơn vị thẩm định", size: 14)
        };
        var snapshot = new LocalScanSnapshot("sha256:legal-basis-blocks", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings
            .Where(item => item.RuleCode == "ND30-PL1-M2-K6A-PUNCT")
            .ToArray();

        Assert.Equal(3, findings.Length);
        Assert.Contains(findings, item => item.Anchor.ParagraphIndex == 1 &&
            item.Expected.Contains("chấm phẩy (;)", StringComparison.Ordinal));
        Assert.Contains(findings, item => item.Anchor.ParagraphIndex == 2 &&
            item.Expected.Contains("dấu chấm (.)", StringComparison.Ordinal));
        Assert.Contains(findings, item => item.Anchor.ParagraphIndex == 5 &&
            item.Expected.Contains("dấu chấm (.)", StringComparison.Ordinal));
    }

    [Fact]
    public void Body_sentence_starting_with_căn_cứ_is_not_checked_as_a_legal_basis()
    {
        var paragraphs = new[]
        {
            P(1, "BÁO CÁO", size: 14, bold: true, alignment: 1),
            P(2, "Về kết quả thẩm định hồ sơ mời thầu", size: 14, bold: true, alignment: 1),
            P(3, "1. NỘI DUNG THẨM ĐỊNH", size: 14, bold: true),
            P(4, "a) Ý kiến thẩm định về cơ sở pháp lý:", size: 14),
            P(5, "Căn cứ các tài liệu được cung cấp, kết quả thẩm định được tổng hợp tại Bảng số 01.",
                size: 14, italic: true, alignment: 3)
        };
        var snapshot = new LocalScanSnapshot("sha256:body-căn-cứ", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.Anchor.ParagraphIndex == 5 &&
            (item.RuleCode == "ND30-PL1-M2-K6A-PUNCT" || item.RuleCode == "ND30-PL1-M2-K6A-STYLE"));
    }

    [Fact]
    public void Narrative_căn_cứ_after_subject_is_not_checked_as_a_legal_basis()
    {
        var paragraphs = new[]
        {
            P(1, "BÁO CÁO", size: 14, bold: true, alignment: 1),
            P(2, "Về kết quả thẩm định hồ sơ mời thầu", size: 14, bold: true, alignment: 1),
            P(3, "Căn cứ các tài liệu được cung cấp, kết quả thẩm định được tổng hợp tại Bảng số 01.",
                size: 14, italic: true, alignment: 3)
        };
        var snapshot = new LocalScanSnapshot("sha256:narrative-căn-cứ-after-subject", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.Anchor.ParagraphIndex == 3 &&
            item.RuleCode.StartsWith("ND30-PL1-M2-K6A", StringComparison.Ordinal));
    }

    [Fact]
    public void Point_style_finding_reports_actual_format_and_an_explicit_fix()
    {
        var point = P(1, "b) Ý kiến thẩm định về nội dung không tuân thủ", size: 14,
            bold: true, italic: true);
        var snapshot = new LocalScanSnapshot("sha256:point-style", 1,
            new[] { ValidSection() }, new[] { point }, Array.Empty<AnnotationProtectedSpan>());

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K6D-POINT");

        Assert.Equal("Điểm sai kiểu chữ: đang in đậm, đang in nghiêng.", finding.CurrentIssue);
        Assert.Equal("Bỏ in đậm và in nghiêng cho toàn bộ điểm; trình bày bằng chữ đứng, không đậm.",
            finding.Expected);
    }

    [Fact]
    public void Point_style_finding_detects_mixed_bold_or_italic_formatting()
    {
        var point = P(1, "c) Cách thức làm việc", size: 14, bold: null, italic: null);
        var snapshot = new LocalScanSnapshot("sha256:mixed-point-style", 1,
            new[] { ValidSection() }, new[] { point }, Array.Empty<AnnotationProtectedSpan>());

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K6D-POINT");

        Assert.Contains("đậm/không đậm chưa đồng nhất", finding.CurrentIssue, StringComparison.Ordinal);
        Assert.Contains("nghiêng/đứng chưa đồng nhất", finding.CurrentIssue, StringComparison.Ordinal);
    }

    [Fact]
    public void Place_and_issued_date_is_centered_under_the_national_identity_block()
    {
        var paragraph = P(1, "Hà Nội, ngày 02 tháng 09 năm 2026", "placeAndIssuedDate",
            size: 13, italic: true, alignment: 1, indent: 0);
        var snapshot = new LocalScanSnapshot("sha256:place-date-center", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K4-STYLE");
    }

    [Theory]
    [InlineData(12, 13, 13)]
    [InlineData(13, 14, 14)]
    public void Nd30_header_accepts_only_the_two_consistent_font_size_tiers(
        double nationalSize, double mottoSize, double dateSize)
    {
        var paragraphs = new[]
        {
            P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "nationalTitle", size: nationalSize,
                bold: true, alignment: 1),
            P(2, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: mottoSize,
                bold: true, alignment: 1, page: 1, left: 300, top: 80, width: 180),
            P(3, "Hà Nội, ngày 02 tháng 09 năm 2026", "placeAndIssuedDate", size: dateSize,
                italic: true, alignment: 1, indent: 0)
        };
        var snapshot = new LocalScanSnapshot("sha256:header-tier-ok", 1, new[] { ValidSection() },
            paragraphs, Array.Empty<AnnotationProtectedSpan>(),
            new[] { Line(1, paragraphs[1], 300, 98, 180) }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-MV-CT1" &&
            item.CurrentIssue.Contains("bậc cỡ chữ", StringComparison.Ordinal));
    }

    [Fact]
    public void Nd30_header_reports_cross_tier_sizes_even_when_each_size_is_individually_legal()
    {
        var paragraphs = new[]
        {
            P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "nationalTitle", size: 12,
                bold: true, alignment: 1),
            P(2, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", size: 14,
                bold: true, alignment: 1, page: 1, left: 300, top: 80, width: 180),
            P(3, "Hà Nội, ngày 02 tháng 09 năm 2026", "placeAndIssuedDate", size: 14,
                italic: true, alignment: 1, indent: 0)
        };
        var snapshot = new LocalScanSnapshot("sha256:header-tier-bad", 1, new[] { ValidSection() },
            paragraphs, Array.Empty<AnnotationProtectedSpan>(),
            new[] { Line(1, paragraphs[1], 300, 98, 180) }, "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-MV-CT1" &&
            item.Anchor.ParagraphIndex == 2 && item.Expected.Contains("Quốc hiệu cỡ 12, Tiêu ngữ cỡ 13", StringComparison.Ordinal));
        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-MV-CT1" &&
            item.Anchor.ParagraphIndex == 3);
    }

    [Fact]
    public void Appendix_models_accept_arabic_or_roman_ordinals_and_model_22_signature_information()
    {
        var section = new LocalSectionSnapshot(1, ValidSection().PageWidthPoints,
            ValidSection().PageHeightPoints, ValidSection().TopMarginPoints,
            ValidSection().BottomMarginPoints, ValidSection().LeftMarginPoints,
            ValidSection().RightMarginPoints, false, true, true, 1, 1);
        var paragraphs = new[]
        {
            P(1, "Ban hành kèm theo 02 phụ lục.", section: 1),
            P(2, "Phụ lục 1", "appendixLabel", size: 14, bold: true, alignment: 1, section: 1),
            P(3, "DANH MỤC BIỂU MẪU", size: 14, bold: true, alignment: 1, section: 1),
            P(4, "ÁP DỤNG NĂM 2026", size: 13, bold: true, alignment: 1, section: 1),
            P(5, "(Kèm theo Văn bản số 01/QĐ-A ngày 02 tháng 09 năm 2026 của Cơ quan A)",
                size: 14, italic: true, alignment: 1, section: 1),
            P(6, "Số: 02/QĐ-A; ngày 02/09/2026; 08:30:15", size: 10, alignment: 2, section: 1),
            P(7, "Phụ lục II", "appendixLabel", size: 14, bold: true, alignment: 1, section: 1),
            P(8, "DỮ LIỆU ĐIỆN TỬ", size: 13, bold: true, alignment: 1, section: 1),
            P(9, "(Kèm theo Văn bản số ... ngày ... tháng ... năm ... của Cơ quan A)",
                size: 13, italic: true, alignment: 1, section: 1)
        };
        var snapshot = new LocalScanSnapshot("sha256:appendix-models", 1, new[] { section },
            paragraphs, Array.Empty<AnnotationProtectedSpan>(), regimeCode: "STATE_ND30");

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M3-K1A-NUM");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M3-K1B");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M3-K1C");
    }

    [Fact]
    public void Style_finding_describes_current_values_and_the_complete_required_format()
    {
        var paragraph = P(1, "CƠ QUAN CHỦ QUẢN", "superiorOrganName", "Arial", 10,
            bold: true, italic: true, alignment: 0);
        var snapshot = new LocalScanSnapshot("sha256:explicit-style-guidance", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var finding = Assert.Single(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            item => item.RuleCode == "ND30-PL1-M2-K2-SUP");

        Assert.Equal(
            "Tên cơ quan chủ quản sai thể thức: đang dùng phông Arial, đang dùng cỡ chữ 10, đang in đậm, đang in nghiêng, đang căn trái.",
            finding.CurrentIssue);
        Assert.Equal(
            "Định dạng Tên cơ quan chủ quản: phông Times New Roman, cỡ chữ 12–13, không in đậm, không in nghiêng, căn giữa.",
            finding.Expected);
    }

    [Fact]
    public void Abbreviated_citation_date_in_legal_basis_is_accepted()
    {
        var paragraph = P(1, "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05/03/2020 của Chính phủ;",
            "legalBasis", size: 14, italic: true, alignment: 3);
        var snapshot = new LocalScanSnapshot("sha256:citation-short-date", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-DATE");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-CITE");
    }

    [Fact]
    public void Short_date_without_ngay_is_reported_by_spelling_but_still_counts_for_citation_completeness()
    {
        var paragraph = P(1, "Căn cứ Nghị định số 30/2020/NĐ-CP 05/03/2020 của Chính phủ;",
            "legalBasis", size: 14, italic: true, alignment: 3);
        var snapshot = new LocalScanSnapshot("sha256:citation-bare-short-date", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        var spelling = new LocalDocumentScanner().ScanSpelling(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-DATE");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-CITE");
        var missingNgay = Assert.Single(spelling, item => item.RuleCode == "ND30-PL1-M2-K6B-DATE");
        Assert.Equal("05/03/2020", missingNgay.Anchor.ExpectedText);
        Assert.Equal("Ngày viết dạng số nhưng thiếu từ “ngày” phía trước.", missingNgay.CurrentIssue);
        Assert.Equal("Thêm từ “ngày” trước ngày tháng, ví dụ: ngày 05/03/2020.", missingNgay.Expected);
    }

    [Fact]
    public void Short_date_with_ngay_is_accepted_inside_and_outside_legal_basis()
    {
        var paragraph = P(1, "Thực hiện Nghị định số 30/2020/NĐ-CP ngày 05/03/2020 của Chính phủ.",
            size: 14, italic: false, alignment: 3);
        var snapshot = new LocalScanSnapshot("sha256:citation-short-date-outside-basis", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        var spelling = new LocalDocumentScanner().ScanSpelling(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-DATE");
        Assert.DoesNotContain(spelling, item => item.RuleCode == "ND30-PL1-M2-K6B-DATE");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-CITE");
    }

    [Fact]
    public void Bare_short_date_is_not_reported_in_body_table_or_parenthesized_citation()
    {
        var paragraphs = new[]
        {
            P(1, "Thời hạn thực hiện: 01/07/2025."),
            new LocalParagraphSnapshot(2, "01/07/2025", "wdMainTextStory", 1, 200,
                "Times New Roman", tableIndex: 1, rowIndex: 1, cellIndex: 1,
                isInTable: true, fontSizePoints: 13),
            P(3, "Căn cứ Nghị định số 30/2020/NĐ-CP (05/03/2020) của Chính phủ;",
                "legalBasis", size: 14, italic: true, alignment: 3)
        };
        var snapshot = new LocalScanSnapshot("sha256:context-aware-date", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanSpelling(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-DATE");
    }

    [Theory]
    [InlineData("Bộ luật số 45/2019/QH14")]
    [InlineData("Lệnh số 01/2025/L-CTN")]
    [InlineData("Công văn số 12/BNV-TCBC")]
    [InlineData("Tờ trình số 09/TTr-CP")]
    [InlineData("Nghị quyết liên tịch số 01/2025/NQLT")]
    [InlineData("Thông tư liên tịch số 02/2025/TTLT")]
    public void Shared_citation_vocabulary_recognizes_required_document_types(string citation)
    {
        var paragraph = P(1, citation + " ngày 01 tháng 7 năm 2025 của Chính phủ.",
            size: 14, italic: true, alignment: 3);
        var snapshot = new LocalScanSnapshot("sha256:citation-vocabulary", 1,
            new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-SO");
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6B-CITE");
    }

    [Fact]
    public void Citation_completeness_is_checked_inside_each_citation_segment()
    {
        var paragraph = P(1,
            "Nghị định số 30/2020/NĐ-CP ngày 05/03/2020 của Chính phủ; Thông tư số 01/2026/TT-BNV ngày 01/06/2026;",
            size: 14, italic: true, alignment: 3);
        var snapshot = new LocalScanSnapshot("sha256:citation-segments", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        var completeness = findings.Where(item => item.RuleCode == "ND30-PL1-M2-K6B-CITE").ToArray();
        Assert.Single(completeness);
        Assert.Equal("Thông tư số 01/2026/TT-BNV", completeness[0].Anchor.ExpectedText,
            ignoreCase: true, ignoreLineEndingDifferences: false, ignoreWhiteSpaceDifferences: false);
    }

    [Fact]
    public void Body_after_last_legal_basis_is_not_formatted_as_another_legal_basis()
    {
        var paragraphs = new[]
        {
            P(1, "QUYẾT ĐỊNH", size: 14, bold: true, alignment: 1),
            P(2, "Về việc phê duyệt hồ sơ", size: 14, bold: true, alignment: 1),
            P(3, "Căn cứ Nghị định số 30/2020/NĐ-CP;", size: 14, italic: true, alignment: 3),
            P(4, "Căn cứ đề nghị của cơ quan chuyên môn.", size: 14, italic: true, alignment: 3),
            P(5, "Nội dung quyết định này có hiệu lực kể từ ngày ký và được tổ chức thực hiện.",
                size: 14, bold: false, italic: false, alignment: 3, indent: 10 * 72d / 25.4d, after: 6)
        };
        var snapshot = new LocalScanSnapshot("sha256:legal-body", 1, new[] { ValidSection() }, paragraphs,
            Array.Empty<AnnotationProtectedSpan>(), documentTypeCode: LocalDocumentTypeCodes.Decision,
            documentTypeWasSelectedManually: true);

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("legalBasis", roles[3]);
        Assert.Equal("legalBasis", roles[4]);
        Assert.False(roles.ContainsKey(5));
    }

    [Fact]
    public void Article_indent_uses_millimeters_converted_to_points()
    {
        const double pointsPerMillimeter = 72d / 25.4d;
        var valid = P(1, "Điều 1. Phạm vi điều chỉnh", bold: null,
            indent: 10d * pointsPerMillimeter);
        var invalid = P(1, "Điều 1. Phạm vi điều chỉnh", bold: null, indent: 10d);
        var scanner = new LocalDocumentScanner();

        var validFindings = scanner.ScanFormat(new LocalScanSnapshot(
            "sha256:article-indent-valid", 1, new[] { ValidSection() },
            new[] { valid }, Array.Empty<AnnotationProtectedSpan>()), Rules()).Findings;
        var invalidFindings = scanner.ScanFormat(new LocalScanSnapshot(
            "sha256:article-indent-invalid", 1, new[] { ValidSection() },
            new[] { invalid }, Array.Empty<AnnotationProtectedSpan>()), Rules()).Findings;

        Assert.DoesNotContain(validFindings, item =>
            item.RuleCode == "ND30-PL1-M2-K6D-ARTICLE" &&
            item.CurrentIssue.Contains("thụt đầu dòng sai", StringComparison.Ordinal));
        Assert.Contains(invalidFindings, item =>
            item.RuleCode == "ND30-PL1-M2-K6D-ARTICLE" &&
            item.CurrentIssue.Contains("thụt đầu dòng sai", StringComparison.Ordinal));
    }

    [Fact]
    public void Listing_colon_intro_followed_by_clauses_is_not_flagged_as_invalid_content_end()
    {
        var paragraphs = new[]
        {
            P(1, "QUYẾT ĐỊNH", "typeName", size: 14, bold: true, alignment: 1),
            P(2, "Về việc phê duyệt kế hoạch lựa chọn nhà thầu", "subject", size: 14, bold: true, alignment: 1),
            P(3, "4. Giải pháp và phương pháp luận:", size: 13, bold: true),
            P(4, "Nhà thầu chuẩn bị đề xuất giải pháp, phương pháp luận tổng quát thực hiện dịch vụ theo các nội dung quy định tại Chương này, gồm các phần như sau:", size: 13),
            P(5, "1. Giải pháp và phương pháp luận;", size: 13),
            P(6, "2. Kế hoạch công tác.", size: 13),
            P(7, "5. Quy định về kiểm tra, nghiệm thu sản phẩm: Quy định cụ thể trong hợp đồng.", size: 13),
            P(8, "Nơi nhận:", "recipientLabel", size: 12, bold: true),
            P(9, "- Như trên;", "recipientList", size: 11),
            P(10, "- Lưu: VT.", "recipientList", size: 11)
        };
        var snapshot = new LocalScanSnapshot("sha256:listing-colon", 1, new[] { ValidSection() }, paragraphs,
            Array.Empty<AnnotationProtectedSpan>(), documentTypeCode: LocalDocumentTypeCodes.Decision,
            documentTypeWasSelectedManually: true);

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6E-DOTSLASH");
    }

    [Fact]
    public void Trailing_blank_paragraph_before_recipients_does_not_produce_invalid_dotslash_anchor()
    {
        var paragraphs = new[]
        {
            P(1, "TỜ TRÌNH", "typeName", size: 14, bold: true, alignment: 1),
            P(2, "Về việc mua sắm trang thiết bị", "subject", size: 14, bold: true, alignment: 1),
            P(3, "Nội dung tờ trình kết thúc bằng dấu chấm.", size: 13),
            P(4, "   \r", size: 13),
            P(5, "Nơi nhận:", "recipientLabel", size: 12, bold: true),
            P(6, "- Như trên;", "recipientList", size: 11),
            P(7, "- Lưu: VT.", "recipientList", size: 11)
        };
        var snapshot = new LocalScanSnapshot("sha256:trailing-blank", 1, new[] { ValidSection() }, paragraphs,
            Array.Empty<AnnotationProtectedSpan>(), documentTypeCode: "TO_TRINH",
            documentTypeWasSelectedManually: true);

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K6E-DOTSLASH");
    }

    [Fact]
    public void Scanner_checks_each_embedded_document_against_its_own_type_and_lines()
    {
        var paragraphs = new[]
        {
            P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size: 12, bold: true, alignment: 1, page: 1),
            P(2, "Độc lập - Tự do - Hạnh phúc", size: 13, bold: true, alignment: 1, page: 1),
            P(3, "Số: 01/QĐ-ABC", size: 13, alignment: 1, page: 1),
            P(4, "QUYẾT ĐỊNH", size: 14, bold: true, alignment: 1, page: 1),
            P(5, "Về việc phê duyệt", size: 14, bold: true, alignment: 1, page: 1),
            P(10, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size: 13, bold: true, alignment: 1, page: 2),
            P(11, "Độc lập - Tự do - Hạnh phúc", size: 14, bold: true, alignment: 1, page: 2),
            P(12, "Số: 02/QĐ-ABC", size: 13, alignment: 1, page: 2),
            P(13, "THÔNG BÁO", size: 14, bold: true, alignment: 1, page: 2),
            P(14, "Về việc triển khai", size: 14, bold: true, alignment: 1, page: 2)
        };
        var snapshot = new LocalScanSnapshot("sha256:two-scanner-blocks", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;

        Assert.Contains(findings, item => item.RuleCode == "ND30-PL1-M2-K3-ABBR" &&
            item.Anchor.ParagraphIndex == 12 && item.Expected.Contains("TB", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-M2-K3-ABBR" &&
            item.Anchor.ParagraphIndex == 3);
        Assert.Equal(2, findings.Count(item =>
            item.RuleCode == "ND30-PL1-M2-K1-TN-LINE"));
        Assert.DoesNotContain(findings, item => item.RuleCode == "ND30-PL1-MV-CT1" &&
            (item.Anchor.ParagraphIndex == 1 || item.Anchor.ParagraphIndex == 2 ||
             item.Anchor.ParagraphIndex == 10 || item.Anchor.ParagraphIndex == 11));
    }

    private static LocalScanSnapshot BadFormatSnapshot()
    {
        var paragraphs = new List<LocalParagraphSnapshot>
        {
            P(1, "Cộng hòa xã hội chủ nghĩa Việt Nam", "nationalTitle", "Arial", 10, false, true, 0),
            P(2, "Độc lập–Tự do–Hạnh phúc", "nationalMotto", "Arial", 10, false, true, 0, before: 12),
            P(3, "BỘ CHỦ QUẢN", "superiorOrganName", "Arial", 10, true, true, 0),
            P(4, "CƠ QUAN BAN HÀNH", "organName", "Arial", 10, false, true, 0),
            P(5, "Quyết định", "typeName", "Arial", 10, false, true, 0),
            P(6, "Số 5 qđ abc", "codeNumber", "Times New Roman", 13),
            P(7, "hà nội ngày 1 tháng 2 năm 2026", "placeAndIssuedDate", "Arial", 10, true, false, 0),
            P(8, "Trích yếu văn bản", "subject", "Arial", 10, false, true, 0),
            P(9, "Trích yếu công văn", "officialLetterSubject", "Arial", 10, true, true, 0),
            P(10, "Căn cứ Nghị định 12/2020 ngày 1/2/2020", "legalBasis", "Arial", 10, true, false, 0),
            P(11, "Điều 1", font: "Times New Roman", size: 13, bold: false, indent: 0),
            P(12, "1. Nội dung khoản", font: "Times New Roman", size: 13, italic: true),
            P(13, "a) Nội dung điểm", font: "Times New Roman", size: 13, bold: true),
            P(14, "c) Nội dung điểm tiếp theo", font: "Times New Roman", size: 13),
            P(15, "Đây là phần nội dung chính đủ dài nhưng kết thúc chưa có dấu", font: "Arial", size: 11,
                alignment: 0, indent: 0, after: 0, lineSpacing: 30, color: 255, lineRule: 5),
            P(16, "kt. GIÁM ĐỐC", "signerAuthority", "Arial", 10, false, true, 0),
            P(17, "Kính gửi", "recipientSalutation", "Times New Roman", 13),
            P(18, "Kính gửi: Cơ quan A", "recipientSalutationInline", "Times New Roman", 13),
            P(19, "- Cơ quan A,", "recipientSalutationList", "Times New Roman", 13),
            P(20, "Nơi nhận", "recipientLabel", "Arial", 14, false, false, 2),
            P(21, "Cơ quan B", "recipientList", "Arial", 14),
            P(22, "- Không phải dòng lưu;", "recipientList", "Arial", 14),
            P(23, "Phụ lục A", "appendixLabel", "Arial", 10, false, true, 0, section: 1),
            P(24, "TÊN PHỤ LỤC", font: "Arial", size: 10, bold: false, italic: true, alignment: 0, section: 1),
            P(25, "(Kèm theo Quyết định)", font: "Arial", size: 10, bold: true, italic: false, alignment: 0, section: 1),
            P(26, "Phụ lục 2", "appendixLabel", "Arial", 10, false, true, 0, section: 1)
            ,P(27, "ĐẢNG CỘNG SẢN VIỆT NAM", "partyTitle", "Arial", 10, false, false, 1),
            P(28, "Thực hiện Nghị định số 30/2020/NĐ-CP ngày 05/03/2020 của Chính phủ.",
                font: "Times New Roman", size: 13)
        };
        return new LocalScanSnapshot("sha256:bad-format", 1,
            new[] { new LocalSectionSnapshot(1, 700, 900, 10, 10, 10, 10, true, false, false, 1, 0) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());
    }

    private static LocalScanSnapshot LayoutOnlySnapshot()
    {
        return new LocalScanSnapshot("sha256:layout", 1, new[] { ValidSection() },
            new[] { P(1, "Kính gửi", "recipientSalutation", "Times New Roman", 13) },
            Array.Empty<AnnotationProtectedSpan>());
    }

    private static LocalScanSnapshot BadSpellingSnapshot()
    {
        var text = "Căn cứ Nghị định số 30/2020/NĐ-CP 05/03/2020 của Chính phủ; nội dung đầu. câu sau  sai , ông nguyễn văn an doof. hà nội, việt nam, sông hồng, tây nguyên, bộ nội vụ, ban chấp hành trung ương, ngày quốc khánh, giáp thìn, phường iv, chương I và Điều 2 Khoản 1 Điểm a), sát nhập.\u200B";
        return new LocalScanSnapshot("sha256:bad-spelling", 1, new[] { ValidSection() },
            new[] { P(1, text, "legalBasis", font: "Times New Roman", size: 13) }, Array.Empty<AnnotationProtectedSpan>());
    }

    private static LocalSectionSnapshot ValidSection()
    {
        const double pt = 72d / 25.4d;
        return new LocalSectionSnapshot(1, 210 * pt, 297 * pt, 20 * pt, 20 * pt, 30 * pt, 15 * pt, false, true);
    }

    [Fact]
    public void Combined_national_header_reports_the_missing_independent_motto_line()
    {
        var paragraph = P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\vĐộc lập - Tự do - Hạnh phúc\v---------------");
        var snapshot = new LocalScanSnapshot("sha256:combined-header", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());
        var scan = new LocalDocumentScanner().ScanFormat(snapshot, Rules());
        Assert.Contains(scan.Findings, f => f.RuleCode == "ND30-PL1-M2-K1-TN-LINE" &&
            f.Anchor.ParagraphIndex == 1);
    }

    [Fact]
    public void Long_decision_subject_is_centered_not_body_text()
    {
        var text = "Về việc phê duyệt " + string.Join(" ", Enumerable.Repeat("nhiệm vụ của dự án", 60));
        var snapshot = new LocalScanSnapshot("long-subject", 1, new[] { ValidSection() },
            new[] { P(1, "QUYẾT ĐỊNH", bold: true, alignment: 1),
                P(2, text, bold: true, alignment: 1, indent: 0, after: 0),
                P(3, "Căn cứ Luật Đấu thầu;", italic: true) }, Array.Empty<AnnotationProtectedSpan>());
        Assert.Equal("subject", new DocumentRoleDetector().Detect(snapshot)[2]);
        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        Assert.DoesNotContain(findings, f => f.Anchor.ParagraphIndex == 2 &&
            new[] { "ND30-PL1-M2-K6E-ALIGN", "ND30-PL1-M2-K6E-INDENT", "ND30-PL1-M2-K6E-SPACEAFTER" }.Contains(f.RuleCode));
    }

    [Fact]
    public void Two_valid_motto_lines_are_reported_instead_of_accepting_the_first()
    {
        var motto = P(1, "Độc lập - Tự do - Hạnh phúc", "nationalMotto", alignment: 1,
            page: 1, left: 300, top: 70, width: 170);
        var snapshot = new LocalScanSnapshot("sha256:duplicate-lines", 1, new[] { ValidSection() },
            new[] { motto }, Array.Empty<AnnotationProtectedSpan>(),
            new[] { Line(1, motto, 300, 88, 170), Line(2, motto, 300, 91, 170) });
        Assert.Contains(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            f => f.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
    }

    [Theory]
    [InlineData(15, -5, true)]
    [InlineData(25, -5, true)]
    [InlineData(15, 0, true)]
    public void Scanner_checks_list_geometry_not_only_negative_first_line(double left, double first, bool invalid)
    {
        const double points = 72d / 25.4d;
        var paragraph = new LocalParagraphSnapshot(1, "- Nội dung cam kết thực hiện nhiệm vụ", "wdMainTextStory", 1, 0,
            "Times New Roman", fontSizePoints: 14, alignment: 3, firstLineIndentPoints: first * points,
            leftIndentPoints: left * points);
        var snapshot = new LocalScanSnapshot("sha256:hanging", 1, new[] { ValidSection() },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());
        var scan = new LocalDocumentScanner().ScanFormat(snapshot, Rules());
        Assert.Equal(invalid, scan.Findings.Any(f => f.RuleCode == "ND30-PL1-M2-K6E-INDENT"));
    }

    [Fact]
    public void Automatic_recipient_list_is_checked_without_adding_bullet_to_source_text()
    {
        var paragraphs = new[] {
            P(1, "Nơi nhận:", size: 12, bold: true, italic: true, alignment: 0),
            new LocalParagraphSnapshot(2, "Như Điều 4;", "wdMainTextStory", 1, 100,
                "Times New Roman", fontSizePoints: 14, italic: true, alignment: 3,
                firstLineIndentPoints: 28, leftIndentPoints: 28, listMarker: "-"),
            new LocalParagraphSnapshot(3, "Lưu VT.", "wdMainTextStory", 1, 200,
                "Times New Roman", fontSizePoints: 14, italic: true, alignment: 3, listMarker: "-") };
        var snapshot = new LocalScanSnapshot("auto-recipients", 1, new[] { ValidSection() },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());
        Assert.Equal("recipientList", new DocumentRoleDetector().Detect(snapshot)[2]);
        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        Assert.Contains(findings, f => f.RuleCode == "ND30-PL1-M2-K9B-LIST" && f.Anchor.ParagraphIndex == 2);
        Assert.DoesNotContain(findings, f => f.RuleCode == "ND30-PL1-M2-K9B-LUU");
        Assert.Equal("Như Điều 4;", paragraphs[1].Text);
    }

    [Fact]
    public void Proper_auto_list_archive_line_accepts_comma_and_final_period()
    {
        var snapshot = new LocalScanSnapshot("valid-recipients", 1, new[] { ValidSection() },
            new[] { P(1, "Nơi nhận:", size:12, bold:true, italic:true, alignment:0),
                new LocalParagraphSnapshot(2, "Lưu: VT, HC (02).", "wdMainTextStory", 1, 100,
                    "Times New Roman", fontSizePoints:11, italic:false, bold:false, alignment:0,
                    firstLineIndentPoints:0, leftIndentPoints:0, listMarker:"-") },
            Array.Empty<AnnotationProtectedSpan>());
        Assert.DoesNotContain(new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings,
            f => f.RuleCode == "ND30-PL1-M2-K9B-LUU");
    }

    [Theory]
    [InlineData(1, true)]
    [InlineData(12, false)]
    public void Motto_line_on_cover_cannot_satisfy_later_physical_page(int anchorPhysicalPage, bool missing)
    {
        var motto = P(510, "Độc lập - Tự do - Hạnh phúc", "nationalMotto",
            size:13, bold:true, alignment:1, section:3, page:12, left:300, top:70, width:170);
        var line = new LocalLineShapeSnapshot(1, "CHUANHOA2_MOTTO_P510", 9, "wdMainTextStory",
            3, motto.AbsoluteStart, 510, anchorPhysicalPage, 300, 88, 170, 0, 300, 88,
            1, 1, true, 1, .75, 0, 1, 1);
        var snapshot = new LocalScanSnapshot("physical-page-match", 1, Array.Empty<LocalSectionSnapshot>(),
            new[] { motto }, Array.Empty<AnnotationProtectedSpan>(), new[] { line });
        var findings = new LocalDocumentScanner().ScanFormat(snapshot, Rules()).Findings;
        Assert.Equal(missing, findings.Any(f => f.RuleCode == "ND30-PL1-M2-K1-TN-LINE"));
    }

    private static LocalParagraphSnapshot P(int index, string text, string role = "Unknown", string font = "Times New Roman",
        double? size = 13, bool? bold = false, bool? italic = false, int? alignment = 3, double? indent = 30,
        double? before = 0, double? after = 6, double? lineSpacing = 12, int? color = 0, int section = 1,
        int page = 0, double? left = null, double? top = null, double? width = null, int lineRule = 0)
    {
        return new LocalParagraphSnapshot(index, text, "wdMainTextStory", section, index * 100, font,
            fontSizePoints: size, bold: bold, italic: italic, alignment: alignment,
            firstLineIndentPoints: indent, spaceBeforePoints: before, spaceAfterPoints: after,
            role: role, fontColor: color, lineSpacingPoints: lineSpacing, lineSpacingRule: lineRule,
            pageNumber: page, pageLeftPoints: left, pageTopPoints: top, textWidthPoints: width);
    }

    private static LocalLineShapeSnapshot Line(int index, LocalParagraphSnapshot paragraph, double left,
        double top, double width, int dashStyle = 1)
    {
        return new LocalLineShapeSnapshot(index, "Line " + index, 9, paragraph.StoryType,
            paragraph.SectionIndex, paragraph.AbsoluteStart, paragraph.Index, paragraph.PageNumber,
            left, top, width, 0, left, top, 1, 1, true, dashStyle, .75, 0, 1, 1);
    }

    private static LocalRulePack Rules(IReadOnlyList<string>? lexicon = null)
    {
        return new LocalRulePack("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
            210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
            new[] { new TextCorrectionRule("sát nhập", "sáp nhập") },
            new[] { new TelexRule(@"\bdoof\b", "đồ") }, new[] { '\u200B' },
            capitalizations: new[]
            {
                new CapitalizationRule("administrative", "Hà Nội"),
                new CapitalizationRule("geographic", "Việt Nam"),
                new CapitalizationRule("terrain", "sông Hồng"),
                new CapitalizationRule("region", "Tây Nguyên"),
                new CapitalizationRule("organ", "Bộ Nội vụ"),
                new CapitalizationRule("specialOrgan", "Ban Chấp hành Trung ương"),
                new CapitalizationRule("holiday", "Ngày Quốc khánh"),
                new CapitalizationRule("lunarYear", "Giáp Thìn")
            },
            documentTypeAbbreviations: new[]
            {
                new DocumentTypeAbbreviationRule("Quyết định", "QĐ"),
                new DocumentTypeAbbreviationRule("Thông báo", "TB")
            },
            lexicon: lexicon);
    }
}
