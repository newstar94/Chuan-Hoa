using ChuanHoa.Application;
using ChuanHoa.Application.Scanning;
using ChuanHoa.Contracts;
using ChuanHoa.Domain.Common;

namespace ChuanHoa.Application.Tests;

public sealed class TechnicalDocumentScannerTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");

    [Fact]
    public void Valid_technical_baseline_returns_no_findings()
    {
        var result = Scanner().Scan(Request(Snapshot()));

        Assert.Empty(result.Findings);
        Assert.Equal("sha256:document", result.DocumentFingerprint);
        Assert.Equal(1, result.Revision);
        Assert.Equal(TechnicalDocumentScanner.RuleReleaseId, result.RuleReleaseId);
        Assert.Equal(Now, result.CompletedAtUtc);
    }

    [Fact]
    public void Detects_exact_whitespace_spans_without_marking_another_occurrence()
    {
        var paragraph = Paragraph("Đây  là lỗi , cần tô đúng.", fontName: "Times New Roman");
        var result = Scanner().Scan(Request(Snapshot(paragraphs: new[] { paragraph })));

        var repeated = Assert.Single(result.Findings, item => item.RuleId == "TECH-WHITESPACE-REPEATED");
        Assert.NotNull(repeated.Anchor);
        Assert.Equal(FindingAnchorKind.TextSpan, repeated.Anchor.Kind);
        Assert.Equal(3, repeated.Anchor.StartOffset);
        Assert.Equal(2, repeated.Anchor.Length);
        Assert.Equal("  ", repeated.Anchor.ExpectedText);

        var beforePunctuation = Assert.Single(
            result.Findings,
            item => item.RuleId == "TECH-WHITESPACE-BEFORE-PUNCTUATION");
        Assert.Equal(11, beforePunctuation.Anchor!.StartOffset);
        Assert.Equal(" ", beforePunctuation.Anchor.ExpectedText);
        Assert.All(result.Findings, item => Assert.Equal(RiskTier.ReportOnly, item.RiskTier));
    }

    [Fact]
    public void Detects_non_a4_margin_and_font_with_section_and_paragraph_anchors()
    {
        var section = new SectionSnapshot(1, 700, 900, 10, 10, 10, 10, false);
        var paragraph = Paragraph("Nội dung", fontName: "Arial");

        var result = Scanner().Scan(Request(Snapshot(sections: new[] { section }, paragraphs: new[] { paragraph })));

        Assert.Contains(result.Findings, item =>
            item.RuleId == "TECH-PAGE-A4" && item.Anchor?.Kind == FindingAnchorKind.Section);
        Assert.Contains(result.Findings, item =>
            item.RuleId == "TECH-ND30-MARGINS" && item.Anchor?.Kind == FindingAnchorKind.Section);
        Assert.Contains(result.Findings, item =>
            item.RuleId == "TECH-FONT-TNR" && item.Anchor?.Kind == FindingAnchorKind.Paragraph &&
            item.ParagraphIndex == paragraph.Index);
    }

    [Theory]
    [InlineData("chuanhoa.scan-request.v2", ".docx", "SCAN_SCHEMA_UNSUPPORTED")]
    [InlineData("chuanhoa.scan-request.v1", ".pdf", "DOCUMENT_FORMAT_UNSUPPORTED")]
    public void Rejects_unknown_schema_and_unsupported_format(
        string schema,
        string format,
        string expectedCode)
    {
        var error = Assert.Throws<DomainException>(() =>
            Scanner().Scan(new ScanRequest(schema, ScanLane.Format, Snapshot(fileFormat: format))));

        Assert.Equal(expectedCode, error.Code);
    }

    [Fact]
    public void Spelling_lane_stays_fail_closed_until_verified_checker_exists()
    {
        var error = Assert.Throws<DomainException>(() =>
            Scanner().Scan(Request(Snapshot(), ScanLane.Spelling)));

        Assert.Equal("SCAN_LANE_NOT_IMPLEMENTED", error.Code);
    }

    private static TechnicalDocumentScanner Scanner() => new(new StaticClock(Now));

    private static ScanRequest Request(DocumentSnapshot snapshot, ScanLane lane = ScanLane.Format) =>
        new(TechnicalDocumentScanner.RequestSchema, lane, snapshot);

    private static DocumentSnapshot Snapshot(
        string fileFormat = ".docx",
        IReadOnlyList<SectionSnapshot>? sections = null,
        IReadOnlyList<ParagraphSnapshot>? paragraphs = null)
    {
        return new DocumentSnapshot(
            1,
            "sha256:document",
            1,
            RegimeCode.Nd30,
            DocumentTypeCode.OfficialLetter,
            true,
            true,
            new DocumentPreflight(false, false, false, false, fileFormat, true),
            sections ?? new[] { ValidA4Section() },
            paragraphs ?? new[] { Paragraph("Nội dung hợp lệ") },
            Array.Empty<TableSnapshot>(),
            Array.Empty<ProtectedSpanSnapshot>());
    }

    private static SectionSnapshot ValidA4Section()
    {
        const double pointPerMillimeter = 72.0d / 25.4d;
        return new SectionSnapshot(
            1,
            210.0d * pointPerMillimeter,
            297.0d * pointPerMillimeter,
            20.0d * pointPerMillimeter,
            20.0d * pointPerMillimeter,
            30.0d * pointPerMillimeter,
            15.0d * pointPerMillimeter,
            false);
    }

    private static ParagraphSnapshot Paragraph(string text, string fontName = "Times New Roman") =>
        new(
            1,
            text,
            ComponentRole.BodyText,
            1.0d,
            fontName,
            14.0d,
            false,
            false,
            3,
            28.0d,
            0.0d,
            6.0d,
            false,
            "MainTextStory",
            1,
            0);

    private sealed class StaticClock : IClock
    {
        public StaticClock(DateTimeOffset utcNow) => UtcNow = utcNow;

        public DateTimeOffset UtcNow { get; }
    }
}
