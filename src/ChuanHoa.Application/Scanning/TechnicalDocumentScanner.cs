using System.Text.RegularExpressions;
using ChuanHoa.Contracts;
using ChuanHoa.Domain.Common;

namespace ChuanHoa.Application.Scanning;

public sealed class TechnicalDocumentScanner
{
    public const string RequestSchema = "chuanhoa.scan-request.v1";
    public const string RuleReleaseId = "TECHNICAL-BASELINE-V1";
    private const string BaselineReasonCode = "TECHNICAL_BASELINE_NOT_LEGAL_SIGNOFF";
    private const double PointsPerMillimeter = 72.0d / 25.4d;
    private static readonly Regex RepeatedWhitespace = new(@"[ \t]{2,}", RegexOptions.Compiled);
    private static readonly Regex WhitespaceBeforePunctuation = new(@"[ \t]+(?=[,.:;?!])", RegexOptions.Compiled);
    private readonly IClock _clock;

    public TechnicalDocumentScanner(IClock clock)
    {
        _clock = clock ?? throw new ArgumentNullException(nameof(clock));
    }

    public ScanResult Scan(ScanRequest request)
    {
        ValidateRequest(request);
        if (request.Lane != ScanLane.Format)
        {
            throw new DomainException(
                "SCAN_LANE_NOT_IMPLEMENTED",
                "The requested scan lane has no verified server checker.");
        }

        var snapshot = request.Snapshot;
        var findings = new List<Finding>();
        CheckSections(snapshot, findings);
        CheckParagraphs(snapshot, findings);

        return new ScanResult(
            Guid.NewGuid(),
            snapshot.DocumentFingerprint,
            snapshot.Revision,
            snapshot.Regime,
            snapshot.DocumentType,
            RuleReleaseId,
            findings,
            _clock.UtcNow);
    }

    private static void ValidateRequest(ScanRequest? request)
    {
        if (request is null || request.Snapshot is null)
        {
            throw new DomainException("SCAN_REQUEST_REQUIRED", "A document scan request is required.");
        }

        if (!string.Equals(request.Schema, RequestSchema, StringComparison.Ordinal))
        {
            throw new DomainException("SCAN_SCHEMA_UNSUPPORTED", "The document scan schema is not supported.");
        }

        var snapshot = request.Snapshot;
        if (snapshot.SchemaVersion != 1)
        {
            throw new DomainException("SNAPSHOT_SCHEMA_UNSUPPORTED", "The document snapshot schema is not supported.");
        }

        if (snapshot.Revision <= 0 ||
            string.IsNullOrWhiteSpace(snapshot.DocumentFingerprint) ||
            !snapshot.DocumentFingerprint.StartsWith("sha256:", StringComparison.Ordinal))
        {
            throw new DomainException("SNAPSHOT_IDENTITY_INVALID", "The document snapshot has no valid fingerprint or revision.");
        }

        var format = snapshot.Preflight.FileFormat?.Trim().ToLowerInvariant();
        if (format is not ".doc" and not ".docx")
        {
            throw new DomainException("DOCUMENT_FORMAT_UNSUPPORTED", "Only saved .doc and .docx documents are supported.");
        }

        if (!snapshot.Preflight.HasActiveWindow)
        {
            throw new DomainException("ACTIVE_WINDOW_REQUIRED", "The snapshot was not captured from an active Word window.");
        }

        if (snapshot.Paragraphs.Count > 200_000 || snapshot.Sections.Count > 5_000 || snapshot.Tables.Count > 20_000)
        {
            throw new DomainException("SNAPSHOT_LIMIT_EXCEEDED", "The document snapshot exceeds the supported safety limits.");
        }
    }

    private static void CheckSections(DocumentSnapshot snapshot, ICollection<Finding> findings)
    {
        foreach (var section in snapshot.Sections)
        {
            var expectedWidth = section.IsLandscape ? 297.0d * PointsPerMillimeter : 210.0d * PointsPerMillimeter;
            var expectedHeight = section.IsLandscape ? 210.0d * PointsPerMillimeter : 297.0d * PointsPerMillimeter;
            if (!Approximately(section.PageWidthPoints, expectedWidth, 3.0d) ||
                !Approximately(section.PageHeightPoints, expectedHeight, 3.0d))
            {
                findings.Add(CreateSectionFinding(
                    "TECH-PAGE-A4",
                    section.Index,
                    "Khổ giấy của section không phải A4",
                    "Kích thước hiện tại: " + FormatPoints(section.PageWidthPoints) + " × " +
                    FormatPoints(section.PageHeightPoints) + " pt.",
                    section.IsLandscape ? "A4 ngang 841,9 × 595,3 pt" : "A4 dọc 595,3 × 841,9 pt"));
            }

            if (snapshot.Regime == RegimeCode.Nd30 && !MarginsAreWithinNd30TechnicalBaseline(section))
            {
                findings.Add(CreateSectionFinding(
                    "TECH-ND30-MARGINS",
                    section.Index,
                    "Lề trang nằm ngoài baseline kỹ thuật NĐ30",
                    "Lề trên/dưới/trái/phải hiện tại: " +
                    FormatMillimeters(section.TopMarginPoints) + "/" +
                    FormatMillimeters(section.BottomMarginPoints) + "/" +
                    FormatMillimeters(section.LeftMarginPoints) + "/" +
                    FormatMillimeters(section.RightMarginPoints) + " mm.",
                    "Trên 20–25 mm; dưới 20–25 mm; trái 30–35 mm; phải 15–20 mm"));
            }
        }
    }

    private static void CheckParagraphs(DocumentSnapshot snapshot, ICollection<Finding> findings)
    {
        foreach (var paragraph in snapshot.Paragraphs)
        {
            var printableText = (paragraph.Text ?? string.Empty).TrimEnd('\r', '\a');
            if (printableText.Length == 0)
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(paragraph.FontName) &&
                !string.Equals(paragraph.FontName.Trim(), "Times New Roman", StringComparison.OrdinalIgnoreCase))
            {
                findings.Add(new Finding(
                    "TECH-FONT-TNR",
                    FindingStatus.Fail,
                    FindingSeverity.Error,
                    RiskTier.ReportOnly,
                    "Phông chữ không phải Times New Roman",
                    "Đoạn đang dùng phông '" + paragraph.FontName + "'.",
                    paragraph.Index,
                    paragraph.Role,
                    RuleReleaseId,
                    BaselineReasonCode,
                    CreateFindingId("TECH-FONT-TNR", paragraph, 0),
                    "Times New Roman",
                    null,
                    CreateParagraphAnchor(paragraph)));
            }

            AddTextSpanFindings(
                paragraph,
                printableText,
                RepeatedWhitespace,
                "TECH-WHITESPACE-REPEATED",
                "Có khoảng trắng hoặc tab lặp",
                "Vùng được đánh dấu có nhiều ký tự trắng liên tiếp.",
                "Một dấu cách",
                findings);
            AddTextSpanFindings(
                paragraph,
                printableText,
                WhitespaceBeforePunctuation,
                "TECH-WHITESPACE-BEFORE-PUNCTUATION",
                "Có khoảng trắng trước dấu câu",
                "Vùng được đánh dấu là khoảng trắng đứng ngay trước dấu câu.",
                "Không có khoảng trắng trước dấu câu",
                findings);
        }
    }

    private static void AddTextSpanFindings(
        ParagraphSnapshot paragraph,
        string text,
        Regex pattern,
        string ruleId,
        string title,
        string message,
        string expected,
        ICollection<Finding> findings)
    {
        foreach (Match match in pattern.Matches(text))
        {
            findings.Add(new Finding(
                ruleId,
                FindingStatus.Fail,
                FindingSeverity.Warning,
                RiskTier.ReportOnly,
                title,
                message,
                paragraph.Index,
                paragraph.Role,
                RuleReleaseId,
                BaselineReasonCode,
                CreateFindingId(ruleId, paragraph, match.Index),
                expected,
                null,
                new FindingAnchor(
                    FindingAnchorKind.TextSpan,
                    paragraph.StoryType,
                    paragraph.Index,
                    match.Index,
                    match.Length,
                    match.Value,
                    paragraph.SectionIndex,
                    paragraph.TableIndex,
                    paragraph.RowIndex,
                    paragraph.CellIndex)));
        }
    }

    private static Finding CreateSectionFinding(
        string ruleId,
        int sectionIndex,
        string title,
        string message,
        string expected)
    {
        return new Finding(
            ruleId,
            FindingStatus.Fail,
            FindingSeverity.Error,
            RiskTier.ReportOnly,
            title,
            message,
            null,
            null,
            RuleReleaseId,
            BaselineReasonCode,
            ruleId + ":section:" + sectionIndex,
            expected,
            null,
            new FindingAnchor(
                FindingAnchorKind.Section,
                "MainTextStory",
                null,
                null,
                null,
                null,
                sectionIndex));
    }

    private static FindingAnchor CreateParagraphAnchor(ParagraphSnapshot paragraph)
    {
        return new FindingAnchor(
            FindingAnchorKind.Paragraph,
            paragraph.StoryType,
            paragraph.Index,
            null,
            null,
            null,
            paragraph.SectionIndex,
            paragraph.TableIndex,
            paragraph.RowIndex,
            paragraph.CellIndex);
    }

    private static string CreateFindingId(string ruleId, ParagraphSnapshot paragraph, int offset)
    {
        return ruleId + ":" + paragraph.StoryType + ":" + paragraph.SectionIndex + ":" +
            paragraph.Index + ":" + offset;
    }

    private static bool MarginsAreWithinNd30TechnicalBaseline(SectionSnapshot section)
    {
        return InMillimeterRange(section.TopMarginPoints, 20.0d, 25.0d) &&
            InMillimeterRange(section.BottomMarginPoints, 20.0d, 25.0d) &&
            InMillimeterRange(section.LeftMarginPoints, 30.0d, 35.0d) &&
            InMillimeterRange(section.RightMarginPoints, 15.0d, 20.0d);
    }

    private static bool InMillimeterRange(double points, double minimum, double maximum)
    {
        var millimeters = points / PointsPerMillimeter;
        return millimeters >= minimum - 0.25d && millimeters <= maximum + 0.25d;
    }

    private static bool Approximately(double actual, double expected, double tolerance)
    {
        return Math.Abs(actual - expected) <= tolerance;
    }

    private static string FormatPoints(double value) => value.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);

    private static string FormatMillimeters(double value) =>
        (value / PointsPerMillimeter).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
}
