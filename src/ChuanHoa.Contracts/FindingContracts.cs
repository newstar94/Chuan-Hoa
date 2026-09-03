using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

[JsonConverter(typeof(JsonStringEnumConverter<FindingStatus>))]
public enum FindingStatus
{
    Pass,
    Fail,
    NotChecked,
    NotApplicable
}

[JsonConverter(typeof(JsonStringEnumConverter<FindingSeverity>))]
public enum FindingSeverity
{
    Info,
    Warning,
    Error
}

[JsonConverter(typeof(JsonStringEnumConverter<RiskTier>))]
public enum RiskTier
{
    Safe,
    Confirm,
    ReportOnly,
    Blocked
}

[JsonConverter(typeof(JsonStringEnumConverter<FindingAnchorKind>))]
public enum FindingAnchorKind
{
    TextSpan,
    Paragraph,
    Section,
    Document
}

[JsonConverter(typeof(JsonStringEnumConverter<ScanLane>))]
public enum ScanLane
{
    Format,
    Spelling
}

public sealed record ScanRequest(
    string Schema,
    ScanLane Lane,
    DocumentSnapshot Snapshot);

public sealed record FindingAnchor(
    FindingAnchorKind Kind,
    string StoryType,
    int? ParagraphIndex,
    int? StartOffset,
    int? Length,
    string? ExpectedText,
    int? SectionIndex = null,
    int? TableIndex = null,
    int? RowIndex = null,
    int? CellIndex = null);

public sealed record Finding(
    string RuleId,
    FindingStatus Status,
    FindingSeverity Severity,
    RiskTier RiskTier,
    string Title,
    string Message,
    int? ParagraphIndex,
    ComponentRole? ComponentRole,
    string RuleReleaseId,
    string? ReasonCode,
    string? FindingId = null,
    string? Expected = null,
    string? Citation = null,
    FindingAnchor? Anchor = null);

public sealed record ScanResult(
    Guid JobId,
    string DocumentFingerprint,
    long Revision,
    RegimeCode Regime,
    DocumentTypeCode DocumentType,
    string RuleReleaseId,
    IReadOnlyList<Finding> Findings,
    DateTimeOffset CompletedAtUtc);
