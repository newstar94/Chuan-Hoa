using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

[JsonConverter(typeof(JsonStringEnumConverter<RuleFixtureKind>))]
public enum RuleFixtureKind
{
    Positive,
    Negative,
    Boundary
}

[JsonConverter(typeof(JsonStringEnumConverter<RuleFixtureReviewStatus>))]
public enum RuleFixtureReviewStatus
{
    Draft,
    Approved,
    Rejected
}

public sealed record CanonicalRuleFixture(
    string Schema,
    string FixtureId,
    string RuleCode,
    RuleFixtureKind Kind,
    RuleFixtureReviewStatus ReviewStatus,
    string? Reviewer,
    DateTimeOffset? ReviewedAtUtc,
    string SourceDocumentHash,
    DocumentSnapshot Snapshot,
    FindingStatus ExpectedStatus,
    string? ExpectedReasonCode);
