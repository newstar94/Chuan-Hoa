using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

[JsonConverter(typeof(JsonStringEnumConverter<RuleReleaseStatus>))]
public enum RuleReleaseStatus
{
    BaselineDraft,
    Candidate,
    Published,
    Retired
}

[JsonConverter(typeof(JsonStringEnumConverter<RuleLifecycleStatus>))]
public enum RuleLifecycleStatus
{
    Draft,
    Active,
    Suspended,
    Retired
}

[JsonConverter(typeof(JsonStringEnumConverter<RuleImplementationStatus>))]
public enum RuleImplementationStatus
{
    Unrouted,
    HardwiredNotChecked,
    BaselineLogicPath,
    ImplementedVerified
}

[JsonConverter(typeof(JsonStringEnumConverter<LegalReviewStatus>))]
public enum LegalReviewStatus
{
    Unreviewed,
    Approved,
    Rejected
}

[JsonConverter(typeof(JsonStringEnumConverter<RuleFixPolicy>))]
public enum RuleFixPolicy
{
    ReportOnly,
    SignedFixPlanSafe,
    SignedFixPlanConfirm,
    Blocked
}

public sealed record CanonicalRuleRelease(
    string Schema,
    string ReleaseId,
    RuleReleaseStatus Status,
    DateTimeOffset CreatedAtUtc,
    DateTimeOffset? EffectiveFromUtc,
    DateTimeOffset? EffectiveUntilUtc,
    string SourceBaselineId,
    IReadOnlyList<CanonicalRuleDefinition> Rules);

public sealed record CanonicalRuleDefinition(
    string RuleCode,
    string RegimeCode,
    RuleLifecycleStatus LifecycleStatus,
    string Title,
    string Description,
    RuleProvenance Provenance,
    RuleImplementation Implementation,
    RuleLegalReview LegalReview,
    RuleFixtureSet Fixtures,
    RuleFixPolicy FixPolicy,
    IReadOnlyList<string> DocumentTypeCodes,
    DateTimeOffset? EffectiveFromUtc,
    DateTimeOffset? EffectiveUntilUtc);

public sealed record RuleProvenance(
    string SourceKind,
    string SourcePath,
    string SourceSymbol,
    string SourceHash);

public sealed record RuleImplementation(
    RuleImplementationStatus Status,
    string? RouteFunction,
    string? DetectorId,
    string? VerifiedEngineVersion);

public sealed record RuleLegalReview(
    LegalReviewStatus Status,
    string? Authority,
    string? Instrument,
    string? Provision,
    string? Reviewer,
    DateTimeOffset? ReviewedAtUtc);

public sealed record RuleFixtureSet(
    IReadOnlyList<string> PositiveFixtureIds,
    IReadOnlyList<string> NegativeFixtureIds,
    IReadOnlyList<string> BoundaryFixtureIds);
