namespace ChuanHoa.Domain.Trials;

public enum TrialCampaignStatus
{
    Draft,
    Scheduled,
    Active,
    Ended,
    Terminated
}

public enum TrialKind
{
    Launch,
    Personal
}

public enum TrialGrantStatus
{
    Active,
    Expired,
    Converted,
    Revoked
}

public sealed record TrialCampaign(
    Guid Id,
    string ProductCode,
    DateTimeOffset StartsAtUtc,
    DateTimeOffset EndsAtUtc,
    TrialCampaignStatus Status);

public sealed record TrialGrant(
    Guid Id,
    Guid SubjectId,
    string ProductCode,
    TrialKind Kind,
    DateTimeOffset StartsAtUtc,
    DateTimeOffset EndsAtUtc,
    TrialGrantStatus Status,
    string SourceReference);

public sealed record TrialEligibilityInput(
    Guid SubjectId,
    string ProductCode,
    DateTimeOffset AccountCreatedAtUtc,
    DateTimeOffset ServerNowUtc,
    TrialCampaign? LaunchCampaign,
    TimeSpan PersonalTrialDuration,
    bool HasActivePaidEntitlement,
    bool HasActiveManualGrant,
    IReadOnlyCollection<TrialGrant> HistoricalTrialGrants);

public sealed record TrialResolution(
    string AccessState,
    TrialKind? TrialKind,
    DateTimeOffset? StartsAtUtc,
    DateTimeOffset? EndsAtUtc,
    bool ShouldCreateGrant,
    string ReasonCode);
