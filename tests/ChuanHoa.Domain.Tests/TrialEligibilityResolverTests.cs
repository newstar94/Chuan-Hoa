using ChuanHoa.Domain.Trials;

namespace ChuanHoa.Domain.Tests;

public sealed class TrialEligibilityResolverTests
{
    private static readonly Guid SubjectId = Guid.Parse("11111111-1111-1111-1111-111111111111");
    private static readonly DateTimeOffset CampaignStart = DateTimeOffset.Parse("2026-10-01T00:00:00Z");
    private static readonly DateTimeOffset CampaignEnd = DateTimeOffset.Parse("2026-11-01T00:00:00Z");
    private readonly TrialEligibilityResolver _resolver = new();

    [Fact]
    public void Paid_entitlement_has_highest_priority()
    {
        var result = _resolver.Resolve(Input(CampaignStart, hasPaid: true));
        Assert.Equal("PAID_ACTIVE", result.AccessState);
        Assert.False(result.ShouldCreateGrant);
    }

    [Fact]
    public void Launch_trial_starts_at_first_premium_use_and_ends_at_campaign_boundary()
    {
        var firstUse = CampaignStart.AddDays(20);
        var result = _resolver.Resolve(Input(firstUse));
        Assert.Equal("LAUNCH_TRIAL", result.AccessState);
        Assert.Equal(firstUse, result.StartsAtUtc);
        Assert.Equal(CampaignEnd, result.EndsAtUtc);
        Assert.True(result.ShouldCreateGrant);
    }

    [Fact]
    public void Campaign_end_is_exclusive()
    {
        var result = _resolver.Resolve(Input(CampaignEnd, accountCreatedAt: CampaignEnd));
        Assert.Equal("PERSONAL_TRIAL", result.AccessState);
        Assert.Equal(TrialKind.Personal, result.TrialKind);
    }

    [Fact]
    public void Account_created_before_launch_end_is_not_given_personal_trial()
    {
        var result = _resolver.Resolve(Input(CampaignEnd.AddDays(1), accountCreatedAt: CampaignStart));
        Assert.Equal("PAID_REQUIRED", result.AccessState);
        Assert.Equal("ACCOUNT_PREDATES_PERSONAL_TRIAL_ERA", result.ReasonCode);
    }

    [Fact]
    public void Historical_launch_trial_prevents_personal_trial()
    {
        var historical = new TrialGrant(
            Guid.NewGuid(),
            SubjectId,
            "CHUAN_HOA",
            TrialKind.Launch,
            CampaignStart,
            CampaignEnd,
            TrialGrantStatus.Expired,
            "campaign");
        var result = _resolver.Resolve(Input(CampaignEnd.AddDays(1), historical: [historical]));
        Assert.Equal("PAID_REQUIRED", result.AccessState);
        Assert.Equal("TRIAL_ALREADY_CONSUMED", result.ReasonCode);
    }

    [Fact]
    public void Existing_active_trial_is_returned_idempotently()
    {
        var historical = new TrialGrant(
            Guid.NewGuid(),
            SubjectId,
            "CHUAN_HOA",
            TrialKind.Personal,
            CampaignEnd.AddDays(1),
            CampaignEnd.AddDays(8),
            TrialGrantStatus.Active,
            "personal");
        var result = _resolver.Resolve(Input(CampaignEnd.AddDays(2), historical: [historical]));
        Assert.Equal("PERSONAL_TRIAL", result.AccessState);
        Assert.False(result.ShouldCreateGrant);
    }

    private static TrialEligibilityInput Input(
        DateTimeOffset now,
        bool hasPaid = false,
        bool hasManual = false,
        IReadOnlyCollection<TrialGrant>? historical = null,
        DateTimeOffset? accountCreatedAt = null)
    {
        return new TrialEligibilityInput(
            SubjectId,
            "CHUAN_HOA",
            accountCreatedAt ?? CampaignStart.AddDays(-30),
            now,
            new TrialCampaign(Guid.NewGuid(), "CHUAN_HOA", CampaignStart, CampaignEnd, TrialCampaignStatus.Active),
            TimeSpan.FromDays(7),
            hasPaid,
            hasManual,
            historical ?? Array.Empty<TrialGrant>());
    }
}
