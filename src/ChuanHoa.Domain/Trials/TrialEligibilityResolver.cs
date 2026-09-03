using ChuanHoa.Domain.Common;

namespace ChuanHoa.Domain.Trials;

public sealed class TrialEligibilityResolver
{
    public TrialResolution Resolve(TrialEligibilityInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        if (string.IsNullOrWhiteSpace(input.ProductCode))
        {
            throw new DomainException("PRODUCT_CODE_REQUIRED", "Product code is required.");
        }

        if (input.PersonalTrialDuration <= TimeSpan.Zero)
        {
            throw new DomainException("PERSONAL_TRIAL_DURATION_INVALID", "Personal trial duration must be positive.");
        }

        if (input.HasActivePaidEntitlement)
        {
            return new TrialResolution("PAID_ACTIVE", null, null, null, false, "ACTIVE_PAID_ENTITLEMENT");
        }

        if (input.HasActiveManualGrant)
        {
            return new TrialResolution("MANUAL_GRANT", null, null, null, false, "ACTIVE_MANUAL_GRANT");
        }

        var existing = input.HistoricalTrialGrants
            .Where(grant => grant.SubjectId == input.SubjectId && grant.ProductCode == input.ProductCode)
            .OrderByDescending(grant => grant.StartsAtUtc)
            .FirstOrDefault();

        if (existing is not null)
        {
            if (existing.Status == TrialGrantStatus.Active && input.ServerNowUtc < existing.EndsAtUtc)
            {
                return new TrialResolution(
                    existing.Kind == TrialKind.Launch ? "LAUNCH_TRIAL" : "PERSONAL_TRIAL",
                    existing.Kind,
                    existing.StartsAtUtc,
                    existing.EndsAtUtc,
                    false,
                    "EXISTING_ACTIVE_TRIAL");
            }

            return new TrialResolution("PAID_REQUIRED", null, null, null, false, "TRIAL_ALREADY_CONSUMED");
        }

        var campaign = input.LaunchCampaign;
        if (campaign is not null &&
            campaign.ProductCode == input.ProductCode &&
            campaign.Status is TrialCampaignStatus.Scheduled or TrialCampaignStatus.Active &&
            input.ServerNowUtc >= campaign.StartsAtUtc &&
            input.ServerNowUtc < campaign.EndsAtUtc)
        {
            return new TrialResolution(
                "LAUNCH_TRIAL",
                TrialKind.Launch,
                input.ServerNowUtc,
                campaign.EndsAtUtc,
                true,
                "LAUNCH_CAMPAIGN_ACTIVE");
        }

        if (campaign is not null && input.ServerNowUtc < campaign.EndsAtUtc)
        {
            return new TrialResolution("PAID_REQUIRED", null, null, null, false, "LAUNCH_CAMPAIGN_NOT_ACTIVE");
        }

        if (campaign is not null && input.AccountCreatedAtUtc < campaign.EndsAtUtc)
        {
            return new TrialResolution("PAID_REQUIRED", null, null, null, false, "ACCOUNT_PREDATES_PERSONAL_TRIAL_ERA");
        }

        var personalEndsAt = input.ServerNowUtc.Add(input.PersonalTrialDuration);
        return new TrialResolution(
            "PERSONAL_TRIAL",
            TrialKind.Personal,
            input.ServerNowUtc,
            personalEndsAt,
            true,
            "FIRST_PERSONAL_TRIAL_ELIGIBLE");
    }
}
