using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using ChuanHoa.Contracts;

namespace ChuanHoa.Rules;

public sealed partial class CanonicalRuleReleaseValidator
{
    public const string SupportedSchema = "chuanhoa.canonical-rule-release.v1";

    public RuleReleaseValidationResult Validate(CanonicalRuleRelease release)
    {
        ArgumentNullException.ThrowIfNull(release);
        var errors = new List<RuleReleaseValidationError>();

        if (!string.Equals(release.Schema, SupportedSchema, StringComparison.Ordinal))
        {
            errors.Add(new("RULE_SCHEMA_UNSUPPORTED", "$.schema", "The canonical rule schema is not supported."));
        }

        if (string.IsNullOrWhiteSpace(release.ReleaseId) || !ReleaseIdPattern().IsMatch(release.ReleaseId))
        {
            errors.Add(new(
                "RULE_RELEASE_ID_INVALID",
                "$.releaseId",
                "Release id must contain 8 to 80 uppercase letters, digits, dot, underscore, or hyphen characters."));
        }

        if (string.IsNullOrWhiteSpace(release.SourceBaselineId))
        {
            errors.Add(new("RULE_BASELINE_ID_REQUIRED", "$.sourceBaselineId", "Source baseline id is required."));
        }

        if (release.EffectiveUntilUtc is not null &&
            (release.EffectiveFromUtc is null || release.EffectiveUntilUtc <= release.EffectiveFromUtc))
        {
            errors.Add(new(
                "RULE_RELEASE_PERIOD_INVALID",
                "$.effectiveUntilUtc",
                "Release end must be after release start."));
        }

        if (release.Rules.Count == 0)
        {
            errors.Add(new("RULE_RELEASE_EMPTY", "$.rules", "A canonical rule release must contain at least one rule."));
        }

        var duplicateCodes = release.Rules
            .GroupBy(rule => rule.RuleCode, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .OrderBy(code => code, StringComparer.Ordinal);
        foreach (var duplicateCode in duplicateCodes)
        {
            errors.Add(new(
                "RULE_CODE_DUPLICATE",
                "$.rules",
                $"Rule code occurs more than once: {duplicateCode}."));
        }

        for (var index = 0; index < release.Rules.Count; index++)
        {
            ValidateRule(release, release.Rules[index], index, errors);
        }

        return new RuleReleaseValidationResult(
            new ReadOnlyCollection<RuleReleaseValidationError>(errors),
            errors.Count == 0,
            errors.Count == 0 && release.Status == RuleReleaseStatus.Published);
    }

    private static void ValidateRule(
        CanonicalRuleRelease release,
        CanonicalRuleDefinition rule,
        int index,
        ICollection<RuleReleaseValidationError> errors)
    {
        var path = $"$.rules[{index}]";
        if (string.IsNullOrWhiteSpace(rule.RuleCode) || !RuleCodePattern().IsMatch(rule.RuleCode))
        {
            errors.Add(new("RULE_CODE_INVALID", $"{path}.ruleCode", "Rule code format is invalid."));
        }

        if (string.IsNullOrWhiteSpace(rule.RegimeCode))
        {
            errors.Add(new("RULE_REGIME_REQUIRED", $"{path}.regimeCode", "Regime code is required."));
        }

        if (string.IsNullOrWhiteSpace(rule.Title))
        {
            errors.Add(new("RULE_TITLE_REQUIRED", $"{path}.title", "Rule title is required."));
        }

        if (string.IsNullOrWhiteSpace(rule.Description))
        {
            errors.Add(new("RULE_DESCRIPTION_REQUIRED", $"{path}.description", "Rule description is required."));
        }

        if (string.IsNullOrWhiteSpace(rule.Provenance.SourceKind) ||
            string.IsNullOrWhiteSpace(rule.Provenance.SourcePath) ||
            string.IsNullOrWhiteSpace(rule.Provenance.SourceSymbol) ||
            !Sha256Pattern().IsMatch(rule.Provenance.SourceHash))
        {
            errors.Add(new(
                "RULE_PROVENANCE_INVALID",
                $"{path}.provenance",
                "Provenance requires source kind/path/symbol and an uppercase SHA-256 hash."));
        }

        if (rule.EffectiveUntilUtc is not null &&
            (rule.EffectiveFromUtc is null || rule.EffectiveUntilUtc <= rule.EffectiveFromUtc))
        {
            errors.Add(new(
                "RULE_PERIOD_INVALID",
                $"{path}.effectiveUntilUtc",
                "Rule end must be after rule start."));
        }

        if (release.Status != RuleReleaseStatus.Published)
        {
            return;
        }

        if (rule.LifecycleStatus != RuleLifecycleStatus.Active)
        {
            errors.Add(new(
                "PUBLISHED_RULE_NOT_ACTIVE",
                $"{path}.lifecycleStatus",
                "Every rule in a published release must be active."));
        }

        if (rule.Implementation.Status != RuleImplementationStatus.ImplementedVerified ||
            string.IsNullOrWhiteSpace(rule.Implementation.RouteFunction) ||
            string.IsNullOrWhiteSpace(rule.Implementation.DetectorId) ||
            string.IsNullOrWhiteSpace(rule.Implementation.VerifiedEngineVersion))
        {
            errors.Add(new(
                "PUBLISHED_RULE_IMPLEMENTATION_UNVERIFIED",
                $"{path}.implementation",
                "Published rules require a verified implementation, route, detector, and engine version."));
        }

        if (rule.LegalReview.Status != LegalReviewStatus.Approved ||
            string.IsNullOrWhiteSpace(rule.LegalReview.Authority) ||
            string.IsNullOrWhiteSpace(rule.LegalReview.Instrument) ||
            string.IsNullOrWhiteSpace(rule.LegalReview.Provision) ||
            string.IsNullOrWhiteSpace(rule.LegalReview.Reviewer) ||
            rule.LegalReview.ReviewedAtUtc is null)
        {
            errors.Add(new(
                "PUBLISHED_RULE_LEGAL_REVIEW_REQUIRED",
                $"{path}.legalReview",
                "Published rules require approved legal traceability and reviewer evidence."));
        }

        if (rule.Fixtures.PositiveFixtureIds.Count == 0 ||
            rule.Fixtures.NegativeFixtureIds.Count == 0 ||
            rule.Fixtures.BoundaryFixtureIds.Count == 0)
        {
            errors.Add(new(
                "PUBLISHED_RULE_FIXTURES_REQUIRED",
                $"{path}.fixtures",
                "Published rules require positive, negative, and boundary fixtures."));
        }

        if (rule.FixPolicy == RuleFixPolicy.Blocked)
        {
            errors.Add(new(
                "PUBLISHED_RULE_FIX_POLICY_BLOCKED",
                $"{path}.fixPolicy",
                "A published rule cannot retain a blocked fix policy."));
        }
    }

    [GeneratedRegex("^[A-Z0-9][A-Z0-9._-]{7,79}$", RegexOptions.CultureInvariant)]
    private static partial Regex ReleaseIdPattern();

    [GeneratedRegex("^[A-Z0-9][A-Z0-9._-]{2,99}$", RegexOptions.CultureInvariant)]
    private static partial Regex RuleCodePattern();

    [GeneratedRegex("^[A-F0-9]{64}$", RegexOptions.CultureInvariant)]
    private static partial Regex Sha256Pattern();
}

public sealed record RuleReleaseValidationError(
    string Code,
    string Path,
    string Message);

public sealed record RuleReleaseValidationResult(
    IReadOnlyList<RuleReleaseValidationError> Errors,
    bool IsValid,
    bool IsPublishable);
