using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using ChuanHoa.Contracts;

namespace ChuanHoa.Rules;

public sealed partial class CanonicalRuleFixtureValidator
{
    public const string SupportedSchema = "chuanhoa.canonical-rule-fixture.v1";

    public RuleFixtureValidationResult Validate(CanonicalRuleFixture fixture)
    {
        ArgumentNullException.ThrowIfNull(fixture);
        var errors = new List<RuleFixtureValidationError>();

        if (!string.Equals(fixture.Schema, SupportedSchema, StringComparison.Ordinal))
        {
            errors.Add(new("FIXTURE_SCHEMA_UNSUPPORTED", "$.schema", "Fixture schema is not supported."));
        }

        if (string.IsNullOrWhiteSpace(fixture.FixtureId) ||
            !FixtureIdPattern().IsMatch(fixture.FixtureId))
        {
            errors.Add(new("FIXTURE_ID_INVALID", "$.fixtureId", "Fixture id format is invalid."));
        }

        if (string.IsNullOrWhiteSpace(fixture.RuleCode))
        {
            errors.Add(new("FIXTURE_RULE_CODE_REQUIRED", "$.ruleCode", "Rule code is required."));
        }

        if (!Sha256Pattern().IsMatch(fixture.SourceDocumentHash))
        {
            errors.Add(new(
                "FIXTURE_SOURCE_HASH_INVALID",
                "$.sourceDocumentHash",
                "Fixture source document hash must be uppercase SHA-256."));
        }

        if (string.IsNullOrWhiteSpace(fixture.Snapshot.DocumentFingerprint))
        {
            errors.Add(new(
                "FIXTURE_DOCUMENT_FINGERPRINT_REQUIRED",
                "$.snapshot.documentFingerprint",
                "Fixture snapshot must have a document fingerprint."));
        }

        if (fixture.Snapshot.Revision < 0)
        {
            errors.Add(new(
                "FIXTURE_REVISION_INVALID",
                "$.snapshot.revision",
                "Fixture snapshot revision cannot be negative."));
        }

        if (fixture.ReviewStatus == RuleFixtureReviewStatus.Approved &&
            (string.IsNullOrWhiteSpace(fixture.Reviewer) || fixture.ReviewedAtUtc is null))
        {
            errors.Add(new(
                "FIXTURE_APPROVAL_EVIDENCE_REQUIRED",
                "$.reviewStatus",
                "Approved fixture requires reviewer and reviewed time."));
        }

        if (fixture.Kind == RuleFixtureKind.Positive && fixture.ExpectedStatus != FindingStatus.Pass)
        {
            errors.Add(new(
                "POSITIVE_FIXTURE_EXPECTATION_INVALID",
                "$.expectedStatus",
                "Positive fixtures must expect Pass."));
        }

        if (fixture.Kind == RuleFixtureKind.Negative && fixture.ExpectedStatus != FindingStatus.Fail)
        {
            errors.Add(new(
                "NEGATIVE_FIXTURE_EXPECTATION_INVALID",
                "$.expectedStatus",
                "Negative fixtures must expect Fail."));
        }

        return new RuleFixtureValidationResult(
            new ReadOnlyCollection<RuleFixtureValidationError>(errors),
            errors.Count == 0,
            errors.Count == 0 && fixture.ReviewStatus == RuleFixtureReviewStatus.Approved);
    }

    public RuleFixturePackValidationResult ValidatePack(
        string ruleCode,
        IReadOnlyCollection<CanonicalRuleFixture> fixtures)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(ruleCode);
        ArgumentNullException.ThrowIfNull(fixtures);
        var errors = new List<RuleFixtureValidationError>();
        var ids = new HashSet<string>(StringComparer.Ordinal);

        foreach (var fixture in fixtures)
        {
            if (!ids.Add(fixture.FixtureId))
            {
                errors.Add(new(
                    "FIXTURE_ID_DUPLICATE",
                    "$.fixtures",
                    $"Fixture id occurs more than once: {fixture.FixtureId}."));
            }

            if (!string.Equals(fixture.RuleCode, ruleCode, StringComparison.Ordinal))
            {
                errors.Add(new(
                    "FIXTURE_RULE_CODE_MISMATCH",
                    "$.fixtures",
                    $"Fixture {fixture.FixtureId} belongs to {fixture.RuleCode}, not {ruleCode}."));
            }

            var fixtureValidation = Validate(fixture);
            errors.AddRange(fixtureValidation.Errors);
            if (!fixtureValidation.IsGoldenEligible)
            {
                errors.Add(new(
                    "FIXTURE_NOT_GOLDEN_ELIGIBLE",
                    "$.fixtures",
                    $"Fixture {fixture.FixtureId} is not approved for golden evidence."));
            }
        }

        foreach (var requiredKind in Enum.GetValues<RuleFixtureKind>())
        {
            if (!fixtures.Any(fixture => fixture.Kind == requiredKind))
            {
                errors.Add(new(
                    "FIXTURE_KIND_MISSING",
                    "$.fixtures",
                    $"Fixture pack is missing {requiredKind}."));
            }
        }

        return new RuleFixturePackValidationResult(
            new ReadOnlyCollection<RuleFixtureValidationError>(errors),
            errors.Count == 0);
    }

    [GeneratedRegex("^[a-z0-9][a-z0-9._-]{7,119}$", RegexOptions.CultureInvariant)]
    private static partial Regex FixtureIdPattern();

    [GeneratedRegex("^[A-F0-9]{64}$", RegexOptions.CultureInvariant)]
    private static partial Regex Sha256Pattern();
}

public sealed record RuleFixtureValidationError(
    string Code,
    string Path,
    string Message);

public sealed record RuleFixtureValidationResult(
    IReadOnlyList<RuleFixtureValidationError> Errors,
    bool IsValid,
    bool IsGoldenEligible);

public sealed record RuleFixturePackValidationResult(
    IReadOnlyList<RuleFixtureValidationError> Errors,
    bool IsComplete);
