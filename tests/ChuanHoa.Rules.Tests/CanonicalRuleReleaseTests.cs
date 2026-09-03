using System.Text;
using System.Text.Json;
using ChuanHoa.Contracts;
using ChuanHoa.Rules;

namespace ChuanHoa.Rules.Tests;

public sealed class CanonicalRuleReleaseTests
{
    private readonly CanonicalRuleReleaseParser _parser = new();
    private readonly CanonicalRuleReleaseValidator _validator = new();

    [Fact]
    public void Baseline_draft_preserves_exact_route_classification_without_becoming_publishable()
    {
        using var stream = File.OpenRead(GetRulePath("baseline-draft.v1.json"));
        var release = _parser.Parse(stream);
        var validation = _validator.Validate(release);

        Assert.True(validation.IsValid);
        Assert.False(validation.IsPublishable);
        Assert.Equal(RuleReleaseStatus.BaselineDraft, release.Status);
        Assert.Equal(96, release.Rules.Count);
        Assert.Equal(
            2,
            release.Rules.Count(rule =>
                rule.Implementation.Status == RuleImplementationStatus.Unrouted));
        Assert.Equal(
            19,
            release.Rules.Count(rule =>
                rule.Implementation.Status == RuleImplementationStatus.HardwiredNotChecked));
        Assert.Equal(
            75,
            release.Rules.Count(rule =>
                rule.Implementation.Status == RuleImplementationStatus.BaselineLogicPath));
        Assert.All(
            release.Rules,
            rule =>
            {
                Assert.Equal(RuleLifecycleStatus.Draft, rule.LifecycleStatus);
                Assert.Equal(LegalReviewStatus.Unreviewed, rule.LegalReview.Status);
                Assert.Equal(RuleFixPolicy.Blocked, rule.FixPolicy);
                Assert.Empty(rule.Fixtures.PositiveFixtureIds);
                Assert.Empty(rule.Fixtures.NegativeFixtureIds);
                Assert.Empty(rule.Fixtures.BoundaryFixtureIds);
            });
    }

    [Fact]
    public void Parser_rejects_unknown_properties()
    {
        var json = Serialize(CreateVerifiedRelease());
        var mutated = json.Replace(
            "\"releaseId\":",
            "\"unknownField\":true,\"releaseId\":",
            StringComparison.Ordinal);

        Assert.Throws<JsonException>(() => _parser.Parse(Encoding.UTF8.GetBytes(mutated)));
    }

    [Fact]
    public void Parser_rejects_unknown_enum_values()
    {
        var json = Serialize(CreateVerifiedRelease()).Replace(
            "\"Published\"",
            "\"Uncontrolled\"",
            StringComparison.Ordinal);

        Assert.Throws<JsonException>(() => _parser.Parse(Encoding.UTF8.GetBytes(json)));
    }

    [Fact]
    public void Published_release_rejects_unverified_implementation()
    {
        var release = CreateVerifiedRelease();
        var rule = release.Rules[0] with
        {
            Implementation = new RuleImplementation(
                RuleImplementationStatus.BaselineLogicPath,
                "CheckPageSizeA4",
                null,
                null)
        };

        var result = _validator.Validate(release with { Rules = [rule] });

        Assert.False(result.IsPublishable);
        Assert.Contains(
            result.Errors,
            error => error.Code == "PUBLISHED_RULE_IMPLEMENTATION_UNVERIFIED");
    }

    [Fact]
    public void Published_release_rejects_missing_legal_traceability()
    {
        var release = CreateVerifiedRelease();
        var rule = release.Rules[0] with
        {
            LegalReview = new RuleLegalReview(
                LegalReviewStatus.Unreviewed,
                null,
                null,
                null,
                null,
                null)
        };

        var result = _validator.Validate(release with { Rules = [rule] });

        Assert.False(result.IsPublishable);
        Assert.Contains(
            result.Errors,
            error => error.Code == "PUBLISHED_RULE_LEGAL_REVIEW_REQUIRED");
    }

    [Fact]
    public void Published_release_rejects_missing_fixture_classes()
    {
        var release = CreateVerifiedRelease();
        var rule = release.Rules[0] with
        {
            Fixtures = new RuleFixtureSet(["positive-001"], [], ["boundary-001"])
        };

        var result = _validator.Validate(release with { Rules = [rule] });

        Assert.False(result.IsPublishable);
        Assert.Contains(
            result.Errors,
            error => error.Code == "PUBLISHED_RULE_FIXTURES_REQUIRED");
    }

    [Fact]
    public void Fully_traced_verified_rule_is_publishable()
    {
        var result = _validator.Validate(CreateVerifiedRelease());

        Assert.True(result.IsPublishable);
        Assert.Empty(result.Errors);
    }

    [Fact]
    public void Json_schema_is_closed_at_release_rule_and_nested_objects()
    {
        using var stream = File.OpenRead(GetRulePath("canonical-rule-release.v1.schema.json"));
        using var document = JsonDocument.Parse(stream);
        var root = document.RootElement;

        Assert.False(root.GetProperty("additionalProperties").GetBoolean());
        var definitions = root.GetProperty("$defs");
        Assert.False(definitions.GetProperty("rule").GetProperty("additionalProperties").GetBoolean());
        Assert.False(definitions.GetProperty("provenance").GetProperty("additionalProperties").GetBoolean());
        Assert.False(definitions.GetProperty("implementation").GetProperty("additionalProperties").GetBoolean());
        Assert.False(definitions.GetProperty("legalReview").GetProperty("additionalProperties").GetBoolean());
        Assert.False(definitions.GetProperty("fixtures").GetProperty("additionalProperties").GetBoolean());
    }

    [Fact]
    public void Draft_fixture_is_valid_but_not_golden_eligible()
    {
        var validator = new CanonicalRuleFixtureValidator();
        var fixture = CreateFixture(
            "nd30-page-size-positive-001",
            RuleFixtureKind.Positive,
            FindingStatus.Pass,
            RuleFixtureReviewStatus.Draft);

        var result = validator.Validate(fixture);

        Assert.True(result.IsValid);
        Assert.False(result.IsGoldenEligible);
    }

    [Fact]
    public void Golden_pack_requires_approved_positive_negative_and_boundary_fixtures()
    {
        var validator = new CanonicalRuleFixtureValidator();
        var incomplete = new[]
        {
            CreateFixture(
                "nd30-page-size-positive-001",
                RuleFixtureKind.Positive,
                FindingStatus.Pass,
                RuleFixtureReviewStatus.Approved),
            CreateFixture(
                "nd30-page-size-negative-001",
                RuleFixtureKind.Negative,
                FindingStatus.Fail,
                RuleFixtureReviewStatus.Approved)
        };

        var incompleteResult = validator.ValidatePack("ND30-PL1-M1-K1", incomplete);
        var complete = incomplete.Append(
            CreateFixture(
                "nd30-page-size-boundary-001",
                RuleFixtureKind.Boundary,
                FindingStatus.NotChecked,
                RuleFixtureReviewStatus.Approved)).ToArray();
        var completeResult = validator.ValidatePack("ND30-PL1-M1-K1", complete);

        Assert.False(incompleteResult.IsComplete);
        Assert.Contains(
            incompleteResult.Errors,
            error => error.Code == "FIXTURE_KIND_MISSING");
        Assert.True(completeResult.IsComplete);
        Assert.Empty(completeResult.Errors);
    }

    private static CanonicalRuleRelease CreateVerifiedRelease()
    {
        var rule = new CanonicalRuleDefinition(
            "ND30-PL1-M1-K1",
            "ND30",
            RuleLifecycleStatus.Active,
            "Khổ giấy A4",
            "Văn bản dùng khổ giấy A4.",
            new RuleProvenance(
                "LEGAL_INSTRUMENT",
                "legal/nd30/appendix-1",
                "PhuLucI.MucI.Khoan1",
                new string('A', 64)),
            new RuleImplementation(
                RuleImplementationStatus.ImplementedVerified,
                "CheckPageSizeA4",
                "PAGE_SIZE_A4_V1",
                "engine-1.0.0"),
            new RuleLegalReview(
                LegalReviewStatus.Approved,
                "Chính phủ",
                "Nghị định 30/2020/NĐ-CP",
                "Phụ lục I, Phần I, Khoản 1",
                "legal-reviewer-001",
                DateTimeOffset.Parse("2026-09-01T00:00:00Z")),
            new RuleFixtureSet(
                ["nd30-page-size-positive-001"],
                ["nd30-page-size-negative-001"],
                ["nd30-page-size-boundary-001"]),
            RuleFixPolicy.SignedFixPlanSafe,
            ["OfficialLetter", "Decision"],
            DateTimeOffset.Parse("2026-09-01T00:00:00Z"),
            null);

        return new CanonicalRuleRelease(
            CanonicalRuleReleaseValidator.SupportedSchema,
            "ND30-RELEASE-0001",
            RuleReleaseStatus.Published,
            DateTimeOffset.Parse("2026-09-01T00:00:00Z"),
            DateTimeOffset.Parse("2026-09-01T00:00:00Z"),
            null,
            "LEGAL-BASELINE-0001",
            [rule]);
    }

    private static CanonicalRuleFixture CreateFixture(
        string fixtureId,
        RuleFixtureKind kind,
        FindingStatus expectedStatus,
        RuleFixtureReviewStatus reviewStatus)
    {
        var snapshot = new DocumentSnapshot(
            1,
            $"fixture:{fixtureId}",
            0,
            RegimeCode.Nd30,
            DocumentTypeCode.OfficialLetter,
            true,
            true,
            new DocumentPreflight(false, false, false, false, "DOCX", true),
            [
                new SectionSnapshot(1, 595.3, 841.9, 56.7, 56.7, 85.0, 56.7, false)
            ],
            [],
            []);
        return new CanonicalRuleFixture(
            CanonicalRuleFixtureValidator.SupportedSchema,
            fixtureId,
            "ND30-PL1-M1-K1",
            kind,
            reviewStatus,
            reviewStatus == RuleFixtureReviewStatus.Approved ? "legal-reviewer-001" : null,
            reviewStatus == RuleFixtureReviewStatus.Approved
                ? DateTimeOffset.Parse("2026-09-01T00:00:00Z")
                : null,
            new string('B', 64),
            snapshot,
            expectedStatus,
            expectedStatus == FindingStatus.NotChecked ? "BOUNDARY_REQUIRES_WORD_VM" : null);
    }

    private static string Serialize(CanonicalRuleRelease release)
    {
        return JsonSerializer.Serialize(
            release,
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                Converters =
                {
                    new System.Text.Json.Serialization.JsonStringEnumConverter()
                }
            });
    }

    private static string GetRulePath(string fileName)
    {
        return Path.Combine(AppContext.BaseDirectory, "Rules", fileName);
    }
}
