using System;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ChuanHoa.Client.Core.Licensing;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Security;

namespace ChuanHoa.Client.Core.Tests;

public sealed class SignedRuleLeaseTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");

    [Fact]
    public void Signed_rule_pack_is_verified_before_parsing()
    {
        using var rsa = RSA.Create(2048);
        var payload = RulePayload(Now.AddMinutes(-1), Now.AddDays(30));
        var encoded = Sign(rsa, "dev-1", "rulePack", payload);
        var verifier = new RsaSha256ArtifactVerifier("dev-1", RsaSha256ArtifactVerifier.ExportPublicKeyXml(rsa));

        var verified = verifier.Verify(encoded, "rulePack");
        var pack = LocalRulePackParser.Parse(verified, Now, "1.0.0.0");

        Assert.Equal("DEV-RULES", pack.PackId);
        Assert.Equal("Times New Roman", pack.BodyFontName);
        Assert.Equal(.3, pack.OrganLineMinRatio, 3);
        Assert.Equal(.55, pack.SubjectLineMaxRatio, 3);
        Assert.Single(pack.Corrections);
        Assert.False(pack.AcademicTypography.Enabled);
        Assert.Equal(AdvisoryProfileStatus.DisabledNotConfigured, pack.AcademicTypography.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.AdvisoryProfileNotConfigured,
            pack.AcademicTypography.DiagnosticCode);
    }

    [Fact]
    public void Legacy_v1_rule_pack_remains_compatible_and_cannot_enable_advisory_rules()
    {
        var root = RuleRoot(LocalRulePackParser.LegacySchema);
        root.Add(AdvisoryProfile(enabled: true));
        var payload = Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting));

        var pack = LocalRulePackParser.Parse(payload, Now, "1.0.0.0");

        Assert.Equal("1.0.0.0", pack.MinimumClientReleaseId);
        Assert.False(pack.AcademicTypography.Enabled);
        Assert.Equal(AdvisoryProfileStatus.DisabledLegacyV1, pack.AcademicTypography.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.LegacyRulePackHasNoAdvisoryProfile,
            pack.AcademicTypography.DiagnosticCode);
        Assert.Empty(pack.AcademicTypography.EnabledRuleCodes);
        Assert.Empty(pack.AcademicTypography.AutoFixRuleCodes);
    }

    [Fact]
    public void V2_rule_pack_parses_typed_academic_typography_profile()
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        root.Add(AdvisoryProfile());

        var pack = Parse(root);

        var profile = pack.AcademicTypography;
        Assert.True(profile.Enabled);
        Assert.Equal(AdvisoryProfileStatus.Enabled, profile.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.ProfileEnabled, profile.DiagnosticCode);
        Assert.Equal(1, profile.DetectorPolicyVersion);
        Assert.Equal(AcademicTypographyRuleCodes.All, profile.EnabledRuleCodes);
        Assert.Equal(new[]
        {
            AcademicTypographyRuleCodes.PaginationKeep,
            AcademicTypographyRuleCodes.PaginationWidow
        }, profile.AutoFixRuleCodes);
        Assert.Equal(.90d, profile.Thresholds.HeadingConfidenceMinimum, 3);
        Assert.Equal(.95d, profile.Thresholds.BodyConfidenceMinimum, 3);
        Assert.Equal(1, profile.Thresholds.CaptionMaxBlankParagraphs);
        Assert.Equal(1, profile.Thresholds.MathMinimumSignalCount);
        Assert.True(profile.IsRuleEnabled(AcademicTypographyRuleCodes.TableBooktabs));
        Assert.True(profile.IsAutoFixEnabled(AcademicTypographyRuleCodes.PaginationKeep));
        Assert.False(profile.IsAutoFixEnabled(AcademicTypographyRuleCodes.SectionStyle));
    }

    [Fact]
    public void V2_rule_pack_accepts_an_enabled_subset_but_only_safe_auto_fix_rules()
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        root.Add(AdvisoryProfile(
            enabledRules: new[]
            {
                AcademicTypographyRuleCodes.SectionStyle,
                AcademicTypographyRuleCodes.PaginationKeep
            },
            autoFixRules: new[] { AcademicTypographyRuleCodes.PaginationKeep }));

        var profile = Parse(root).AcademicTypography;

        Assert.True(profile.Enabled);
        Assert.Equal(2, profile.EnabledRuleCodes.Count);
        Assert.Single(profile.AutoFixRuleCodes);
    }

    [Fact]
    public void Explicitly_disabled_v2_profile_remains_typed_and_cannot_enable_a_rule_at_runtime()
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        root.Add(AdvisoryProfile(enabled: false));

        var profile = Parse(root).AcademicTypography;

        Assert.False(profile.Enabled);
        Assert.Equal(AdvisoryProfileStatus.DisabledByPolicy, profile.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.ProfileDisabledByPolicy, profile.DiagnosticCode);
        Assert.False(profile.IsRuleEnabled(AcademicTypographyRuleCodes.SectionStyle));
        Assert.False(profile.IsAutoFixEnabled(AcademicTypographyRuleCodes.PaginationKeep));
    }

    [Theory]
    [InlineData("duplicate-rule")]
    [InlineData("unknown-rule")]
    [InlineData("unsafe-auto-fix")]
    [InlineData("auto-fix-not-enabled")]
    [InlineData("threshold-out-of-range")]
    [InlineData("unknown-profile-field")]
    [InlineData("duplicate-profile")]
    [InlineData("missing-enabled-rules")]
    [InlineData("missing-auto-fix-rules")]
    [InlineData("missing-thresholds")]
    [InlineData("duplicate-enabled-rules")]
    [InlineData("duplicate-auto-fix-rules")]
    [InlineData("duplicate-thresholds")]
    [InlineData("invalid-enabled")]
    [InlineData("wrong-profile-code")]
    [InlineData("duplicate-auto-fix-rule")]
    public void Malformed_nested_advisory_payload_disables_the_whole_profile_but_keeps_core_rules(
        string malformedCase)
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        var advisory = AdvisoryProfile();
        root.Add(advisory);
        var profile = advisory.Element("profile")!;
        switch (malformedCase)
        {
            case "duplicate-rule":
                profile.Element("enabledRules")!.Add(
                    new XElement("rule", new XAttribute("code", AcademicTypographyRuleCodes.SectionStyle)));
                break;
            case "unknown-rule":
                profile.Element("enabledRules")!.Add(
                    new XElement("rule", new XAttribute("code", "LATEX-UNKNOWN")));
                break;
            case "unsafe-auto-fix":
                profile.Element("autoFixRules")!.Add(
                    new XElement("rule", new XAttribute("code", AcademicTypographyRuleCodes.SectionStyle)));
                break;
            case "auto-fix-not-enabled":
                profile.Element("enabledRules")!.Elements("rule")
                    .Single(item => (string?)item.Attribute("code") == AcademicTypographyRuleCodes.PaginationKeep)
                    .Remove();
                break;
            case "threshold-out-of-range":
                profile.Element("thresholds")!.SetAttributeValue("headingConfidenceMinimum", "1.01");
                break;
            case "unknown-profile-field":
                profile.Add(new XElement("unsupported"));
                break;
            case "duplicate-profile":
                advisory.Add(new XElement(profile));
                break;
            case "missing-enabled-rules":
                profile.Element("enabledRules")!.Remove();
                break;
            case "missing-auto-fix-rules":
                profile.Element("autoFixRules")!.Remove();
                break;
            case "missing-thresholds":
                profile.Element("thresholds")!.Remove();
                break;
            case "duplicate-enabled-rules":
                profile.Add(new XElement(profile.Element("enabledRules")!));
                break;
            case "duplicate-auto-fix-rules":
                profile.Add(new XElement(profile.Element("autoFixRules")!));
                break;
            case "duplicate-thresholds":
                profile.Add(new XElement(profile.Element("thresholds")!));
                break;
            case "invalid-enabled":
                profile.SetAttributeValue("enabled", "yes");
                break;
            case "wrong-profile-code":
                profile.SetAttributeValue("code", "OtherProfile");
                break;
            case "duplicate-auto-fix-rule":
                profile.Element("autoFixRules")!.Add(
                    new XElement("rule", new XAttribute("code", AcademicTypographyRuleCodes.PaginationKeep)));
                break;
            default:
                throw new InvalidOperationException(malformedCase);
        }

        var pack = Parse(root);

        Assert.Equal("Times New Roman", pack.BodyFontName);
        Assert.Single(pack.Corrections);
        Assert.False(pack.AcademicTypography.Enabled);
        Assert.Equal(AdvisoryProfileStatus.DisabledMalformed, pack.AcademicTypography.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.AdvisoryProfileMalformed,
            pack.AcademicTypography.DiagnosticCode);
        Assert.Empty(pack.AcademicTypography.EnabledRuleCodes);
        Assert.Empty(pack.AcademicTypography.AutoFixRuleCodes);
    }

    [Fact]
    public void Unsupported_nested_detector_policy_version_is_disabled_with_a_typed_diagnostic()
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        var advisory = AdvisoryProfile();
        advisory.Element("profile")!.SetAttributeValue("detectorPolicyVersion", "2");
        root.Add(advisory);

        var profile = Parse(root).AcademicTypography;

        Assert.False(profile.Enabled);
        Assert.Equal(2, profile.DetectorPolicyVersion);
        Assert.Equal(AdvisoryProfileStatus.DisabledUnsupportedPolicyVersion, profile.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.DetectorPolicyVersionUnsupported, profile.DiagnosticCode);
    }

    [Theory]
    [InlineData("unknown-root-attribute")]
    [InlineData("missing-format")]
    [InlineData("unknown-core-element")]
    [InlineData("unknown-format-attribute")]
    [InlineData("invalid-hidden-character")]
    [InlineData("duplicate-core-container")]
    [InlineData("invalid-pack-version")]
    [InlineData("invalid-core-number")]
    [InlineData("invalid-telex-regex")]
    [InlineData("invalid-validity-window")]
    public void Malformed_v2_core_payload_is_rejected_instead_of_salvaging_advisory_or_legal_rules(
        string malformedCase)
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        root.Add(AdvisoryProfile());
        switch (malformedCase)
        {
            case "unknown-root-attribute":
                root.SetAttributeValue("unsupported", "true");
                break;
            case "missing-format":
                root.Element("format")!.Remove();
                break;
            case "unknown-core-element":
                root.Add(new XElement("unsupported"));
                break;
            case "unknown-format-attribute":
                root.Element("format")!.SetAttributeValue("unsupported", "true");
                break;
            case "invalid-hidden-character":
                root.Element("hiddenCharacters")!.Element("character")!
                    .SetAttributeValue("codePoint", "D800");
                break;
            case "duplicate-core-container":
                root.Add(new XElement(root.Element("corrections")!));
                break;
            case "invalid-pack-version":
                root.SetAttributeValue("version", "release-two");
                break;
            case "invalid-core-number":
                root.Element("format")!.SetAttributeValue("a4WidthMm", "NaN");
                break;
            case "invalid-telex-regex":
                root.Element("telex")!.Add(new XElement("rule",
                    new XAttribute("pattern", "["), new XAttribute("replacement", "x")));
                break;
            case "invalid-validity-window":
                root.SetAttributeValue("notBeforeUtc", Now.AddDays(31).ToString("O"));
                break;
            default:
                throw new InvalidOperationException(malformedCase);
        }

        Assert.Throws<FormatException>(() => Parse(root));
    }

    [Fact]
    public void V2_minimum_client_release_compatibility_is_preserved()
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        root.SetAttributeValue("minimumClientReleaseId", "2.0.0.0");

        var exception = Assert.Throws<InvalidOperationException>(() => Parse(root, "1.9.9.9"));

        Assert.Equal("RULE_PACK_CLIENT_TOO_OLD", exception.Message);
    }

    [Fact]
    public void V2_rejects_non_version_compatibility_values_while_legacy_v1_behavior_remains_compatible()
    {
        var v2 = RuleRoot(LocalRulePackParser.Schema);
        v2.SetAttributeValue("minimumClientReleaseId", "release-z");
        Assert.Throws<FormatException>(() => Parse(v2, "release-a"));

        var v1 = RuleRoot(LocalRulePackParser.LegacySchema);
        v1.SetAttributeValue("minimumClientReleaseId", "release-a");
        var legacyPack = Parse(v1, "release-z");
        Assert.Equal(AdvisoryProfileStatus.DisabledLegacyV1, legacyPack.AcademicTypography.Status);
    }

    [Fact]
    public void Tampered_payload_is_rejected()
    {
        using var rsa = RSA.Create(2048);
        var encoded = Sign(rsa, "dev-1", "rulePack", RulePayload(Now.AddMinutes(-1), Now.AddDays(30)));
        var artifact = SignedArtifactCodec.Decode(encoded);
        artifact.Payload[artifact.Payload.Length / 2] ^= 1;
        var tampered = SignedArtifactCodec.Encode(artifact);
        var verifier = new RsaSha256ArtifactVerifier("dev-1", RsaSha256ArtifactVerifier.ExportPublicKeyXml(rsa));

        Assert.Throws<CryptographicException>(() => verifier.Verify(tampered, "rulePack"));
    }

    [Fact]
    public void Unsigned_advisory_opt_in_cannot_change_a_signed_disabled_profile()
    {
        using var rsa = RSA.Create(2048);
        var root = RuleRoot(LocalRulePackParser.Schema);
        root.Add(AdvisoryProfile(enabled: false));
        var encoded = Sign(rsa, "dev-1", "rulePack",
            Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting)));
        var artifact = SignedArtifactCodec.Decode(encoded);
        var tamperedRoot = XElement.Parse(Encoding.UTF8.GetString(artifact.Payload));
        tamperedRoot.Element("advisoryProfiles")!.Element("profile")!
            .SetAttributeValue("enabled", "true");
        var tampered = SignedArtifactCodec.Encode(new SignedArtifact(
            artifact.Kind, artifact.KeyId,
            Encoding.UTF8.GetBytes(tamperedRoot.ToString(SaveOptions.DisableFormatting)),
            artifact.Signature));
        var verifier = new RsaSha256ArtifactVerifier(
            "dev-1", RsaSha256ArtifactVerifier.ExportPublicKeyXml(rsa));

        Assert.Throws<CryptographicException>(() => verifier.Verify(tampered, "rulePack"));
    }

    [Theory]
    [InlineData("headingConfidenceMinimum", "0", true)]
    [InlineData("headingConfidenceMinimum", "1", true)]
    [InlineData("headingConfidenceMinimum", "-0.0001", false)]
    [InlineData("headingConfidenceMinimum", "1.0001", false)]
    [InlineData("headingConfidenceMinimum", "NaN", false)]
    [InlineData("headingConfidenceMinimum", "Infinity", false)]
    [InlineData("bodyConfidenceMinimum", "0", true)]
    [InlineData("bodyConfidenceMinimum", "1", true)]
    [InlineData("bodyConfidenceMinimum", "-Infinity", false)]
    [InlineData("captionMaxBlankParagraphs", "0", true)]
    [InlineData("captionMaxBlankParagraphs", "2", true)]
    [InlineData("captionMaxBlankParagraphs", "-1", false)]
    [InlineData("captionMaxBlankParagraphs", "3", false)]
    [InlineData("mathMinimumSignalCount", "1", true)]
    [InlineData("mathMinimumSignalCount", "10", true)]
    [InlineData("mathMinimumSignalCount", "0", false)]
    [InlineData("mathMinimumSignalCount", "11", false)]
    public void Advisory_threshold_boundaries_are_enforced_without_rejecting_the_core_pack(
        string attribute,
        string value,
        bool expectedEnabled)
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        var advisory = AdvisoryProfile();
        advisory.Element("profile")!.Element("thresholds")!.SetAttributeValue(attribute, value);
        root.Add(advisory);

        var pack = Parse(root);

        Assert.Equal(expectedEnabled, pack.AcademicTypography.Enabled);
        Assert.Equal(expectedEnabled ? AdvisoryProfileStatus.Enabled : AdvisoryProfileStatus.DisabledMalformed,
            pack.AcademicTypography.Status);
        Assert.Single(pack.Corrections);
    }

    [Fact]
    public void Unsupported_policy_version_does_not_mask_a_malformed_nested_profile()
    {
        var root = RuleRoot(LocalRulePackParser.Schema);
        var advisory = AdvisoryProfile();
        var profile = advisory.Element("profile")!;
        profile.SetAttributeValue("detectorPolicyVersion", "2");
        profile.Element("thresholds")!.SetAttributeValue("mathMinimumSignalCount", "0");
        root.Add(advisory);

        var parsed = Parse(root).AcademicTypography;

        Assert.Equal(AdvisoryProfileStatus.DisabledMalformed, parsed.Status);
        Assert.Equal(AdvisoryProfileDiagnosticCode.AdvisoryProfileMalformed, parsed.DiagnosticCode);
    }

    [Fact]
    public void Invalid_signed_line_shape_ratio_is_rejected_fail_closed()
    {
        var root = XElement.Parse(Encoding.UTF8.GetString(RulePayload(Now.AddMinutes(-1), Now.AddDays(30))));
        root.Element("format")!.SetAttributeValue("organLineMinRatio", "0.9");
        root.Element("format")!.SetAttributeValue("organLineMaxRatio", "0.5");

        Assert.Throws<FormatException>(() => LocalRulePackParser.Parse(
            Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting)), Now, "1.0.0.0"));
    }

    [Fact]
    public void Lease_is_device_release_feature_and_seven_day_bound()
    {
        var payload = Encoding.UTF8.GetBytes(new XElement("offlineLease",
            new XAttribute("schema", OfflineLeaseParser.Schema), new XAttribute("leaseId", "lease-1"),
            new XAttribute("subjectId", "dev-user"), new XAttribute("deviceThumbprint", "device-1"),
            new XAttribute("clientReleaseId", "1.0.0.0"), new XAttribute("issuedAtUtc", Now.ToString("O")),
            new XAttribute("notBeforeUtc", Now.AddMinutes(-1).ToString("O")),
            new XAttribute("expiresAtUtc", Now.AddDays(7).ToString("O")),
            new XElement("features", new XElement("feature", new XAttribute("code", "FORMAT_SCAN"))))
            .ToString(SaveOptions.DisableFormatting));
        var lease = OfflineLeaseParser.Parse(payload);
        var validator = new OfflineLeaseValidator();

        validator.Validate(lease, "device-1", "1.0.0.0", "FORMAT_SCAN", Now, Now.AddMinutes(-2));
        Assert.Throws<InvalidOperationException>(() =>
            validator.Validate(lease, "other-device", "1.0.0.0", "FORMAT_SCAN", Now));
        Assert.Throws<InvalidOperationException>(() =>
            validator.Validate(lease, "device-1", "1.0.0.0", "SPELLING_SCAN", Now));
    }

    [Fact]
    public void Lease_longer_than_seven_days_is_rejected()
    {
        var payload = Encoding.UTF8.GetBytes(new XElement("offlineLease",
            new XAttribute("schema", OfflineLeaseParser.Schema), new XAttribute("leaseId", "lease-1"),
            new XAttribute("subjectId", "dev-user"), new XAttribute("deviceThumbprint", "device-1"),
            new XAttribute("clientReleaseId", "1.0.0.0"), new XAttribute("issuedAtUtc", Now.ToString("O")),
            new XAttribute("notBeforeUtc", Now.ToString("O")),
            new XAttribute("expiresAtUtc", Now.AddDays(7).AddSeconds(1).ToString("O")),
            new XElement("features", new XElement("feature", new XAttribute("code", "FORMAT_SCAN"))))
            .ToString(SaveOptions.DisableFormatting));

        Assert.Throws<InvalidOperationException>(() => new OfflineLeaseValidator().Validate(
            OfflineLeaseParser.Parse(payload), "device-1", "1.0.0.0", "FORMAT_SCAN", Now));
    }

    internal static byte[] RulePayload(DateTimeOffset notBefore, DateTimeOffset expires,
        string schema = LocalRulePackParser.Schema)
    {
        var root = RuleRoot(schema);
        root.SetAttributeValue("notBeforeUtc", notBefore.ToString("O"));
        root.SetAttributeValue("expiresAtUtc", expires.ToString("O"));
        return Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting));
    }

    private static XElement RuleRoot(string schema) =>
        new XElement("rulePack",
            new XAttribute("schema", schema), new XAttribute("packId", "DEV-RULES"),
            new XAttribute("version", "1.0.0"), new XAttribute("notBeforeUtc", Now.AddMinutes(-1).ToString("O")),
            new XAttribute("expiresAtUtc", Now.AddDays(30).ToString("O")), new XAttribute("minimumClientReleaseId", "1.0.0.0"),
            new XElement("format", new XAttribute("a4WidthMm", "210"), new XAttribute("a4HeightMm", "297"),
                new XAttribute("topMinMm", "20"), new XAttribute("topMaxMm", "25"),
                new XAttribute("bottomMinMm", "20"), new XAttribute("bottomMaxMm", "25"),
                new XAttribute("leftMinMm", "30"), new XAttribute("leftMaxMm", "35"),
                new XAttribute("rightMinMm", "15"), new XAttribute("rightMaxMm", "20"),
                new XAttribute("mottoLineMinRatio", "0.8"), new XAttribute("mottoLineMaxRatio", "1.2"),
                new XAttribute("organLineMinRatio", "0.3"), new XAttribute("organLineMaxRatio", "0.55"),
                new XAttribute("subjectLineMinRatio", "0.3"), new XAttribute("subjectLineMaxRatio", "0.55"),
                new XAttribute("partyTitleLineMinRatio", "0.8"), new XAttribute("partyTitleLineMaxRatio", "1.2"),
                new XAttribute("bodyFontName", "Times New Roman")),
            new XElement("corrections", new XElement("correction", new XAttribute("wrong", "sát nhập"), new XAttribute("replacement", "sáp nhập"))),
            new XElement("telex"), new XElement("hiddenCharacters", new XElement("character", new XAttribute("codePoint", "200B"))))
        ;

    private static XElement AdvisoryProfile(
        bool enabled = true,
        IEnumerable<string>? enabledRules = null,
        IEnumerable<string>? autoFixRules = null) =>
        new XElement("advisoryProfiles",
            new XElement("profile",
                new XAttribute("code", AcademicTypographyAdvisoryProfile.ProfileCode),
                new XAttribute("enabled", enabled.ToString().ToLowerInvariant()),
                new XAttribute("detectorPolicyVersion", "1"),
                new XElement("enabledRules", (enabledRules ?? AcademicTypographyRuleCodes.All).Select(code =>
                    new XElement("rule", new XAttribute("code", code)))),
                new XElement("autoFixRules", (autoFixRules ?? new[]
                {
                    AcademicTypographyRuleCodes.PaginationKeep,
                    AcademicTypographyRuleCodes.PaginationWidow
                }).Select(code => new XElement("rule", new XAttribute("code", code)))),
                new XElement("thresholds",
                    new XAttribute("headingConfidenceMinimum", "0.90"),
                    new XAttribute("bodyConfidenceMinimum", "0.95"),
                    new XAttribute("captionMaxBlankParagraphs", "1"),
                    new XAttribute("mathMinimumSignalCount", "1"))));

    private static LocalRulePack Parse(XElement root, string clientReleaseId = "1.0.0.0") =>
        LocalRulePackParser.Parse(
            Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting)), Now, clientReleaseId);

    private static string Sign(RSA rsa, string keyId, string kind, byte[] payload)
    {
        var signature = rsa.SignData(payload, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return SignedArtifactCodec.Encode(new SignedArtifact(kind, keyId, payload, signature));
    }
}
