using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using ChuanHoa.Api.Development;
using ChuanHoa.Client.Core.Licensing;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Security;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace ChuanHoa.Api.Tests;

public sealed class DevelopmentArtifactIssuerTests
{
    [Fact]
    public async Task Development_issuer_creates_verifiable_seven_day_lease_and_rule_pack()
    {
        await using var fixture = await DevelopmentIssuerFixture.Create(new Dictionary<string, string?>
        {
            ["ChuanHoa:AcademicTypographyAdvisory:Enabled"] = "true",
            ["ChuanHoa:AcademicTypographyAdvisory:HeadingConfidenceMinimum"] = "0.91",
            ["ChuanHoa:AcademicTypographyAdvisory:BodyConfidenceMinimum"] = "0.96",
            ["ChuanHoa:AcademicTypographyAdvisory:CaptionMaxBlankParagraphs"] = "2",
            ["ChuanHoa:AcademicTypographyAdvisory:MathMinimumSignalCount"] = "3"
        });

        var response = fixture.Issuer.Issue(new DevelopmentBootstrapRequest("device-test", "1.0.0.0"));
        var verifier = new RsaSha256ArtifactVerifier(
            DevelopmentArtifactIssuer.KeyId,
            RsaSha256ArtifactVerifier.ExportPublicKeyXml(fixture.Rsa));
        var lease = OfflineLeaseParser.Parse(verifier.Verify(response.Lease, DevelopmentArtifactIssuer.LeaseKind));
        var pack = LocalRulePackParser.Parse(
            verifier.Verify(response.RulePack, DevelopmentArtifactIssuer.RulePackKind),
            response.ServerTimeUtc,
            "1.0.0.0");

        Assert.Equal(TimeSpan.FromDays(7), lease.ExpiresAtUtc - lease.IssuedAtUtc);
        Assert.Contains("FORMAT_SCAN", lease.Features);
        Assert.Contains("SPELLING_SCAN", lease.Features);
        Assert.Contains("DOCUMENT_TOOLS", lease.Features);
        Assert.Contains("AUTOFIX", lease.Features);
        Assert.Equal("device-test", lease.DeviceThumbprint);
        Assert.Equal("1.0.0.0", pack.MinimumClientReleaseId);
        Assert.Single(pack.Corrections);
        Assert.Single(pack.TelexRules);
        Assert.Equal(1001, pack.Lexicon.Count);
        Assert.True(pack.AcademicTypography.Enabled);
        Assert.Equal(AcademicTypographyRuleCodes.All, pack.AcademicTypography.EnabledRuleCodes);
        Assert.Equal(new[]
        {
            AcademicTypographyRuleCodes.PaginationKeep,
            AcademicTypographyRuleCodes.PaginationWidow
        }, pack.AcademicTypography.AutoFixRuleCodes);
        Assert.Equal(.91d, pack.AcademicTypography.Thresholds.HeadingConfidenceMinimum, 3);
        Assert.Equal(.96d, pack.AcademicTypography.Thresholds.BodyConfidenceMinimum, 3);
        Assert.Equal(2, pack.AcademicTypography.Thresholds.CaptionMaxBlankParagraphs);
        Assert.Equal(3, pack.AcademicTypography.Thresholds.MathMinimumSignalCount);
    }

    [Fact]
    public async Task Development_advisory_profile_is_disabled_by_default_but_core_pack_remains_valid()
    {
        await using var fixture = await DevelopmentIssuerFixture.Create();
        var response = fixture.Issuer.Issue(new DevelopmentBootstrapRequest("device-test", "1.0.0.0"));
        var verifier = new RsaSha256ArtifactVerifier(
            DevelopmentArtifactIssuer.KeyId,
            RsaSha256ArtifactVerifier.ExportPublicKeyXml(fixture.Rsa));

        var pack = LocalRulePackParser.Parse(
            verifier.Verify(response.RulePack, DevelopmentArtifactIssuer.RulePackKind),
            response.ServerTimeUtc,
            "1.0.0.0");

        Assert.False(pack.AcademicTypography.Enabled);
        Assert.Equal(AdvisoryProfileStatus.DisabledByPolicy, pack.AcademicTypography.Status);
        Assert.Single(pack.Corrections);
        Assert.Equal(1001, pack.Lexicon.Count);
    }

    [Theory]
    [InlineData("HeadingConfidenceMinimum", "1.01")]
    [InlineData("BodyConfidenceMinimum", "NaN")]
    [InlineData("CaptionMaxBlankParagraphs", "3")]
    [InlineData("MathMinimumSignalCount", "0")]
    [InlineData("DetectorPolicyVersion", "2")]
    public async Task Development_issuer_refuses_to_sign_invalid_advisory_configuration(string name, string value)
    {
        await using var fixture = await DevelopmentIssuerFixture.Create(new Dictionary<string, string?>
        {
            ["ChuanHoa:AcademicTypographyAdvisory:" + name] = value
        });

        Assert.Throws<InvalidOperationException>(() =>
            fixture.Issuer.Issue(new DevelopmentBootstrapRequest("device-test", "1.0.0.0")));
    }

    [Fact]
    public async Task Production_environment_cannot_enable_the_development_issuer_or_advisory_profile()
    {
        await using var fixture = await DevelopmentIssuerFixture.Create(new Dictionary<string, string?>
        {
            ["ChuanHoa:AcademicTypographyAdvisory:Enabled"] = "true"
        }, "Production");

        Assert.False(fixture.Issuer.IsEnabled);
        Assert.Throws<InvalidOperationException>(() =>
            fixture.Issuer.Issue(new DevelopmentBootstrapRequest("device-test", "1.0.0.0")));
    }

    [Fact]
    public async Task Development_issuer_refuses_to_sign_a_malformed_client_release_id()
    {
        await using var fixture = await DevelopmentIssuerFixture.Create();

        Assert.Throws<ArgumentException>(() =>
            fixture.Issuer.Issue(new DevelopmentBootstrapRequest("device-test", "release-not-a-version")));
    }

    [Fact]
    public async Task Development_issuer_is_disabled_without_explicit_flag()
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions { EnvironmentName = "Development" });
        builder.Services.AddSingleton(TimeProvider.System);
        builder.Services.AddSingleton<DevelopmentArtifactIssuer>();
        await using var app = builder.Build();
        var issuer = app.Services.GetRequiredService<DevelopmentArtifactIssuer>();

        Assert.False(issuer.IsEnabled);
        Assert.Throws<InvalidOperationException>(() =>
            issuer.Issue(new DevelopmentBootstrapRequest("device-test", "1.0.0.0")));
    }

    private static string PrivateKeyXml(RSA rsa)
    {
        var key = rsa.ExportParameters(true);
        return new XElement("RSAKeyValue",
            Element("Modulus", key.Modulus), Element("Exponent", key.Exponent), Element("P", key.P),
            Element("Q", key.Q), Element("DP", key.DP), Element("DQ", key.DQ),
            Element("InverseQ", key.InverseQ), Element("D", key.D))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static XElement Element(string name, byte[]? value) =>
        new(name, Convert.ToBase64String(value ?? throw new InvalidOperationException(name + " is missing.")));

    private sealed class DevelopmentIssuerFixture : IAsyncDisposable
    {
        private DevelopmentIssuerFixture(string directory, RSA rsa, WebApplication app,
            DevelopmentArtifactIssuer issuer)
        {
            DirectoryPath = directory;
            Rsa = rsa;
            App = app;
            Issuer = issuer;
        }

        private string DirectoryPath { get; }
        private WebApplication App { get; }
        public RSA Rsa { get; }
        public DevelopmentArtifactIssuer Issuer { get; }

        public static async Task<DevelopmentIssuerFixture> Create(
            IReadOnlyDictionary<string, string?>? overrides = null,
            string environmentName = "Development")
        {
            var directory = Path.Combine(Path.GetTempPath(), "ChuanHoaIssuerTest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var rsa = RSA.Create(2048);
            try
            {
                var privateKeyPath = Path.Combine(directory, "key.xml");
                var dictionaryPath = Path.Combine(directory, "dictionary.json");
                var lexiconDirectory = Path.Combine(directory, "lexicon");
                Directory.CreateDirectory(lexiconDirectory);
                await File.WriteAllLinesAsync(Path.Combine(lexiconDirectory, "vi-test.dic"),
                    new[] { "1001" }.Concat(Enumerable.Range(0, 1001).Select(index => "từ" + index)));
                await File.WriteAllTextAsync(privateKeyPath, PrivateKeyXml(rsa));
                await File.WriteAllTextAsync(dictionaryPath, JsonSerializer.Serialize(new Dictionary<string, string>
                {
                    ["sát nhập"] = "sáp nhập"
                }));

                var builder = WebApplication.CreateBuilder(new WebApplicationOptions
                {
                    EnvironmentName = environmentName,
                    ContentRootPath = directory
                });
                var configuration = new Dictionary<string, string?>
                {
                    ["ChuanHoa:EnableDevelopmentBootstrap"] = "true",
                    ["ChuanHoa:DevelopmentSigningKeyPath"] = privateKeyPath,
                    ["ChuanHoa:DevelopmentRuleDictionaryPath"] = dictionaryPath,
                    ["ChuanHoa:DevelopmentRuleLexiconDirectory"] = lexiconDirectory
                };
                if (overrides != null)
                {
                    foreach (var item in overrides) configuration[item.Key] = item.Value;
                }
                builder.Configuration.AddInMemoryCollection(configuration);
                builder.Services.AddSingleton(TimeProvider.System);
                builder.Services.AddSingleton<DevelopmentArtifactIssuer>();
                var app = builder.Build();
                return new DevelopmentIssuerFixture(
                    directory, rsa, app, app.Services.GetRequiredService<DevelopmentArtifactIssuer>());
            }
            catch
            {
                rsa.Dispose();
                Directory.Delete(directory, true);
                throw;
            }
        }

        public async ValueTask DisposeAsync()
        {
            await App.DisposeAsync();
            Rsa.Dispose();
            Directory.Delete(DirectoryPath, true);
        }
    }
}
