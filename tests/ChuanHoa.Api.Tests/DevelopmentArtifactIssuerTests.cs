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
        var directory = Path.Combine(Path.GetTempPath(), "ChuanHoaIssuerTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            using var rsa = RSA.Create(2048);
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
                EnvironmentName = "Development",
                ContentRootPath = directory
            });
            builder.Configuration.AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["ChuanHoa:EnableDevelopmentBootstrap"] = "true",
                ["ChuanHoa:DevelopmentSigningKeyPath"] = privateKeyPath,
                ["ChuanHoa:DevelopmentRuleDictionaryPath"] = dictionaryPath,
                ["ChuanHoa:DevelopmentRuleLexiconDirectory"] = lexiconDirectory
            });
            builder.Services.AddSingleton(TimeProvider.System);
            builder.Services.AddSingleton<DevelopmentArtifactIssuer>();
            await using var app = builder.Build();
            var issuer = app.Services.GetRequiredService<DevelopmentArtifactIssuer>();

            var response = issuer.Issue(new DevelopmentBootstrapRequest("device-test", "1.0.0.0"));
            var verifier = new RsaSha256ArtifactVerifier(
                DevelopmentArtifactIssuer.KeyId,
                RsaSha256ArtifactVerifier.ExportPublicKeyXml(rsa));
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
            Assert.Single(pack.Corrections);
            Assert.Single(pack.TelexRules);
            Assert.Equal(1001, pack.Lexicon.Count);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
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
}
