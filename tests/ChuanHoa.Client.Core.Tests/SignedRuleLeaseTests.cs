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

    internal static byte[] RulePayload(DateTimeOffset notBefore, DateTimeOffset expires) =>
        Encoding.UTF8.GetBytes(new XElement("rulePack",
            new XAttribute("schema", LocalRulePackParser.Schema), new XAttribute("packId", "DEV-RULES"),
            new XAttribute("version", "1.0.0"), new XAttribute("notBeforeUtc", notBefore.ToString("O")),
            new XAttribute("expiresAtUtc", expires.ToString("O")), new XAttribute("minimumClientReleaseId", "1.0.0.0"),
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
            .ToString(SaveOptions.DisableFormatting));

    private static string Sign(RSA rsa, string keyId, string kind, byte[] payload)
    {
        var signature = rsa.SignData(payload, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return SignedArtifactCodec.Encode(new SignedArtifact(kind, keyId, payload, signature));
    }
}
