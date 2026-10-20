using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.Client.Core.Tests;

public sealed class BundledRulePackTests
{
    [Fact]
    public void Bundled_rules_are_complete_and_have_no_usage_expiry()
    {
        var pack = BundledRulePack.Load("1.0.0.178");
        Assert.Equal("BUNDLED-FREE-ND30-V2", pack.PackId);
        Assert.Equal(DateTimeOffset.MinValue, pack.NotBeforeUtc);
        Assert.Equal(DateTimeOffset.MaxValue, pack.ExpiresAtUtc);
        Assert.True(pack.Lexicon.Count > 1000);
        Assert.NotEmpty(pack.Corrections);
        Assert.NotEmpty(pack.Capitalizations);
        Assert.NotEmpty(pack.DocumentTypeAbbreviations);
        Assert.Equal(210, pack.A4WidthMm);
        Assert.Equal(297, pack.A4HeightMm);
        Assert.False(pack.AcademicTypography.Enabled);
    }

    [Fact]
    public void Bundled_rules_still_reject_incompatible_clients() =>
        Assert.Throws<InvalidOperationException>(() => BundledRulePack.Load("0.9.0.0"));
}
