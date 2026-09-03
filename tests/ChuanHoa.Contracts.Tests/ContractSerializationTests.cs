using System.Text.Json;
using ChuanHoa.Contracts;

namespace ChuanHoa.Contracts.Tests;

public sealed class ContractSerializationTests
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web);

    [Fact]
    public void Risk_tier_serializes_as_stable_string()
    {
        var json = JsonSerializer.Serialize(RiskTier.ReportOnly, Options);
        Assert.Equal("\"ReportOnly\"", json);
    }

    [Fact]
    public void Fix_plan_rejects_unknown_operation_enum()
    {
        const string json = "\"RunScript\"";
        Assert.Throws<JsonException>(() => JsonSerializer.Deserialize<FixOperationType>(json, Options));
    }

    [Fact]
    public void Finding_anchor_serializes_machine_readable_exact_range()
    {
        var finding = new Finding(
            "ND30-FONT-01",
            FindingStatus.Fail,
            FindingSeverity.Error,
            RiskTier.ReportOnly,
            "Sai phông chữ",
            "Hiện tại Arial; yêu cầu Times New Roman.",
            7,
            ComponentRole.BodyText,
            "rules-2026.09",
            null,
            "finding-01",
            "Times New Roman",
            "NĐ 30/2020/NĐ-CP, Phụ lục I",
            new FindingAnchor(
                FindingAnchorKind.TextSpan,
                "MainTextStory",
                7,
                4,
                5,
                "Arial"));

        var json = JsonSerializer.Serialize(finding, Options);
        var roundTrip = JsonSerializer.Deserialize<Finding>(json, Options);

        Assert.NotNull(roundTrip);
        Assert.Equal("finding-01", roundTrip.FindingId);
        Assert.Equal(FindingAnchorKind.TextSpan, roundTrip.Anchor!.Kind);
        Assert.Equal(4, roundTrip.Anchor.StartOffset);
        Assert.Equal(5, roundTrip.Anchor.Length);
        Assert.Equal("Arial", roundTrip.Anchor.ExpectedText);
    }
}
