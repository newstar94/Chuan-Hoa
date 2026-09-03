using System.Text.Json;

namespace ChuanHoa.Contracts.Tests;

public sealed class OpenApiContractTests
{
    [Fact]
    public void OpenApi_snapshot_is_versioned_environment_neutral_and_contains_error_contract()
    {
        var contractPath = Path.Combine(
            AppContext.BaseDirectory,
            "Contracts",
            "chuanhoa-api.v1.json");
        using var stream = File.OpenRead(contractPath);
        using var document = JsonDocument.Parse(stream);
        var root = document.RootElement;

        Assert.Equal("3.1.1", root.GetProperty("openapi").GetString());
        Assert.Equal("Chuẩn Hóa API", root.GetProperty("info").GetProperty("title").GetString());
        Assert.Equal("v1", root.GetProperty("info").GetProperty("version").GetString());
        Assert.False(root.TryGetProperty("servers", out _));

        var paths = root.GetProperty("paths");
        Assert.True(paths.TryGetProperty("/health", out _));
        Assert.True(paths.TryGetProperty("/health/ready", out _));

        var errorSchema = root
            .GetProperty("components")
            .GetProperty("schemas")
            .GetProperty("ApiErrorResponse");
        var required = errorSchema
            .GetProperty("required")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("schema", required);
        Assert.Contains("code", required);
        Assert.Contains("recovery", required);
        Assert.Contains("correlationId", required);
    }
}
