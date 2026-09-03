using System.Text.Json;
using System.Text.Json.Serialization;
using ChuanHoa.Contracts;

namespace ChuanHoa.Rules;

public sealed class CanonicalRuleReleaseParser
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        AllowTrailingCommas = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = false,
        ReadCommentHandling = JsonCommentHandling.Disallow,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        MaxDepth = 64,
        Converters =
        {
            new JsonStringEnumConverter()
        }
    };

    public CanonicalRuleRelease Parse(ReadOnlySpan<byte> utf8Json)
    {
        if (utf8Json.IsEmpty)
        {
            throw new JsonException("Canonical rule release payload is empty.");
        }

        var release = JsonSerializer.Deserialize<CanonicalRuleRelease>(utf8Json, SerializerOptions);
        return release ?? throw new JsonException("Canonical rule release payload resolved to null.");
    }

    public CanonicalRuleRelease Parse(Stream utf8Json)
    {
        ArgumentNullException.ThrowIfNull(utf8Json);
        var release = JsonSerializer.Deserialize<CanonicalRuleRelease>(utf8Json, SerializerOptions);
        return release ?? throw new JsonException("Canonical rule release payload resolved to null.");
    }
}
