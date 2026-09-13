using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;

namespace ChuanHoa.Api.Security;

public sealed class IntegrationRequestAuthenticator(TimeProvider timeProvider)
{
    private readonly ConcurrentDictionary<string, long> _seenNonces = new(StringComparer.Ordinal);

    public bool TryAuthenticate(HttpRequest request, ReadOnlySpan<byte> body, IConfiguration configuration)
    {
        if (!bool.TryParse(configuration["ChuanHoa:AdminIntegration:Enabled"], out var enabled) || !enabled)
            return false;
        var expectedClient = configuration["ChuanHoa:AdminIntegration:ClientId"];
        var secret = configuration["ChuanHoa:AdminIntegration:SharedSecret"];
        if (string.IsNullOrWhiteSpace(expectedClient) || string.IsNullOrWhiteSpace(secret) || secret.Length < 32)
            return false;
        var client = request.Headers["X-Integration-Client"].ToString();
        var timestampText = request.Headers["X-Integration-Timestamp"].ToString();
        var nonce = request.Headers["X-Integration-Nonce"].ToString();
        var signature = request.Headers["X-Integration-Signature"].ToString();
        if (!string.Equals(client, expectedClient, StringComparison.Ordinal)
            || nonce.Length is < 16 or > 128
            || !long.TryParse(timestampText, out var timestamp)
            || string.IsNullOrWhiteSpace(signature))
            return false;
        var now = timeProvider.GetUtcNow().ToUnixTimeSeconds();
        var skew = int.TryParse(configuration["ChuanHoa:AdminIntegration:ClockSkewSeconds"], out var configuredSkew)
            ? Math.Clamp(configuredSkew, 30, 900)
            : 300;
        if (Math.Abs(now - timestamp) > skew) return false;
        var bodyHash = Convert.ToHexString(SHA256.HashData(body)).ToLowerInvariant();
        var canonical = string.Join('\n', request.Method.ToUpperInvariant(), request.Path, timestampText, nonce, bodyHash);
        var expected = Convert.ToHexString(HMACSHA256.HashData(Encoding.UTF8.GetBytes(secret), Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
        var validSignature = CryptographicOperations.FixedTimeEquals(
            Encoding.UTF8.GetBytes(expected),
            Encoding.UTF8.GetBytes(signature.Trim().ToLowerInvariant()));
        if (!validSignature) return false;
        var nonceKey = $"{client}:{nonce}";
        if (!_seenNonces.TryAdd(nonceKey, now + skew)) return false;
        foreach (var item in _seenNonces.Where(item => item.Value < now).ToArray())
            _seenNonces.TryRemove(item.Key, out _);
        return true;
    }
}
