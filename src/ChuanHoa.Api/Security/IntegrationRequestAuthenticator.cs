using System.Security.Cryptography;
using System.Text;
using ChuanHoa.Contracts.Integration;

namespace ChuanHoa.Api.Security;

public sealed class DisabledIntegrationReplayStore : IIntegrationReplayStore
{
    public Task<bool> TryClaimAsync(string clientId, string nonce, DateTimeOffset expiresAtUtc, CancellationToken cancellationToken)
        => Task.FromResult(false);
}

public sealed class IntegrationRequestAuthenticator(TimeProvider timeProvider, IIntegrationReplayStore replayStore)
{
    public async Task<bool> TryAuthenticateAsync(HttpRequest request, ReadOnlyMemory<byte> body, IConfiguration configuration, CancellationToken cancellationToken = default)
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
        var bodyHash = Convert.ToHexString(SHA256.HashData(body.Span)).ToLowerInvariant();
        var canonical = string.Join('\n', request.Method.ToUpperInvariant(), request.PathBase + request.Path + request.QueryString, timestampText, nonce, bodyHash);
        var expected = Convert.ToHexString(HMACSHA256.HashData(Encoding.UTF8.GetBytes(secret), Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
        var validSignature = CryptographicOperations.FixedTimeEquals(
            Encoding.UTF8.GetBytes(expected),
            Encoding.UTF8.GetBytes(signature.Trim().ToLowerInvariant()));
        if (!validSignature) return false;
        return await replayStore.TryClaimAsync(client, nonce, DateTimeOffset.FromUnixTimeSeconds(timestamp + skew), cancellationToken);
    }
}
