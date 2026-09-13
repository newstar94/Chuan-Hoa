using System.Security.Cryptography;
using System.Text;
using System.Collections.Concurrent;
using ChuanHoa.Contracts.Integration;

namespace ChuanHoa.Api.Security;

public sealed class DisabledIntegrationReplayStore : IIntegrationReplayStore
{
    public Task<bool> TryClaimAsync(string clientId, string nonce, DateTimeOffset expiresAtUtc, CancellationToken cancellationToken)
        => Task.FromResult(false);
}

public sealed class InMemoryIntegrationReplayStore : IIntegrationReplayStore
{
    private readonly ConcurrentDictionary<string, DateTimeOffset> _claims = new(StringComparer.Ordinal);

    public Task<bool> TryClaimAsync(string clientId, string nonce, DateTimeOffset expiresAtUtc, CancellationToken cancellationToken)
    {
        var now = DateTimeOffset.UtcNow;
        foreach (var item in _claims)
            if (item.Value <= now) _claims.TryRemove(item.Key, out _);
        return Task.FromResult(_claims.TryAdd($"{clientId}:{nonce}", expiresAtUtc));
    }
}

public sealed class IntegrationRequestAuthenticator(TimeProvider timeProvider, IIntegrationReplayStore replayStore)
{
    public async Task<bool> TryAuthenticateAsync(HttpRequest request, ReadOnlyMemory<byte> body, IConfiguration configuration, CancellationToken cancellationToken = default)
    {
        var localDevelopment = IsLocalDevelopment(configuration);
        if (!localDevelopment && (!bool.TryParse(configuration["ChuanHoa:AdminIntegration:Enabled"], out var enabled) || !enabled))
            return false;
        var expectedClient = configuration["ChuanHoa:AdminIntegration:ClientId"] ?? "bidding-admin";
        var secret = configuration["ChuanHoa:AdminIntegration:SharedSecret"];
        if (localDevelopment && string.IsNullOrWhiteSpace(secret)) secret = LocalBootstrapSecret();
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

    private static bool IsLocalDevelopment(IConfiguration configuration)
        => configuration["ASPNETCORE_ENVIRONMENT"]?.Equals("Development", StringComparison.OrdinalIgnoreCase) == true
            && Uri.TryCreate(configuration["ChuanHoa:AdminIntegration:BaseUrl"] ?? "http://127.0.0.1:5206", UriKind.Absolute, out var uri)
            && uri.Scheme == Uri.UriSchemeHttp && (uri.Host == "127.0.0.1" || uri.Host == "localhost" || uri.Host == "::1");

    private static string LocalBootstrapSecret()
    {
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var path = Path.Combine(root, "ChuanHoa", "Development", "admin-integration-secret.txt");
        try
        {
            if (File.Exists(path))
            {
                var existing = File.ReadAllText(path).Trim();
                if (existing.Length >= 32) return existing;
            }
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var value = Convert.ToBase64String(RandomNumberGenerator.GetBytes(48));
            try
            {
                using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
                using var writer = new StreamWriter(stream);
                writer.Write(value);
                return value;
            }
            catch (IOException)
            {
                return File.Exists(path) ? File.ReadAllText(path).Trim() : string.Empty;
            }
        }
        catch (IOException) { return string.Empty; }
    }
}
