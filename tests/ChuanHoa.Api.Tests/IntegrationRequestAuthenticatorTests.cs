using System.Security.Cryptography;
using System.Text;
using ChuanHoa.Api.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;

namespace ChuanHoa.Api.Tests;

public sealed class IntegrationRequestAuthenticatorTests
{
    [Fact]
    public void Valid_signature_is_accepted_once_and_replay_is_rejected()
    {
        var clock = new FixedTimeProvider(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000));
        var authenticator = new IntegrationRequestAuthenticator(clock);
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ChuanHoa:AdminIntegration:Enabled"] = "true",
            ["ChuanHoa:AdminIntegration:ClientId"] = "bidding-admin",
            ["ChuanHoa:AdminIntegration:SharedSecret"] = new string('s', 32),
            ["ChuanHoa:AdminIntegration:ClockSkewSeconds"] = "300"
        }).Build();
        var request = BuildRequest("nonce-unique-001", "bidding-admin", new string('s', 32), clock.GetUtcNow()).Request;

        Assert.True(authenticator.TryAuthenticate(request, ReadOnlySpan<byte>.Empty, configuration));
        Assert.False(authenticator.TryAuthenticate(request, ReadOnlySpan<byte>.Empty, configuration));
    }

    [Fact]
    public void Expired_or_wrong_signature_is_rejected_without_consuming_nonce()
    {
        var clock = new FixedTimeProvider(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000));
        var authenticator = new IntegrationRequestAuthenticator(clock);
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ChuanHoa:AdminIntegration:Enabled"] = "true",
            ["ChuanHoa:AdminIntegration:ClientId"] = "bidding-admin",
            ["ChuanHoa:AdminIntegration:SharedSecret"] = new string('s', 32)
        }).Build();
        var expired = BuildRequest("nonce-expired-001", "bidding-admin", new string('s', 32), DateTimeOffset.FromUnixTimeSeconds(1_699_000_000)).Request;
        Assert.False(authenticator.TryAuthenticate(expired, ReadOnlySpan<byte>.Empty, configuration));
        var valid = BuildRequest("nonce-expired-001", "bidding-admin", new string('s', 32), clock.GetUtcNow()).Request;
        Assert.True(authenticator.TryAuthenticate(valid, ReadOnlySpan<byte>.Empty, configuration));
    }

    [Fact]
    public void Post_signature_covers_exact_json_body_and_tampering_is_rejected()
    {
        var clock = new FixedTimeProvider(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000));
        var authenticator = new IntegrationRequestAuthenticator(clock);
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ChuanHoa:AdminIntegration:Enabled"] = "true",
            ["ChuanHoa:AdminIntegration:ClientId"] = "bidding-admin",
            ["ChuanHoa:AdminIntegration:SharedSecret"] = new string('s', 32)
        }).Build();
        var body = Encoding.UTF8.GetBytes("{\"userId\":\"u-1\",\"durationDays\":30}");
        var request = BuildRequest("nonce-post-001-xx", "bidding-admin", new string('s', 32), clock.GetUtcNow(), "POST", "/v1/admin/integration/entitlements/extend", body).Request;
        Assert.True(authenticator.TryAuthenticate(request, body, configuration));

        var tampered = BuildRequest("nonce-post-002-xx", "bidding-admin", new string('s', 32), clock.GetUtcNow(), "POST", "/v1/admin/integration/entitlements/extend", body).Request;
        Assert.False(authenticator.TryAuthenticate(tampered, Encoding.UTF8.GetBytes("{\"userId\":\"u-2\",\"durationDays\":30}"), configuration));
    }

    [Fact]
    public void Disabled_integration_revokes_otherwise_valid_client_credentials()
    {
        var clock = new FixedTimeProvider(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000));
        var authenticator = new IntegrationRequestAuthenticator(clock);
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ChuanHoa:AdminIntegration:Enabled"] = "false",
            ["ChuanHoa:AdminIntegration:ClientId"] = "bidding-admin",
            ["ChuanHoa:AdminIntegration:SharedSecret"] = new string('s', 32)
        }).Build();
        var request = BuildRequest("nonce-revoked-001", "bidding-admin", new string('s', 32), clock.GetUtcNow()).Request;
        Assert.False(authenticator.TryAuthenticate(request, ReadOnlySpan<byte>.Empty, configuration));
    }

    [Fact]
    public void Missing_integration_toggle_is_fail_closed()
    {
        var clock = new FixedTimeProvider(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000));
        var authenticator = new IntegrationRequestAuthenticator(clock);
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ChuanHoa:AdminIntegration:ClientId"] = "bidding-admin",
            ["ChuanHoa:AdminIntegration:SharedSecret"] = new string('s', 32)
        }).Build();
        var request = BuildRequest("nonce-missing-toggle", "bidding-admin", new string('s', 32), clock.GetUtcNow()).Request;
        Assert.False(authenticator.TryAuthenticate(request, ReadOnlySpan<byte>.Empty, configuration));
    }

    private static DefaultHttpContext BuildRequest(string nonce, string client, string secret, DateTimeOffset timestamp, string method = "GET", string path = "/v1/admin/integration/capabilities", byte[]? body = null)
    {
        var context = new DefaultHttpContext();
        context.Request.Method = method;
        context.Request.Path = path;
        var timestampText = timestamp.ToUnixTimeSeconds().ToString();
        var bodyHash = Convert.ToHexString(SHA256.HashData(body ?? Array.Empty<byte>())).ToLowerInvariant();
        var canonical = string.Join('\n', method, context.Request.Path, timestampText, nonce, bodyHash);
        var signature = Convert.ToHexString(HMACSHA256.HashData(Encoding.UTF8.GetBytes(secret), Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
        context.Request.Headers["X-Integration-Client"] = client;
        context.Request.Headers["X-Integration-Timestamp"] = timestampText;
        context.Request.Headers["X-Integration-Nonce"] = nonce;
        context.Request.Headers["X-Integration-Signature"] = signature;
        return context;
    }

    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => now;
    }
}
