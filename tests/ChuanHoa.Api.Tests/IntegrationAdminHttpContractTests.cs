#pragma warning disable ASPDEPR004, ASPDEPR008
using System.Net;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ChuanHoa.Api.Controllers;
using ChuanHoa.Api.Security;
using ChuanHoa.Infrastructure.Admin;
using ChuanHoa.Contracts.Integration;
using ChuanHoa.Contracts;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace ChuanHoa.Api.Tests;

public sealed class IntegrationAdminHttpContractTests
{
    private const string ClientId = "bidding-admin";
    private static readonly string Secret = new('s', 32);

    [Fact]
    public async Task Signed_capabilities_request_reaches_controller_over_http()
    {
        using var server = CreateServer();
        using var client = server.CreateClient();
        using var request = SignedRequest(HttpMethod.Get, "/v1/admin/integration/capabilities", "nonce-http-cap-001", ReadOnlySpan<byte>.Empty);

        using var response = await client.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var payload = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("chuanhoa.admin.integration.v1", payload.GetProperty("schema").GetString());
        Assert.Equal("chuan-hoa", payload.GetProperty("application").GetString());
    }

    [Fact]
    public async Task Signed_mutation_body_is_verified_and_dispatched_once()
    {
        using var server = CreateServer();
        using var client = server.CreateClient();
        var correlation = Guid.NewGuid();
        var body = JsonSerializer.SerializeToUtf8Bytes(new
        {
            userId = Guid.NewGuid(), productId = Guid.NewGuid(), featureCodes = new[] { "feature.a" },
            durationDays = 30, reason = "contract-test", actorId = "actor-1", correlationId = correlation
        });
        using var request = SignedRequest(HttpMethod.Post, "/v1/admin/integration/entitlements/extend", "nonce-http-mut-001", body);
        request.Headers.Add("Idempotency-Key", "http-contract-key-0001");
        request.Content = new ByteArrayContent(body);
        request.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

        using var response = await client.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal(1, server.Services.GetRequiredService<FakeStore>().MutationCalls);
    }

    [Fact]
    public async Task Signed_mutation_preserves_exact_unicode_body_bytes()
    {
        using var server = CreateServer();
        using var client = server.CreateClient();
        var body = Encoding.UTF8.GetBytes($"{{\"userId\":\"{Guid.NewGuid():D}\",\"productId\":\"{Guid.NewGuid():D}\",\"featureCodes\":[\"tính.năng\"],\"durationDays\":30,\"reason\":\"Gia hạn hỗ trợ Đà Nẵng\",\"actorId\":\"actor-1\",\"correlationId\":\"{Guid.NewGuid():D}\"}}");
        using var request = SignedRequest(HttpMethod.Post, "/v1/admin/integration/entitlements/extend", "nonce-http-unicode-001", body);
        request.Headers.Add("Idempotency-Key", "http-unicode-key-0001");
        request.Content = new ByteArrayContent(body);
        request.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

        using var response = await client.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal(1, server.Services.GetRequiredService<FakeStore>().MutationCalls);
    }

    [Fact]
    public async Task Signed_query_is_accepted_and_tampered_query_is_rejected()
    {
        using var server = CreateServer();
        using var client = server.CreateClient();
        using var valid = SignedRequest(HttpMethod.Get, "/v1/admin/integration/accounts?search=%C4%90%C3%A0%20N%E1%BA%B5ng&page=2&pageSize=25", "nonce-http-query-001", ReadOnlySpan<byte>.Empty);
        using var accepted = await client.SendAsync(valid);
        Assert.Equal(HttpStatusCode.OK, accepted.StatusCode);

        using var tampered = SignedRequest(HttpMethod.Get, "/v1/admin/integration/accounts?page=2", "nonce-http-query-002", ReadOnlySpan<byte>.Empty);
        tampered.RequestUri = new Uri("/v1/admin/integration/accounts?page=3", UriKind.Relative);
        using var rejected = await client.SendAsync(tampered);
        Assert.Equal(HttpStatusCode.Unauthorized, rejected.StatusCode);
    }

    [Fact]
    public async Task Signed_collection_request_returns_stable_application_envelope()
    {
        using var server = CreateServer();
        using var client = server.CreateClient();
        using var request = SignedRequest(HttpMethod.Get, "/v1/admin/integration/accounts", "nonce-http-page-001", ReadOnlySpan<byte>.Empty);

        using var response = await client.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var payload = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("chuanhoa.admin.integration.v1", payload.GetProperty("schema").GetString());
        Assert.Equal("chuan-hoa", payload.GetProperty("application").GetString());
        Assert.Equal(0, payload.GetProperty("data").GetProperty("total").GetInt64());
    }

    private static TestServer CreateServer()
    {
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ChuanHoa:AdminIntegration:Enabled"] = "true",
            ["ChuanHoa:AdminIntegration:ClientId"] = ClientId,
            ["ChuanHoa:AdminIntegration:SharedSecret"] = Secret,
            ["ConnectionStrings:ChuanHoa"] = "Host=contract-test"
        }).Build();
        return new TestServer(new WebHostBuilder()
            .ConfigureServices(services =>
            {
                services.AddSingleton<IConfiguration>(configuration);
                services.AddSingleton<TimeProvider>(new FixedTimeProvider(DateTimeOffset.FromUnixTimeSeconds(1_700_000_000)));
                services.AddSingleton<IntegrationRequestAuthenticator>();
                services.AddSingleton<IIntegrationReplayStore, MemoryReplayStore>();
                services.AddSingleton<FakeStore>();
                services.AddSingleton<IIntegrationAdminStore>(provider => provider.GetRequiredService<FakeStore>());
                services.AddSingleton<IActivationKeyAdminStore, FakeActivationKeyStore>();
                services.AddControllers().AddApplicationPart(typeof(IntegrationAdminController).Assembly);
            })
            .Configure(app => app.UseRouting().UseEndpoints(endpoints => endpoints.MapControllers())));
    }

    private static HttpRequestMessage SignedRequest(HttpMethod method, string path, string nonce, ReadOnlySpan<byte> body)
    {
        const string timestamp = "1700000000";
        var bodyHash = Convert.ToHexString(SHA256.HashData(body)).ToLowerInvariant();
        var uri = new Uri(path, UriKind.RelativeOrAbsolute);
        var signedPath = uri.IsAbsoluteUri ? uri.PathAndQuery : path;
        var canonical = string.Join('\n', method.Method.ToUpperInvariant(), signedPath, timestamp, nonce, bodyHash);
        var signature = Convert.ToHexString(HMACSHA256.HashData(Encoding.UTF8.GetBytes(Secret), Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
        var request = new HttpRequestMessage(method, path);
        request.Headers.Add("X-Integration-Client", ClientId);
        request.Headers.Add("X-Integration-Timestamp", timestamp);
        request.Headers.Add("X-Integration-Nonce", nonce);
        request.Headers.Add("X-Integration-Signature", signature);
        return request;
    }

    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => now;
    }

    private sealed class MemoryReplayStore : IIntegrationReplayStore
    {
        private readonly HashSet<string> _seen = new(StringComparer.Ordinal);
        public Task<bool> TryClaimAsync(string clientId, string nonce, DateTimeOffset expiresAtUtc, CancellationToken cancellationToken)
            => Task.FromResult(_seen.Add($"{clientId}:{nonce}"));
    }

    private sealed class FakeStore : IIntegrationAdminStore
    {
        public int MutationCalls { get; private set; }
        public Task<AdminPage<AdminAccount>> AccountsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new AdminPage<AdminAccount>([], page, pageSize, 0));
        public Task<AdminPage<AdminOffer>> OffersAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new AdminPage<AdminOffer>([], page, pageSize, 0));
        public Task<AdminPage<AdminOrder>> OrdersAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new AdminPage<AdminOrder>([], page, pageSize, 0));
        public Task<AdminPage<AdminSubscription>> SubscriptionsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new AdminPage<AdminSubscription>([], page, pageSize, 0));
        public Task<AdminPage<AdminPayment>> PaymentsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new AdminPage<AdminPayment>([], page, pageSize, 0));
        public Task<AdminPage<AdminAudit>> AuditAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new AdminPage<AdminAudit>([], page, pageSize, 0));
        public Task RecordFailedEntitlementAsync(Guid userId, Guid productId, string clientId, string actorId, Guid correlationId, string resultCode, CancellationToken cancellationToken) => Task.CompletedTask;
        public Task<AdminMutation> ExtendEntitlementAsync(Guid userId, Guid productId, IReadOnlyList<string> featureCodes, TimeSpan duration, string reason, string clientId, string idempotencyKey, string actorId, Guid correlationId, CancellationToken cancellationToken)
        {
            MutationCalls++;
            var now = DateTimeOffset.UtcNow;
            return Task.FromResult(new AdminMutation(Guid.NewGuid(), userId, productId, "ACTIVE", now, now.Add(duration), featureCodes));
        }
    }

    private sealed class FakeActivationKeyStore : IActivationKeyAdminStore
    {
        public Task<CreatedVipKeyResponse> CreateVipKeyAsync(CreateVipKeyRequest request, CancellationToken cancellationToken)
            => Task.FromResult(new CreatedVipKeyResponse(Guid.NewGuid(), "CHV-TEST", "CHV-TEST", request.MaxDevices, request.ExpiresAtUtc));
        public Task ReleaseKeyDeviceAsync(AdminDeviceMutationRequest request, CancellationToken cancellationToken) => Task.CompletedTask;
        public Task RevokeKeyAsync(AdminKeyMutationRequest request, CancellationToken cancellationToken) => Task.CompletedTask;
        public Task ResetAccountDeviceAsync(AdminDeviceMutationRequest request, CancellationToken cancellationToken) => Task.CompletedTask;
        public Task<ActivationKeyAdminPage> ListAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) => Task.FromResult(new ActivationKeyAdminPage([], page, pageSize, 0));
    }
}
#pragma warning restore ASPDEPR004, ASPDEPR008
