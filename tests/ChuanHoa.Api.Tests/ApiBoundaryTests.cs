using System.Net;
using System.Net.Http.Json;
using ChuanHoa.Api.Errors;
using ChuanHoa.Api.Middleware;
using ChuanHoa.Api.Security;
using ChuanHoa.Api.Controllers;
using ChuanHoa.Application;
using ChuanHoa.Application.Scanning;
using ChuanHoa.Contracts;
using ChuanHoa.Domain.Common;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;

namespace ChuanHoa.Api.Tests;

public sealed class ApiBoundaryTests
{
    [Fact]
    public async Task Valid_correlation_id_is_preserved_in_header_and_context()
    {
        await using var fixture = await ApiFixture.CreateAsync();
        var correlationId = Guid.NewGuid().ToString("D");
        using var request = new HttpRequestMessage(HttpMethod.Get, "/echo-correlation");
        request.Headers.Add(ApiHeaders.CorrelationId, correlationId);

        using var response = await fixture.Client.SendAsync(request);
        var echoedCorrelationId = await response.Content.ReadAsStringAsync();

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal(correlationId, response.Headers.GetValues(ApiHeaders.CorrelationId).Single());
        Assert.Equal(correlationId, echoedCorrelationId);
    }

    [Fact]
    public async Task Invalid_correlation_id_is_replaced_with_server_guid()
    {
        await using var fixture = await ApiFixture.CreateAsync();
        using var request = new HttpRequestMessage(HttpMethod.Get, "/echo-correlation");
        request.Headers.Add(ApiHeaders.CorrelationId, "not-a-guid");

        using var response = await fixture.Client.SendAsync(request);
        var returned = response.Headers.GetValues(ApiHeaders.CorrelationId).Single();

        Assert.True(Guid.TryParseExact(returned, "D", out var parsed));
        Assert.NotEqual(Guid.Empty, parsed);
        Assert.NotEqual("not-a-guid", returned);
    }

    [Fact]
    public async Task Mutation_without_idempotency_key_returns_versioned_error()
    {
        await using var fixture = await ApiFixture.CreateAsync();

        using var response = await fixture.Client.PostAsync("/v1/mutate", null);
        var error = await response.Content.ReadFromJsonAsync<ApiErrorResponse>();

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.NotNull(error);
        Assert.Equal(ApiContractVersions.ErrorV1, error.Schema);
        Assert.Equal("IDEMPOTENCY_KEY_REQUIRED", error.Code);
        Assert.Equal(response.Headers.GetValues(ApiHeaders.CorrelationId).Single(), error.CorrelationId);
        Assert.False(string.IsNullOrWhiteSpace(error.Recovery));
    }

    [Fact]
    public async Task Invalid_idempotency_key_returns_stable_error_code()
    {
        await using var fixture = await ApiFixture.CreateAsync();
        using var request = new HttpRequestMessage(HttpMethod.Post, "/v1/mutate");
        request.Headers.Add(ApiHeaders.IdempotencyKey, "short");

        using var response = await fixture.Client.SendAsync(request);
        var error = await response.Content.ReadFromJsonAsync<ApiErrorResponse>();

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.NotNull(error);
        Assert.Equal("IDEMPOTENCY_KEY_INVALID", error.Code);
    }

    [Fact]
    public async Task Valid_idempotency_key_reaches_mutation_handler()
    {
        await using var fixture = await ApiFixture.CreateAsync();
        using var request = new HttpRequestMessage(HttpMethod.Post, "/v1/mutate");
        request.Headers.Add(ApiHeaders.IdempotencyKey, "mutation-key-00000001");

        using var response = await fixture.Client.SendAsync(request);

        Assert.Equal(HttpStatusCode.NoContent, response.StatusCode);
    }

    [Fact]
    public async Task Unhandled_exception_does_not_leak_internal_detail()
    {
        await using var fixture = await ApiFixture.CreateAsync();

        using var response = await fixture.Client.GetAsync("/throw-unhandled");
        var error = await response.Content.ReadFromJsonAsync<ApiErrorResponse>();

        Assert.Equal(HttpStatusCode.InternalServerError, response.StatusCode);
        Assert.NotNull(error);
        Assert.Equal("INTERNAL_ERROR", error.Code);
        Assert.DoesNotContain("sensitive-diagnostic", error.Detail, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(error.CorrelationId));
    }

    [Fact]
    public async Task Grant_domain_failure_is_mapped_to_forbidden_with_recovery()
    {
        await using var fixture = await ApiFixture.CreateAsync();

        using var response = await fixture.Client.GetAsync("/throw-grant");
        var error = await response.Content.ReadFromJsonAsync<ApiErrorResponse>();

        Assert.Equal(HttpStatusCode.Forbidden, response.StatusCode);
        Assert.NotNull(error);
        Assert.Equal("GRANT_DEVICE_MISMATCH", error.Code);
        Assert.Contains("new execution grant", error.Recovery, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Document_scan_is_fail_closed_without_configured_identity_provider()
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            EnvironmentName = "Testing"
        });
        builder.WebHost.UseTestServer();
        builder.Services.AddProblemDetails();
        builder.Services.AddExceptionHandler<ApiExceptionHandler>();
        builder.Services.AddSingleton<IClock, SystemClock>();
        builder.Services.AddSingleton<TechnicalDocumentScanner>();
        builder.Services.AddControllers().AddApplicationPart(typeof(DocumentScanController).Assembly);
        builder.Services
            .AddAuthentication(FailClosedAuthenticationDefaults.Scheme)
            .AddScheme<Microsoft.AspNetCore.Authentication.AuthenticationSchemeOptions, FailClosedAuthenticationHandler>(
                FailClosedAuthenticationDefaults.Scheme,
                _ => { });
        builder.Services.AddAuthorization(options =>
        {
            options.AddPolicy("DocumentScan", policy =>
            {
                policy.RequireAuthenticatedUser();
                policy.RequireClaim("scope", "document.scan");
            });
        });

        await using var app = builder.Build();
        app.UseMiddleware<CorrelationIdMiddleware>();
        app.UseExceptionHandler();
        app.UseMiddleware<IdempotencyKeyValidationMiddleware>();
        app.UseAuthentication();
        app.UseAuthorization();
        app.MapControllers();
        await app.StartAsync();
        using var client = app.GetTestClient();
        using var request = new HttpRequestMessage(HttpMethod.Post, "/v1/document-jobs/scan")
        {
            Content = JsonContent.Create(new { })
        };
        request.Headers.Add(ApiHeaders.IdempotencyKey, "scan-key-000000000001");

        using var response = await client.SendAsync(request);

        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
    }

    private sealed class ApiFixture : IAsyncDisposable
    {
        private ApiFixture(WebApplication application)
        {
            Application = application;
            Client = application.GetTestClient();
        }

        public WebApplication Application { get; }

        public HttpClient Client { get; }

        public static async Task<ApiFixture> CreateAsync()
        {
            var builder = WebApplication.CreateBuilder(new WebApplicationOptions
            {
                EnvironmentName = "Testing"
            });
            builder.WebHost.UseTestServer();
            builder.Services.AddProblemDetails();
            builder.Services.AddExceptionHandler<ApiExceptionHandler>();

            var app = builder.Build();
            app.UseMiddleware<CorrelationIdMiddleware>();
            app.UseExceptionHandler();
            app.UseMiddleware<IdempotencyKeyValidationMiddleware>();
            app.MapGet(
                "/echo-correlation",
                (HttpContext context) => Results.Text(CorrelationIdMiddleware.GetCorrelationId(context)));
            app.MapPost("/v1/mutate", () => Results.NoContent());
            app.MapGet("/throw-unhandled", () => ThrowUnhandled());
            app.MapGet("/throw-grant", () => ThrowGrantFailure());
            await app.StartAsync();
            return new ApiFixture(app);
        }

        public async ValueTask DisposeAsync()
        {
            Client.Dispose();
            await Application.DisposeAsync();
        }

        private static IResult ThrowUnhandled()
        {
            throw new InvalidOperationException("sensitive-diagnostic");
        }

        private static IResult ThrowGrantFailure()
        {
            throw new DomainException("GRANT_DEVICE_MISMATCH", "The device binding does not match.");
        }
    }
}
