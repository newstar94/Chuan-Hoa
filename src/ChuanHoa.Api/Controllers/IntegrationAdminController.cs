using ChuanHoa.Api.Security;
using ChuanHoa.Infrastructure.Admin;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace ChuanHoa.Api.Controllers;

[ApiController]
[Route("v1/admin/integration")]
public sealed class IntegrationAdminController(
    IntegrationRequestAuthenticator authenticator,
    IConfiguration configuration) : ControllerBase
{
    private const string IntegrationSchema = "chuanhoa.admin.integration.v1";
    public sealed record ExtendEntitlementRequest(Guid UserId, Guid ProductId, string[] FeatureCodes, int DurationDays, string Reason, string ActorId, Guid CorrelationId);
    [HttpGet("capabilities")]
    public IActionResult Capabilities()
    {
        if (!authenticator.TryAuthenticate(Request, ReadOnlySpan<byte>.Empty, configuration))
            return Unauthorized(new { code = "INTEGRATION_AUTH_INVALID", message = "Server-to-server authentication failed." });

        var configured = !string.IsNullOrWhiteSpace(configuration["ConnectionStrings:ChuanHoa"]);
        return Ok(new
        {
            schema = "chuanhoa.admin.integration.v1",
            application = "chuan-hoa",
            status = configured ? "available" : "not_configured",
            capabilities = new { overview = configured, accounts = configured, plans = configured, subscriptions = configured, billing = configured, activation = configured },
            reason = configured ? "Production administration sources and idempotent entitlement extension are available." : "Production administration stores are not configured."
        });
    }

    [HttpGet("accounts")]
    public async Task<IActionResult> Accounts([FromQuery] string? search, [FromQuery] int page = 1, [FromQuery] int pageSize = 25, CancellationToken cancellationToken = default)
        => await Read(async store => await store.AccountsAsync(search, page, pageSize, cancellationToken));

    [HttpGet("offers")]
    public async Task<IActionResult> Offers([FromQuery] string? search, [FromQuery] int page = 1, [FromQuery] int pageSize = 25, CancellationToken cancellationToken = default)
        => await Read(async store => await store.OffersAsync(search, page, pageSize, cancellationToken));

    [HttpGet("orders")]
    public async Task<IActionResult> Orders([FromQuery] string? search, [FromQuery] int page = 1, [FromQuery] int pageSize = 25, CancellationToken cancellationToken = default)
        => await Read(async store => await store.OrdersAsync(search, page, pageSize, cancellationToken));

    [HttpGet("subscriptions")]
    public async Task<IActionResult> Subscriptions([FromQuery] string? search, [FromQuery] int page = 1, [FromQuery] int pageSize = 25, CancellationToken cancellationToken = default)
        => await Read(async store => await store.SubscriptionsAsync(search, page, pageSize, cancellationToken));

    [HttpGet("payments")]
    public async Task<IActionResult> Payments([FromQuery] string? search, [FromQuery] int page = 1, [FromQuery] int pageSize = 25, CancellationToken cancellationToken = default)
        => await Read(async store => await store.PaymentsAsync(search, page, pageSize, cancellationToken));

    [HttpGet("audit")]
    public async Task<IActionResult> Audit([FromQuery] string? search, [FromQuery] int page = 1, [FromQuery] int pageSize = 25, CancellationToken cancellationToken = default)
        => await Read(async store => await store.AuditAsync(search, page, pageSize, cancellationToken));

    [HttpPost("entitlements/extend")]
    public async Task<IActionResult> ExtendEntitlement([FromBody] JsonElement document, CancellationToken cancellationToken)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(document);
        if (!authenticator.TryAuthenticate(Request, bytes, configuration))
            return Unauthorized(new { code = "INTEGRATION_AUTH_INVALID", message = "Server-to-server authentication failed." });
        var body = document.Deserialize<ExtendEntitlementRequest>(new JsonSerializerOptions(JsonSerializerDefaults.Web));
        if (body is null) return BadRequest(new { code = "ADMIN_ENTITLEMENT_REQUEST_INVALID" });
        var key = Request.Headers["Idempotency-Key"].ToString().Trim();
        var client = Request.Headers["X-Integration-Client"].ToString().Trim();
        if (key.Length is < 16 or > 128 || body.ActorId.Length is < 1 or > 128 || body.CorrelationId == Guid.Empty)
            return BadRequest(new { code = "ADMIN_ENTITLEMENT_REQUEST_INVALID" });
        var store = HttpContext.RequestServices.GetService<IIntegrationAdminStore>();
        if (store is null) return StatusCode(StatusCodes.Status503ServiceUnavailable, new { code = "CHUAN_HOA_ADMIN_NOT_CONFIGURED" });
        async Task RecordFailure(string code)
        {
            try
            {
                await store.RecordFailedEntitlementAsync(body.UserId, body.ProductId, client, body.ActorId, body.CorrelationId, code, cancellationToken);
            }
            catch
            {
                // Preserve the original command error; audit availability is reported by health/operations separately.
            }
        }
        try
        {
            var mutation = await store.ExtendEntitlementAsync(body.UserId, body.ProductId, body.FeatureCodes, TimeSpan.FromDays(body.DurationDays), body.Reason, client, key, body.ActorId, body.CorrelationId, cancellationToken);
            return Ok(new { schema = IntegrationSchema, application = "chuan-hoa", status = "available", data = mutation });
        }
        catch (ArgumentException error) { await RecordFailure(error.Message); return BadRequest(new { code = error.Message }); }
        catch (InvalidOperationException error) when (error.Message == "IDEMPOTENCY_KEY_CONFLICT") { await RecordFailure(error.Message); return Conflict(new { code = error.Message }); }
        catch (InvalidOperationException error) when (error.Message == "ADMIN_ENTITLEMENT_TARGET_INVALID") { await RecordFailure(error.Message); return NotFound(new { code = error.Message }); }
    }

    private async Task<IActionResult> Read<T>(Func<IIntegrationAdminStore, Task<AdminPage<T>>> query)
    {
        if (!authenticator.TryAuthenticate(Request, ReadOnlySpan<byte>.Empty, configuration))
            return Unauthorized(new { code = "INTEGRATION_AUTH_INVALID", message = "Server-to-server authentication failed." });
        var store = HttpContext.RequestServices.GetService<IIntegrationAdminStore>();
        if (store is null) return StatusCode(StatusCodes.Status503ServiceUnavailable, new { code = "CHUAN_HOA_ADMIN_NOT_CONFIGURED" });
        var page = await query(store);
        return Ok(new { schema = IntegrationSchema, application = "chuan-hoa", status = "available", data = page });
    }
}
