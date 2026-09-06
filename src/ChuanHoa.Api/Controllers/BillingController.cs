using System.Security.Claims;
using ChuanHoa.Api.Billing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using PayOS.Models.Webhooks;

namespace ChuanHoa.Api.Controllers;

[ApiController]
[Route("v1/billing")]
public sealed class BillingController(PayOsGateway gateway, IServiceProvider services) : ControllerBase
{
    public sealed record CreatePurchase(Guid OfferId);

    [Authorize]
    [HttpPost("checkout")]
    public async Task<IActionResult> Checkout(CreatePurchase request, CancellationToken cancellationToken)
    {
        if (!Guid.TryParse(User.FindFirstValue(ClaimTypes.NameIdentifier), out var userId)) return Forbid();
        var store = services.GetService<IPurchaseStore>();
        if (!gateway.IsConfigured || store == null)
            return StatusCode(503, new { code = "BILLING_NOT_CONFIGURED", message = "Thanh toán chưa được cấu hình." });
        var flow = new PurchaseWorkflow(store, gateway, TimeProvider.System);
        return Ok(await flow.CreateAsync(userId, request.OfferId,
            Request.Headers["Idempotency-Key"].ToString(), cancellationToken));
    }

    [AllowAnonymous]
    [HttpPost("payos/webhook")]
    public async Task<IActionResult> Webhook(Webhook webhook, CancellationToken cancellationToken)
    {
        var store = services.GetService<IPurchaseStore>();
        if (!gateway.IsConfigured || store == null) return StatusCode(503);
        VerifiedPayment verified;
        try { verified = await gateway.VerifyAsync(webhook); }
        catch (PayOS.Exceptions.WebhookException) { return BadRequest(new { code = "PAYMENT_SIGNATURE_INVALID" }); }
        await new PurchaseWorkflow(store, gateway, TimeProvider.System).ConfirmAsync(verified, cancellationToken);
        return Ok(new { received = true });
    }
}
