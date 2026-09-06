using PayOS;
using PayOS.Models.V2.PaymentRequests;
using PayOS.Models.Webhooks;

namespace ChuanHoa.Api.Billing;

public sealed record CheckoutOrder(long OrderCode, long AmountVnd, string Description);
public sealed record CheckoutPayment(string PaymentLinkId, string QrCode, string CheckoutUrl);
public sealed record VerifiedPayment(long OrderCode, long AmountVnd, string Currency,
    string PaymentLinkId, string Reference, string Code);

/// <summary>Server-side only. Never accepts price, secrets or entitlement flags from an add-in.</summary>
public sealed class PayOsGateway(IConfiguration configuration)
{
    private string Value(string key) => configuration["PayOS:" + key] ?? string.Empty;

    public bool IsConfigured => new[] { "ClientId", "ApiKey", "ChecksumKey" }.All(k =>
        !string.IsNullOrWhiteSpace(Value(k))) && IsHttps(Value("ReturnUrl")) && IsHttps(Value("CancelUrl"));

    private static bool IsHttps(string value) => Uri.TryCreate(value, UriKind.Absolute, out var uri) &&
        uri.Scheme == Uri.UriSchemeHttps && string.IsNullOrEmpty(uri.UserInfo);

    private PayOSClient Client()
    {
        if (!IsConfigured) throw new InvalidOperationException("PAYOS_NOT_CONFIGURED");
        return new PayOSClient(Value("ClientId"), Value("ApiKey"), Value("ChecksumKey"));
    }

    public async Task<CheckoutPayment> CreateAsync(CheckoutOrder order)
    {
        if (order.OrderCode <= 0 || order.AmountVnd <= 0 || string.IsNullOrWhiteSpace(order.Description))
            throw new ArgumentException("INVALID_CHECKOUT_ORDER");
        var payment = await Client().PaymentRequests.CreateAsync(new CreatePaymentLinkRequest
        {
            OrderCode = order.OrderCode, Amount = order.AmountVnd, Description = order.Description,
            ReturnUrl = Value("ReturnUrl"), CancelUrl = Value("CancelUrl")
        });
        return new CheckoutPayment(payment.PaymentLinkId, payment.QrCode, payment.CheckoutUrl);
    }

    public async Task<VerifiedPayment> VerifyAsync(Webhook webhook)
    {
        var data = await Client().Webhooks.VerifyAsync(webhook);
        // Never treat the unsigned outer success flag or redirect as payment proof.
        return new VerifiedPayment(data.OrderCode, data.Amount, data.Currency,
            data.PaymentLinkId, data.Reference, data.Code);
    }
}
