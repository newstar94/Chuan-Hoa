namespace ChuanHoa.Api.Billing;

public sealed record PurchaseReservation(Guid OrderId, Guid UserId, Guid OfferId,
    long OrderCode, long AmountVnd, string Currency, DateTimeOffset ExpiresAtUtc,
    string State, string? PaymentLinkId);

/// <summary>
/// Production adapter must use durable transactions and the existing quotes/orders/
/// entitlement_grants tables. Do not replace this with an in-memory store.
/// </summary>
public interface IPurchaseStore
{
    // Resolve the published offer and freeze its server-side price/term/features.
    // Same user+idempotency key must return the same order, or reject changed offer.
    Task<PurchaseReservation> ReserveAsync(Guid userId, Guid offerId, string idempotencyKey, CancellationToken cancellationToken);
    Task SaveLinkAsync(Guid orderId, CheckoutPayment payment, CancellationToken cancellationToken);
    Task<PurchaseReservation?> FindByCodeAsync(long orderCode, CancellationToken cancellationToken);
    // Atomically verify expected order state, record unique payment reference,
    // grant exactly the purchased term/features, and queue signed lease refresh.
    // Duplicate webhook returns success without extending the term again.
    Task ActivateOnceAsync(PurchaseReservation order, VerifiedPayment payment, CancellationToken cancellationToken);
}

public sealed class PurchaseWorkflow(IPurchaseStore store, PayOsGateway gateway, TimeProvider clock)
{
    public async Task<CheckoutPayment> CreateAsync(Guid userId, Guid offerId, string key, CancellationToken cancellationToken)
    {
        if (userId == Guid.Empty || offerId == Guid.Empty || string.IsNullOrWhiteSpace(key))
            throw new ArgumentException("INVALID_PURCHASE_REQUEST");
        if (!gateway.IsConfigured) throw new InvalidOperationException("PAYOS_NOT_CONFIGURED");
        var order = await store.ReserveAsync(userId, offerId, key, cancellationToken);
        if (order.UserId != userId || order.OfferId != offerId || order.Currency != "VND" ||
            order.ExpiresAtUtc <= clock.GetUtcNow() || order.State != "AWAITING_PAYMENT")
            throw new InvalidOperationException("PURCHASE_NOT_PAYABLE");
        var payment = await gateway.CreateAsync(new CheckoutOrder(order.OrderCode, order.AmountVnd, "CH " + order.OrderCode));
        await store.SaveLinkAsync(order.OrderId, payment, cancellationToken);
        return payment;
    }

    public async Task ConfirmAsync(VerifiedPayment payment, CancellationToken cancellationToken)
    {
        var order = await store.FindByCodeAsync(payment.OrderCode, cancellationToken);
        if (order == null || !Matches(order, payment)) throw new InvalidOperationException("PAYMENT_ORDER_MISMATCH");
        await store.ActivateOnceAsync(order, payment, cancellationToken);
    }

    public static bool Matches(PurchaseReservation order, VerifiedPayment payment) =>
        payment.Code == "00" && payment.Currency == "VND" && order.Currency == "VND" &&
        order.AmountVnd > 0 && order.AmountVnd == payment.AmountVnd &&
        order.OrderCode == payment.OrderCode && !string.IsNullOrWhiteSpace(order.PaymentLinkId) &&
        string.Equals(order.PaymentLinkId, payment.PaymentLinkId, StringComparison.Ordinal) &&
        !string.IsNullOrWhiteSpace(payment.Reference) &&
        (order.State == "AWAITING_PAYMENT" || order.State == "ACTIVATED");
}
