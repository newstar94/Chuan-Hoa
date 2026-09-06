using ChuanHoa.Api.Billing;
using Microsoft.Extensions.Configuration;

namespace ChuanHoa.Api.Tests;

public sealed class PurchaseWorkflowTests
{
    [Fact]
    public void Empty_configuration_disables_real_payments() =>
        Assert.False(new PayOsGateway(new ConfigurationBuilder().Build()).IsConfigured);

    [Theory]
    [InlineData(50000, "VND", "link", "00", true)]
    [InlineData(1, "VND", "link", "00", false)]
    [InlineData(50000, "USD", "link", "00", false)]
    [InlineData(50000, "VND", "other", "00", false)]
    [InlineData(50000, "VND", "link", "01", false)]
    public void Signed_webhook_must_match_server_order(long amount, string currency, string link, string code, bool expected)
    {
        var order = new PurchaseReservation(Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(),
            123, 50000, "VND", DateTimeOffset.UtcNow.AddMinutes(15), "AWAITING_PAYMENT", "link");
        Assert.Equal(expected, PurchaseWorkflow.Matches(order,
            new VerifiedPayment(123, amount, currency, link, "reference", code)));
    }
}
