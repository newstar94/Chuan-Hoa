using ChuanHoa.Api.Development;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;

namespace ChuanHoa.Api.Tests;

public sealed class DevelopmentAdminStoreTests
{
    [Fact]
    public async Task Development_admin_persists_users_trial_and_versioned_prices()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ChuanHoaAdminTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var statePath = Path.Combine(directory, "admin-state.json");
            var builder = WebApplication.CreateBuilder(new WebApplicationOptions
            {
                EnvironmentName = "Development",
                ContentRootPath = directory
            });
            builder.Configuration.AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["ChuanHoa:DevelopmentAdminStatePath"] = statePath
            });
            await using var app = builder.Build();
            var now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");
            var store = new DevelopmentAdminStore(app.Environment, app.Configuration, TimeProvider.System);

            var withUser = store.CreateUser(new CreateDevelopmentUserRequest(
                "user@example.test", "Người dùng thử", now.AddDays(7)), now);
            var withOffer = store.CreateOffer(new CreateDevelopmentOfferRequest(
                "Gói cá nhân", 490_000, now.AddMonths(1), null, "SCHEDULED"), now);
            var withTrial = store.SetTrial(new DevelopmentTrialSettings(
                now, now.AddMonths(1), true, false, 30), now);

            Assert.Single(withUser.Users);
            Assert.Single(withOffer.Offers);
            Assert.False(withTrial.Trial.PersonalTrialEnabled);
            var reloaded = new DevelopmentAdminStore(app.Environment, app.Configuration, TimeProvider.System).Read();
            Assert.Single(reloaded.Users);
            Assert.Single(reloaded.Offers);
            Assert.Equal(490_000, reloaded.Offers[0].AmountVnd);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }
}
