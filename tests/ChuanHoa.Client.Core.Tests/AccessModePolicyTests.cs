using ChuanHoa.Client.Core.Licensing;

namespace ChuanHoa.Client.Core.Tests;

public sealed class AccessModePolicyTests
{
    private static readonly DateTimeOffset Login = DateTimeOffset.Parse("2026-09-13T10:00:00Z");

    [Fact]
    public void Account_mode_never_outlives_absolute_72_hour_session_even_with_seven_day_lease()
    {
        var state = new AccessModeState(AccessMode.Account, Login, Login.AddHours(72), "device-a");
        Assert.Equal(Login.AddHours(72), AccessModePolicy.EffectiveExpiry(state, Login.AddDays(7)));
        var error = Assert.Throws<InvalidOperationException>(() => AccessModePolicy.Validate(state, "device-a", Login.AddDays(7), Login.AddHours(72)));
        Assert.Equal("ACCOUNT_SESSION_EXPIRED", error.Message);
    }

    [Fact]
    public void Activation_key_mode_does_not_require_an_account_session()
    {
        var state = new AccessModeState(AccessMode.ActivationKey, Login, Login.AddDays(7), "device-a");
        AccessModePolicy.Validate(state, "device-a", Login.AddDays(2), Login.AddDays(1));
    }

    [Fact]
    public void Cached_access_cannot_be_copied_to_another_device()
    {
        var state = new AccessModeState(AccessMode.Account, Login, Login.AddHours(72), "device-a");
        var error = Assert.Throws<InvalidOperationException>(() => AccessModePolicy.Validate(state, "device-b", Login.AddHours(72), Login));
        Assert.Equal("ACCESS_DEVICE_MISMATCH", error.Message);
    }
}
