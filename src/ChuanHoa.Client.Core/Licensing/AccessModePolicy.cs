using System;

namespace ChuanHoa.Client.Core.Licensing
{
    public enum AccessMode
    {
        None,
        Account,
        ActivationKey
    }

    public sealed class AccessModeState
    {
        public AccessModeState(AccessMode mode, DateTimeOffset issuedAtUtc, DateTimeOffset expiresAtUtc, string deviceThumbprint)
        {
            if (mode == AccessMode.None) throw new ArgumentException("ACCESS_MODE_REQUIRED", nameof(mode));
            if (expiresAtUtc <= issuedAtUtc) throw new ArgumentException("ACCESS_EXPIRY_INVALID", nameof(expiresAtUtc));
            if (string.IsNullOrWhiteSpace(deviceThumbprint)) throw new ArgumentException("DEVICE_REQUIRED", nameof(deviceThumbprint));
            Mode = mode; IssuedAtUtc = issuedAtUtc; ExpiresAtUtc = expiresAtUtc; DeviceThumbprint = deviceThumbprint;
        }

        public AccessMode Mode { get; }
        public DateTimeOffset IssuedAtUtc { get; }
        public DateTimeOffset ExpiresAtUtc { get; }
        public string DeviceThumbprint { get; }
    }

    public static class AccessModePolicy
    {
        public static DateTimeOffset EffectiveExpiry(AccessModeState state, DateTimeOffset signedLeaseExpiryUtc)
            => state.Mode == AccessMode.Account && state.ExpiresAtUtc < signedLeaseExpiryUtc
                ? state.ExpiresAtUtc
                : signedLeaseExpiryUtc;

        public static void Validate(AccessModeState? state, string expectedDeviceThumbprint, DateTimeOffset signedLeaseExpiryUtc, DateTimeOffset nowUtc)
        {
            if (state == null) throw new InvalidOperationException("ACCESS_REQUIRED");
            if (!string.Equals(state.DeviceThumbprint, expectedDeviceThumbprint, StringComparison.Ordinal))
                throw new InvalidOperationException("ACCESS_DEVICE_MISMATCH");
            if (nowUtc >= EffectiveExpiry(state, signedLeaseExpiryUtc))
                throw new InvalidOperationException(state.Mode == AccessMode.Account ? "ACCOUNT_SESSION_EXPIRED" : "ACTIVATION_EXPIRED");
        }
    }
}
