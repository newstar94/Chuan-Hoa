using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace ChuanHoa.Client.Core.Licensing
{
    public sealed class OfflineLease
    {
        public OfflineLease(string leaseId, string subjectId, string deviceThumbprint, string clientReleaseId,
            DateTimeOffset issuedAtUtc, DateTimeOffset notBeforeUtc, DateTimeOffset expiresAtUtc,
            IReadOnlyCollection<string> features)
        {
            LeaseId = leaseId;
            SubjectId = subjectId;
            DeviceThumbprint = deviceThumbprint;
            ClientReleaseId = clientReleaseId;
            IssuedAtUtc = issuedAtUtc;
            NotBeforeUtc = notBeforeUtc;
            ExpiresAtUtc = expiresAtUtc;
            Features = features;
        }

        public string LeaseId { get; }
        public string SubjectId { get; }
        public string DeviceThumbprint { get; }
        public string ClientReleaseId { get; }
        public DateTimeOffset IssuedAtUtc { get; }
        public DateTimeOffset NotBeforeUtc { get; }
        public DateTimeOffset ExpiresAtUtc { get; }
        public IReadOnlyCollection<string> Features { get; }
    }

    public static class OfflineLeaseParser
    {
        public const string Schema = "chuanhoa.offline-lease.v1";
        public static OfflineLease Parse(byte[] payload)
        {
            if (payload == null || payload.Length == 0) throw new FormatException("Lease payload is empty.");
            var root = XElement.Parse(System.Text.Encoding.UTF8.GetString(payload), LoadOptions.None);
            if (root.Name.LocalName != "offlineLease" || (string)root.Attribute("schema") != Schema)
                throw new FormatException("Lease schema is unsupported.");
            var features = root.Element("features")?.Elements("feature")
                .Select(item => Required(item, "code")).Distinct(StringComparer.Ordinal).ToArray()
                ?? throw new FormatException("Lease features are missing.");
            return new OfflineLease(
                Required(root, "leaseId"), Required(root, "subjectId"), Required(root, "deviceThumbprint"),
                Required(root, "clientReleaseId"), Time(root, "issuedAtUtc"), Time(root, "notBeforeUtc"),
                Time(root, "expiresAtUtc"), features);
        }

        private static string Required(XElement element, string name) =>
            (string)element.Attribute(name) ?? throw new FormatException("Missing lease attribute: " + name + ".");
        private static DateTimeOffset Time(XElement element, string name) =>
            DateTimeOffset.ParseExact(Required(element, name), "O", CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal);
    }

    public sealed class OfflineLeaseValidator
    {
        public static readonly TimeSpan MaximumOfflinePeriod = TimeSpan.FromDays(7);

        public void Validate(OfflineLease lease, string expectedDeviceThumbprint, string expectedClientReleaseId,
            string requiredFeature, DateTimeOffset nowUtc, DateTimeOffset? lastTrustedServerTimeUtc = null)
        {
            if (lease == null) throw new ArgumentNullException(nameof(lease));
            Ensure(lease.ExpiresAtUtc > lease.IssuedAtUtc, "LEASE_TIME_RANGE_INVALID");
            Ensure(lease.ExpiresAtUtc - lease.IssuedAtUtc <= MaximumOfflinePeriod, "LEASE_OFFLINE_PERIOD_EXCEEDED");
            Ensure(nowUtc >= lease.NotBeforeUtc, "LEASE_NOT_ACTIVE");
            Ensure(nowUtc < lease.ExpiresAtUtc, "LEASE_EXPIRED");
            Ensure(string.Equals(lease.DeviceThumbprint, expectedDeviceThumbprint, StringComparison.Ordinal), "LEASE_DEVICE_MISMATCH");
            Ensure(string.Equals(lease.ClientReleaseId, expectedClientReleaseId, StringComparison.Ordinal), "LEASE_RELEASE_MISMATCH");
            Ensure(lease.Features.Contains(requiredFeature, StringComparer.Ordinal), "LEASE_FEATURE_MISSING");
            if (lastTrustedServerTimeUtc.HasValue)
                Ensure(nowUtc >= lastTrustedServerTimeUtc.Value.Subtract(TimeSpan.FromMinutes(5)), "LEASE_CLOCK_ROLLBACK");
        }

        private static void Ensure(bool condition, string code)
        {
            if (!condition) throw new InvalidOperationException(code);
        }
    }
}
