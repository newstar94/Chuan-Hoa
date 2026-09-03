using System.Security.Cryptography;
using ChuanHoa.Client.Core.Safety;

namespace ChuanHoa.Client.Core.Tests;

public sealed class MutationAuthorizationSecurityTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-01T12:00:00Z");

    [Fact]
    public void Rsa_verifier_accepts_bound_payload_and_rejects_tampering()
    {
        using var rsa = RSA.Create(2048);
        var unsigned = Authorization(signature: string.Empty);
        var signature = rsa.SignData(
            MutationAuthorizationCanonicalizer.GetCanonicalBytes(unsigned),
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        var signed = Authorization(signature: Base64Url(signature));
        var verifier = new RsaSha256MutationAuthorizationVerifier(
            new StaticKeyProvider("key-2026-01", rsa.ExportParameters(false)));

        Assert.True(verifier.IsAuthentic(signed));

        var tampered = new MutationAuthorization(
            signed.Kind,
            signed.Schema,
            "btnCoChu",
            signed.DocumentFingerprint,
            signed.Scope,
            signed.NotBeforeUtc,
            signed.ExpiresAtUtc,
            signed.Jti,
            signed.KeyId,
            signed.SignatureAlgorithm,
            signed.Signature,
            signed.SubjectUserId,
            signed.DeviceId,
            signed.DocumentRevision,
            signed.Nonce);
        Assert.False(verifier.IsAuthentic(tampered));
    }

    [Fact]
    public void Rsa_verifier_fails_closed_for_unknown_key_and_missing_device_binding()
    {
        using var rsa = RSA.Create(2048);
        var verifier = new RsaSha256MutationAuthorizationVerifier(
            new StaticKeyProvider("another-key", rsa.ExportParameters(false)));

        Assert.False(verifier.IsAuthentic(Authorization(signature: "AA")));
        Assert.False(verifier.IsAuthentic(Authorization(signature: "AA", deviceId: string.Empty)));
    }

    [Fact]
    public void Persistent_replay_store_survives_new_instance_and_expires_records()
    {
        var path = Path.Combine(Path.GetTempPath(), "chuanhoa-replay-" + Guid.NewGuid().ToString("N") + ".v1");
        try
        {
            var first = new PersistentAuthorizationReplayStore(path, new StaticClock(Now));
            Assert.True(first.TryConsume("grant-persistent", Now.AddMinutes(2)));

            var second = new PersistentAuthorizationReplayStore(path, new StaticClock(Now.AddMinutes(1)));
            Assert.False(second.TryConsume("grant-persistent", Now.AddMinutes(3)));

            var afterExpiry = new PersistentAuthorizationReplayStore(path, new StaticClock(Now.AddMinutes(2)));
            Assert.True(afterExpiry.TryConsume("grant-persistent", Now.AddMinutes(4)));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Persistent_replay_store_rejects_malformed_state()
    {
        var path = Path.Combine(Path.GetTempPath(), "chuanhoa-replay-" + Guid.NewGuid().ToString("N") + ".v1");
        try
        {
            File.WriteAllText(path, "untrusted malformed content");
            var store = new PersistentAuthorizationReplayStore(path, new StaticClock(Now));
            Assert.False(store.TryConsume("grant-1", Now.AddMinutes(1)));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    private static MutationAuthorization Authorization(string signature, string deviceId = "device-1") =>
        new(
            MutationAuthorizationKind.ExecutionGrant,
            "execution-grant.v1",
            "btnGianChuNormal",
            "sha256:document",
            "selected-text",
            Now.AddMinutes(-1),
            Now.AddMinutes(4),
            "grant-1",
            "key-2026-01",
            "RS256",
            signature,
            "user-1",
            deviceId,
            "revision-1",
            "nonce-1");

    private static string Base64Url(byte[] value) =>
        Convert.ToBase64String(value).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    private sealed class StaticKeyProvider : IMutationAuthorizationPublicKeyProvider
    {
        private readonly string _keyId;
        private readonly RSAParameters _publicKey;

        public StaticKeyProvider(string keyId, RSAParameters publicKey)
        {
            _keyId = keyId;
            _publicKey = publicKey;
        }

        public bool TryGetPublicKey(string keyId, out RSAParameters publicKey)
        {
            publicKey = _publicKey;
            return string.Equals(keyId, _keyId, StringComparison.Ordinal);
        }
    }

    private sealed class StaticClock : IUtcClock
    {
        public StaticClock(DateTimeOffset utcNow) => UtcNow = utcNow;

        public DateTimeOffset UtcNow { get; }
    }
}
