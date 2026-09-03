using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

[JsonConverter(typeof(JsonStringEnumConverter<EntitlementSource>))]
public enum EntitlementSource
{
    Purchase,
    OrganizationContract,
    ManualGrant,
    LaunchTrial,
    PersonalTrial,
    AdminExtension
}

public sealed record EffectiveEntitlement(
    Guid EntitlementId,
    Guid SubjectId,
    Guid? OrganizationId,
    string ProductCode,
    IReadOnlySet<string> Features,
    EntitlementSource Source,
    DateTimeOffset EffectiveFromUtc,
    DateTimeOffset EffectiveUntilUtc,
    bool Revoked);

public sealed record SignedLeasePayload(
    Guid LeaseId,
    Guid SubjectId,
    Guid? OrganizationId,
    string DeviceKeyThumbprint,
    string ClientReleaseId,
    string ProtocolVersion,
    IReadOnlySet<string> Features,
    EntitlementSource Source,
    DateTimeOffset IssuedAtUtc,
    DateTimeOffset NotBeforeUtc,
    DateTimeOffset ExpiresAtUtc,
    string Jti);

public sealed record ExecutionGrantPayload(
    Guid GrantId,
    string CommandId,
    Guid SubjectId,
    Guid? OrganizationId,
    string DeviceKeyThumbprint,
    string ClientReleaseId,
    string DocumentFingerprint,
    string Scope,
    DateTimeOffset IssuedAtUtc,
    DateTimeOffset NotBeforeUtc,
    DateTimeOffset ExpiresAtUtc,
    string Nonce,
    string Jti);

public sealed record SignedEnvelope<T>(
    string Schema,
    string KeyId,
    string Algorithm,
    string Audience,
    T Payload,
    string Signature);
