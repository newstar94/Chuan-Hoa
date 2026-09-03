using ChuanHoa.Contracts;
using ChuanHoa.Domain.Common;

namespace ChuanHoa.Domain.Security;

public sealed record ExecutionGrantValidationContext(
    string ExpectedAudience,
    string ExpectedCommandId,
    Guid ExpectedSubjectId,
    Guid? ExpectedOrganizationId,
    string ExpectedDeviceKeyThumbprint,
    string ExpectedClientReleaseId,
    string ExpectedDocumentFingerprint,
    DateTimeOffset ServerNowUtc);

public sealed class ExecutionGrantValidator
{
    public void ValidateEnvelopeMetadata(
        SignedEnvelope<ExecutionGrantPayload> envelope,
        ExecutionGrantValidationContext context)
    {
        ArgumentNullException.ThrowIfNull(envelope);
        ArgumentNullException.ThrowIfNull(context);
        if (envelope.Algorithm is not "ES256" and not "EdDSA")
        {
            throw new DomainException("GRANT_ALGORITHM_NOT_ALLOWED", "Execution grant algorithm is not allowed.");
        }

        if (string.IsNullOrWhiteSpace(envelope.KeyId) || string.IsNullOrWhiteSpace(envelope.Signature))
        {
            throw new DomainException("GRANT_SIGNATURE_REQUIRED", "Execution grant signature metadata is required.");
        }

        if (envelope.Audience != context.ExpectedAudience)
        {
            throw new DomainException("GRANT_AUDIENCE_MISMATCH", "Execution grant audience does not match.");
        }

        var payload = envelope.Payload;
        Ensure(payload.CommandId == context.ExpectedCommandId, "GRANT_COMMAND_MISMATCH");
        Ensure(payload.SubjectId == context.ExpectedSubjectId, "GRANT_SUBJECT_MISMATCH");
        Ensure(payload.OrganizationId == context.ExpectedOrganizationId, "GRANT_ORGANIZATION_MISMATCH");
        Ensure(payload.DeviceKeyThumbprint == context.ExpectedDeviceKeyThumbprint, "GRANT_DEVICE_MISMATCH");
        Ensure(payload.ClientReleaseId == context.ExpectedClientReleaseId, "GRANT_RELEASE_MISMATCH");
        Ensure(payload.DocumentFingerprint == context.ExpectedDocumentFingerprint, "GRANT_DOCUMENT_MISMATCH");
        Ensure(context.ServerNowUtc >= payload.NotBeforeUtc, "GRANT_NOT_ACTIVE");
        Ensure(context.ServerNowUtc < payload.ExpiresAtUtc, "GRANT_EXPIRED");
        Ensure(payload.ExpiresAtUtc > payload.IssuedAtUtc, "GRANT_TIME_RANGE_INVALID");
        Ensure(!string.IsNullOrWhiteSpace(payload.Nonce), "GRANT_NONCE_REQUIRED");
        Ensure(!string.IsNullOrWhiteSpace(payload.Jti), "GRANT_JTI_REQUIRED");
    }

    private static void Ensure(bool condition, string code)
    {
        if (!condition)
        {
            throw new DomainException(code, $"Execution grant validation failed: {code}.");
        }
    }
}
