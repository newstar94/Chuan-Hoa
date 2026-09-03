using ChuanHoa.Contracts;
using ChuanHoa.Domain.Common;
using ChuanHoa.Domain.Security;

namespace ChuanHoa.Domain.Tests;

public sealed class ExecutionGrantValidatorTests
{
    private static readonly Guid SubjectId = Guid.Parse("11111111-1111-1111-1111-111111111111");
    private static readonly Guid OrganizationId = Guid.Parse("22222222-2222-2222-2222-222222222222");
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-10-01T00:00:00Z");
    private readonly ExecutionGrantValidator _validator = new();

    [Fact]
    public void Accepts_matching_unexpired_metadata()
    {
        _validator.ValidateEnvelopeMetadata(Envelope(), Context());
    }

    [Theory]
    [InlineData("different-command", "GRANT_COMMAND_MISMATCH")]
    [InlineData("different-document", "GRANT_DOCUMENT_MISMATCH")]
    public void Rejects_rebinding(string change, string expectedCode)
    {
        var envelope = Envelope();
        var context = change == "different-command"
            ? Context(expectedCommand: "btnChenTrangDoc")
            : Context(expectedDocument: "sha256:other");
        var error = Assert.Throws<DomainException>(() => _validator.ValidateEnvelopeMetadata(envelope, context));
        Assert.Equal(expectedCode, error.Code);
    }

    [Fact]
    public void Rejects_expired_grant()
    {
        var error = Assert.Throws<DomainException>(() =>
            _validator.ValidateEnvelopeMetadata(Envelope(), Context(now: Now.AddMinutes(6))));
        Assert.Equal("GRANT_EXPIRED", error.Code);
    }

    [Fact]
    public void Rejects_disallowed_algorithm_before_apply()
    {
        var envelope = Envelope() with { Algorithm = "HS256" };
        var error = Assert.Throws<DomainException>(() => _validator.ValidateEnvelopeMetadata(envelope, Context()));
        Assert.Equal("GRANT_ALGORITHM_NOT_ALLOWED", error.Code);
    }

    private static SignedEnvelope<ExecutionGrantPayload> Envelope()
    {
        return new SignedEnvelope<ExecutionGrantPayload>(
            "execution-grant.v1",
            "grant-key-2026-01",
            "ES256",
            "chuanhoa-vsto",
            new ExecutionGrantPayload(
                Guid.NewGuid(),
                "btnChenTrangNgang",
                SubjectId,
                OrganizationId,
                "sha256:device",
                "vsto-1.0.0",
                "sha256:document",
                "selection-and-new-sections",
                Now,
                Now,
                Now.AddMinutes(5),
                "nonce",
                "jti"),
            "base64url-signature");
    }

    private static ExecutionGrantValidationContext Context(
        string expectedCommand = "btnChenTrangNgang",
        string expectedDocument = "sha256:document",
        DateTimeOffset? now = null)
    {
        return new ExecutionGrantValidationContext(
            "chuanhoa-vsto",
            expectedCommand,
            SubjectId,
            OrganizationId,
            "sha256:device",
            "vsto-1.0.0",
            expectedDocument,
            now ?? Now);
    }
}
