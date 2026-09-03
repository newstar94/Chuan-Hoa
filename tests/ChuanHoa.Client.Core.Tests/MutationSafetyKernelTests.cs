using ChuanHoa.Client.Core.Safety;

namespace ChuanHoa.Client.Core.Tests;

public sealed class MutationSafetyKernelTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-01T12:00:00Z");

    [Fact]
    public void Executes_allowlisted_local_operation_and_restores_application_state()
    {
        var document = new RecordingDocument();
        var replay = new RecordingReplayStore();
        var kernel = CreateKernel(replayStore: replay);

        var result = kernel.Execute(Request(), document);

        Assert.Equal(1, result.AppliedOperationCount);
        Assert.Equal(1, document.ApplyCount);
        Assert.Equal(1, document.VerifyCount);
        Assert.True(document.UndoStarted);
        Assert.True(document.UndoEnded);
        Assert.True(document.ApplicationStateRestored);
        Assert.False(document.RollbackCalled);
        Assert.Equal("grant-1", replay.ConsumedJti);
    }

    [Theory]
    [InlineData(true, false, "DOCUMENT_READ_ONLY")]
    [InlineData(false, true, "DOCUMENT_PROTECTED")]
    public void Rejects_unsafe_preflight_without_consuming_grant_or_mutating(
        bool readOnly,
        bool protectedDocument,
        string expectedCode)
    {
        var document = new RecordingDocument
        {
            Preflight = new DocumentMutationPreflight(true, true, readOnly, protectedDocument, false, true)
        };
        var replay = new RecordingReplayStore();
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel(replayStore: replay).Execute(Request(), document));

        Assert.Equal(expectedCode, error.Code);
        Assert.Null(replay.ConsumedJti);
        Assert.Equal(0, document.ApplyCount);
    }

    [Fact]
    public void Rejects_unsupported_document_format_without_consuming_grant_or_mutating()
    {
        var document = new RecordingDocument
        {
            Preflight = new DocumentMutationPreflight(true, true, false, false, false, true, false)
        };
        var replay = new RecordingReplayStore();
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel(replayStore: replay).Execute(Request(), document));

        Assert.Equal("DOCUMENT_FORMAT_UNSUPPORTED", error.Code);
        Assert.Null(replay.ConsumedJti);
        Assert.Equal(0, document.ApplyCount);
    }

    [Fact]
    public void Rejects_fix_plan_for_local_execution_grant_lane()
    {
        var authorization = Authorization(MutationAuthorizationKind.FixPlan, "fix-plan.v1");
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel().Execute(Request(authorization: authorization), new RecordingDocument()));

        Assert.Equal("AUTHORIZATION_LANE_MISMATCH", error.Code);
    }

    [Fact]
    public void Rejects_untrusted_signature_before_consuming_or_mutating()
    {
        var verifier = new StaticVerifier(false);
        var replay = new RecordingReplayStore();
        var document = new RecordingDocument();
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel(verifier, replay).Execute(Request(), document));

        Assert.Equal("AUTHORIZATION_SIGNATURE_INVALID", error.Code);
        Assert.Null(replay.ConsumedJti);
        Assert.Equal(0, document.ApplyCount);
    }

    [Fact]
    public void Rejects_expired_authorization_at_exact_expiry_boundary()
    {
        var expired = Authorization(expiresAtUtc: Now);
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel().Execute(Request(authorization: expired), new RecordingDocument()));

        Assert.Equal("AUTHORIZATION_EXPIRED", error.Code);
    }

    [Fact]
    public void Rejects_unknown_operation_and_report_only_operation()
    {
        var unknown = Request(operation: new MutationOperation("DeleteArbitraryRange", "selected-text", MutationRiskTier.Confirm));
        var unknownError = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel().Execute(unknown, new RecordingDocument()));
        Assert.Equal("OPERATION_NOT_ALLOWLISTED", unknownError.Code);

        var reportOnly = Request(operation: new MutationOperation("SetCharacterSpacing", "selected-text", MutationRiskTier.ReportOnly));
        var riskError = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel().Execute(reportOnly, new RecordingDocument()));
        Assert.Equal("OPERATION_RISK_NOT_ALLOWED", riskError.Code);
    }

    [Fact]
    public void Rejects_fingerprint_change_immediately_before_apply()
    {
        var document = new RecordingDocument
        {
            Fingerprints = new Queue<string>(new[] { "sha256:document", "sha256:document", "sha256:changed" })
        };
        var replay = new RecordingReplayStore();
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel(replayStore: replay).Execute(Request(), document));

        Assert.Equal("DOCUMENT_FINGERPRINT_MISMATCH", error.Code);
        Assert.Null(replay.ConsumedJti);
        Assert.Equal(0, document.ApplyCount);
        Assert.True(document.ApplicationStateRestored);
    }

    [Fact]
    public void Rejects_replayed_authorization_before_undo_or_apply()
    {
        var replay = new RecordingReplayStore { AllowConsume = false };
        var document = new RecordingDocument();
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel(replayStore: replay).Execute(Request(), document));

        Assert.Equal("AUTHORIZATION_REPLAYED", error.Code);
        Assert.False(document.UndoStarted);
        Assert.Equal(0, document.ApplyCount);
        Assert.True(document.ApplicationStateRestored);
    }

    [Fact]
    public void Rolls_back_and_restores_state_when_postcondition_fails()
    {
        var document = new RecordingDocument { VerificationResult = false };
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel().Execute(Request(), document));

        Assert.Equal("POSTCONDITION_FAILED", error.Code);
        Assert.True(document.RollbackCalled);
        Assert.True(document.UndoEnded);
        Assert.True(document.ApplicationStateRestored);
        Assert.True(document.Events.IndexOf("end-undo") < document.Events.IndexOf("rollback"));
    }

    [Fact]
    public void Rejects_concurrent_mutation_lock_without_consuming_authorization()
    {
        var replay = new RecordingReplayStore();
        var document = new RecordingDocument();
        var error = Assert.Throws<MutationSafetyException>(() =>
            CreateKernel(replayStore: replay, mutationLock: new DenyingMutationLock())
                .Execute(Request(), document));

        Assert.Equal("DOCUMENT_MUTATION_IN_PROGRESS", error.Code);
        Assert.Null(replay.ConsumedJti);
        Assert.Equal(0, document.ApplyCount);
    }

    private static MutationSafetyKernel CreateKernel(
        IMutationAuthorizationVerifier? verifier = null,
        RecordingReplayStore? replayStore = null,
        IDocumentMutationLock? mutationLock = null)
    {
        var policy = new MutationCommandPolicy(
            "btnGianChuNormal",
            MutationAuthorizationKind.ExecutionGrant,
            MutationRiskTier.Confirm,
            true,
            true,
            new[] { "SetCharacterSpacing" },
            new[] { "selected-text" });

        return new MutationSafetyKernel(
            new[] { policy },
            verifier ?? new StaticVerifier(true),
            replayStore ?? new RecordingReplayStore(),
            mutationLock ?? new AllowingMutationLock(),
            new StaticClock(Now));
    }

    private static MutationRequest Request(
        MutationAuthorization? authorization = null,
        MutationOperation? operation = null)
    {
        return new MutationRequest(
            "btnGianChuNormal",
            authorization ?? Authorization(),
            new[] { operation ?? new MutationOperation("SetCharacterSpacing", "selected-text", MutationRiskTier.Confirm) },
            true);
    }

    private static MutationAuthorization Authorization(
        MutationAuthorizationKind kind = MutationAuthorizationKind.ExecutionGrant,
        string schema = "execution-grant.v1",
        DateTimeOffset? expiresAtUtc = null)
    {
        return new MutationAuthorization(
            kind,
            schema,
            "btnGianChuNormal",
            "sha256:document",
            "selected-text",
            Now.AddMinutes(-1),
            expiresAtUtc ?? Now.AddMinutes(4),
            "grant-1");
    }

    private sealed class RecordingDocument : IMutationDocumentAdapter
    {
        public string DocumentIdentity { get; } = "document-1";

        public DocumentMutationPreflight Preflight { get; set; } =
            new(true, true, false, false, false, true);

        public Queue<string> Fingerprints { get; set; } =
            new(new[] { "sha256:document", "sha256:document", "sha256:document", "sha256:final" });

        public int ApplyCount { get; private set; }

        public int VerifyCount { get; private set; }

        public bool VerificationResult { get; set; } = true;

        public bool UndoStarted { get; private set; }

        public bool UndoEnded { get; private set; }

        public bool RollbackCalled { get; private set; }

        public bool ApplicationStateRestored { get; private set; }

        public List<string> Events { get; } = new();

        public DocumentMutationPreflight ReadPreflight() => Preflight;

        public string CaptureFingerprint() => Fingerprints.Count == 0 ? "sha256:final" : Fingerprints.Dequeue();

        public object CaptureApplicationState() => new object();

        public object CreateBackup() => new object();

        public void BeginUndoRecord(string commandId)
        {
            UndoStarted = true;
            Events.Add("begin-undo");
        }

        public void Apply(MutationOperation operation) => ApplyCount++;

        public bool Verify(MutationOperation operation)
        {
            VerifyCount++;
            return VerificationResult;
        }

        public void EndUndoRecord()
        {
            UndoEnded = true;
            Events.Add("end-undo");
        }

        public void Rollback(object? backup)
        {
            RollbackCalled = true;
            Events.Add("rollback");
        }

        public void RestoreApplicationState(object applicationState) => ApplicationStateRestored = true;
    }

    private sealed class StaticVerifier : IMutationAuthorizationVerifier
    {
        private readonly bool _result;

        public StaticVerifier(bool result) => _result = result;

        public bool IsAuthentic(MutationAuthorization authorization) => _result;
    }

    private sealed class RecordingReplayStore : IAuthorizationReplayStore
    {
        public bool AllowConsume { get; set; } = true;

        public string? ConsumedJti { get; private set; }

        public bool TryConsume(string jti, DateTimeOffset expiresAtUtc)
        {
            ConsumedJti = jti;
            return AllowConsume;
        }
    }

    private sealed class AllowingMutationLock : IDocumentMutationLock
    {
        public IDisposable TryAcquire(string documentIdentity) => new Lease();
    }

    private sealed class DenyingMutationLock : IDocumentMutationLock
    {
        public IDisposable? TryAcquire(string documentIdentity) => null;
    }

    private sealed class Lease : IDisposable
    {
        public void Dispose()
        {
        }
    }

    private sealed class StaticClock : IUtcClock
    {
        public StaticClock(DateTimeOffset utcNow) => UtcNow = utcNow;

        public DateTimeOffset UtcNow { get; }
    }
}
