using System;

namespace ChuanHoa.Client.Core.Safety
{
    public interface IMutationDocumentAdapter
    {
        string DocumentIdentity { get; }

        DocumentMutationPreflight ReadPreflight();

        string CaptureFingerprint();

        object CaptureApplicationState();

        object CreateBackup();

        void BeginUndoRecord(string commandId);

        void Apply(MutationOperation operation);

        bool Verify(MutationOperation operation);

        void EndUndoRecord();

        void Rollback(object? backup);

        void RestoreApplicationState(object applicationState);
    }

    public interface IMutationAuthorizationVerifier
    {
        bool IsAuthentic(MutationAuthorization authorization);
    }

    public interface IAuthorizationReplayStore
    {
        bool TryConsume(string jti, DateTimeOffset expiresAtUtc);
    }

    public interface IDocumentMutationLock
    {
        IDisposable? TryAcquire(string documentIdentity);
    }

    public interface IUtcClock
    {
        DateTimeOffset UtcNow { get; }
    }
}
