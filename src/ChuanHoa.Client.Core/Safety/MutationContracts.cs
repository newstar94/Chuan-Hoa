using System;
using System.Collections.Generic;

namespace ChuanHoa.Client.Core.Safety
{
    public enum MutationAuthorizationKind
    {
        ExecutionGrant,
        FixPlan
    }

    public enum MutationRiskTier
    {
        Safe,
        Confirm,
        ReportOnly,
        Blocked
    }

    public sealed class DocumentMutationPreflight
    {
        public DocumentMutationPreflight(
            bool hasActiveDocument,
            bool hasActiveWindow,
            bool isReadOnly,
            bool isProtected,
            bool hasConcurrentMutation,
            bool canCreateBackup,
            bool isSupportedDocumentFormat = true)
        {
            HasActiveDocument = hasActiveDocument;
            HasActiveWindow = hasActiveWindow;
            IsReadOnly = isReadOnly;
            IsProtected = isProtected;
            HasConcurrentMutation = hasConcurrentMutation;
            CanCreateBackup = canCreateBackup;
            IsSupportedDocumentFormat = isSupportedDocumentFormat;
        }

        public bool HasActiveDocument { get; }

        public bool HasActiveWindow { get; }

        public bool IsReadOnly { get; }

        public bool IsProtected { get; }

        public bool HasConcurrentMutation { get; }

        public bool CanCreateBackup { get; }

        public bool IsSupportedDocumentFormat { get; }
    }

    public sealed class MutationAuthorization
    {
        public MutationAuthorization(
            MutationAuthorizationKind kind,
            string schema,
            string commandId,
            string documentFingerprint,
            string scope,
            DateTimeOffset notBeforeUtc,
            DateTimeOffset expiresAtUtc,
            string jti,
            string keyId = "",
            string signatureAlgorithm = "",
            string signature = "",
            string subjectUserId = "",
            string deviceId = "",
            string documentRevision = "",
            string nonce = "")
        {
            Kind = kind;
            Schema = schema;
            CommandId = commandId;
            DocumentFingerprint = documentFingerprint;
            Scope = scope;
            NotBeforeUtc = notBeforeUtc;
            ExpiresAtUtc = expiresAtUtc;
            Jti = jti;
            KeyId = keyId;
            SignatureAlgorithm = signatureAlgorithm;
            Signature = signature;
            SubjectUserId = subjectUserId;
            DeviceId = deviceId;
            DocumentRevision = documentRevision;
            Nonce = nonce;
        }

        public MutationAuthorizationKind Kind { get; }

        public string Schema { get; }

        public string CommandId { get; }

        public string DocumentFingerprint { get; }

        public string Scope { get; }

        public DateTimeOffset NotBeforeUtc { get; }

        public DateTimeOffset ExpiresAtUtc { get; }

        public string Jti { get; }

        public string KeyId { get; }

        public string SignatureAlgorithm { get; }

        public string Signature { get; }

        public string SubjectUserId { get; }

        public string DeviceId { get; }

        public string DocumentRevision { get; }

        public string Nonce { get; }
    }

    public sealed class MutationOperation
    {
        public MutationOperation(string operationType, string scope, MutationRiskTier riskTier)
        {
            OperationType = operationType;
            Scope = scope;
            RiskTier = riskTier;
        }

        public string OperationType { get; }

        public string Scope { get; }

        public MutationRiskTier RiskTier { get; }
    }

    public sealed class MutationCommandPolicy
    {
        public MutationCommandPolicy(
            string commandId,
            MutationAuthorizationKind authorizationKind,
            MutationRiskTier maximumRiskTier,
            bool requiresPreviewConfirmation,
            bool requiresBackup,
            IReadOnlyCollection<string> allowedOperationTypes,
            IReadOnlyCollection<string> allowedScopes)
        {
            CommandId = commandId;
            AuthorizationKind = authorizationKind;
            MaximumRiskTier = maximumRiskTier;
            RequiresPreviewConfirmation = requiresPreviewConfirmation;
            RequiresBackup = requiresBackup;
            AllowedOperationTypes = allowedOperationTypes;
            AllowedScopes = allowedScopes;
        }

        public string CommandId { get; }

        public MutationAuthorizationKind AuthorizationKind { get; }

        public MutationRiskTier MaximumRiskTier { get; }

        public bool RequiresPreviewConfirmation { get; }

        public bool RequiresBackup { get; }

        public IReadOnlyCollection<string> AllowedOperationTypes { get; }

        public IReadOnlyCollection<string> AllowedScopes { get; }
    }

    public sealed class MutationRequest
    {
        public MutationRequest(
            string commandId,
            MutationAuthorization authorization,
            IReadOnlyList<MutationOperation> operations,
            bool previewConfirmed)
        {
            CommandId = commandId;
            Authorization = authorization;
            Operations = operations;
            PreviewConfirmed = previewConfirmed;
        }

        public string CommandId { get; }

        public MutationAuthorization Authorization { get; }

        public IReadOnlyList<MutationOperation> Operations { get; }

        public bool PreviewConfirmed { get; }
    }

    public sealed class MutationExecutionResult
    {
        public MutationExecutionResult(string commandId, int appliedOperationCount, string finalFingerprint)
        {
            CommandId = commandId;
            AppliedOperationCount = appliedOperationCount;
            FinalFingerprint = finalFingerprint;
        }

        public string CommandId { get; }

        public int AppliedOperationCount { get; }

        public string FinalFingerprint { get; }
    }
}
