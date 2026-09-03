using System;
using System.Collections.Generic;

namespace ChuanHoa.Client.Core.Safety
{
    public sealed class MutationSafetyKernel
    {
        private const string ExecutionGrantSchema = "execution-grant.v1";
        private const string FixPlanSchema = "fix-plan.v1";

        private readonly Dictionary<string, MutationCommandPolicy> _policies;
        private readonly IMutationAuthorizationVerifier _authorizationVerifier;
        private readonly IAuthorizationReplayStore _replayStore;
        private readonly IDocumentMutationLock _mutationLock;
        private readonly IUtcClock _clock;

        public MutationSafetyKernel(
            IEnumerable<MutationCommandPolicy> policies,
            IMutationAuthorizationVerifier authorizationVerifier,
            IAuthorizationReplayStore replayStore,
            IDocumentMutationLock mutationLock,
            IUtcClock clock)
        {
            if (policies == null)
            {
                throw new ArgumentNullException(nameof(policies));
            }

            _authorizationVerifier = authorizationVerifier ?? throw new ArgumentNullException(nameof(authorizationVerifier));
            _replayStore = replayStore ?? throw new ArgumentNullException(nameof(replayStore));
            _mutationLock = mutationLock ?? throw new ArgumentNullException(nameof(mutationLock));
            _clock = clock ?? throw new ArgumentNullException(nameof(clock));
            _policies = new Dictionary<string, MutationCommandPolicy>(StringComparer.Ordinal);

            foreach (var policy in policies)
            {
                if (policy == null || string.IsNullOrWhiteSpace(policy.CommandId))
                {
                    throw new ArgumentException("Every mutation policy must have a command ID.", nameof(policies));
                }

                if (_policies.ContainsKey(policy.CommandId))
                {
                    throw new ArgumentException("Duplicate mutation command policy: " + policy.CommandId, nameof(policies));
                }

                _policies.Add(policy.CommandId, policy);
            }
        }

        public MutationExecutionResult Execute(MutationRequest request, IMutationDocumentAdapter document)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            var policy = ValidateRequest(request);
            ValidatePreflight(document.ReadPreflight(), policy);
            EnsureFingerprint(document.CaptureFingerprint(), request.Authorization.DocumentFingerprint);

            using (var lease = _mutationLock.TryAcquire(document.DocumentIdentity))
            {
                if (lease == null)
                {
                    throw Reject("DOCUMENT_MUTATION_IN_PROGRESS", "Another mutation is already running for this document.");
                }

                ValidatePreflight(document.ReadPreflight(), policy);
                EnsureFingerprint(document.CaptureFingerprint(), request.Authorization.DocumentFingerprint);

                object? applicationState = null;
                object? backup = null;
                var undoStarted = false;
                var mutationStarted = false;
                var appliedCount = 0;

                try
                {
                    applicationState = document.CaptureApplicationState();
                    if (policy.RequiresBackup)
                    {
                        backup = document.CreateBackup();
                        if (backup == null)
                        {
                            throw Reject("BACKUP_CREATION_FAILED", "The required document backup was not created.");
                        }
                    }

                    // This is intentionally the last document read before consuming the one-time
                    // authorization and entering the apply loop.
                    EnsureFingerprint(document.CaptureFingerprint(), request.Authorization.DocumentFingerprint);
                    if (!_replayStore.TryConsume(request.Authorization.Jti, request.Authorization.ExpiresAtUtc))
                    {
                        throw Reject("AUTHORIZATION_REPLAYED", "The mutation authorization was already consumed.");
                    }

                    document.BeginUndoRecord(request.CommandId);
                    undoStarted = true;

                    foreach (var operation in request.Operations)
                    {
                        mutationStarted = true;
                        document.Apply(operation);
                        if (!document.Verify(operation))
                        {
                            throw Reject("POSTCONDITION_FAILED", "A mutation operation did not satisfy its postcondition.");
                        }

                        appliedCount++;
                    }

                    return new MutationExecutionResult(
                        request.CommandId,
                        appliedCount,
                        document.CaptureFingerprint());
                }
                catch (MutationSafetyException exception)
                {
                    try
                    {
                        EndUndoBeforeRollback(document, ref undoStarted);
                        RollbackIfNeeded(document, backup, mutationStarted);
                    }
                    catch (Exception recoveryException)
                    {
                        throw new MutationSafetyException(
                            "MUTATION_RECOVERY_FAILED",
                            "The mutation was rejected and its recovery sequence failed.",
                            new AggregateException(exception, recoveryException));
                    }

                    throw;
                }
                catch (Exception exception)
                {
                    try
                    {
                        EndUndoBeforeRollback(document, ref undoStarted);
                        RollbackIfNeeded(document, backup, mutationStarted);
                    }
                    catch (Exception rollbackException)
                    {
                        throw new MutationSafetyException(
                            "APPLY_AND_ROLLBACK_FAILED",
                            "The mutation failed and rollback also failed.",
                            new AggregateException(exception, rollbackException));
                    }

                    throw new MutationSafetyException("MUTATION_APPLY_FAILED", "The mutation failed and was rolled back.", exception);
                }
                finally
                {
                    try
                    {
                        if (undoStarted)
                        {
                            document.EndUndoRecord();
                        }
                    }
                    finally
                    {
                        if (applicationState != null)
                        {
                            document.RestoreApplicationState(applicationState);
                        }
                    }
                }
            }
        }

        private MutationCommandPolicy ValidateRequest(MutationRequest request)
        {
            MutationCommandPolicy policy;
            if (string.IsNullOrWhiteSpace(request.CommandId) || !_policies.TryGetValue(request.CommandId, out policy))
            {
                throw Reject("COMMAND_NOT_ALLOWLISTED", "The command is not in the client mutation allowlist.");
            }

            var authorization = request.Authorization;
            if (authorization == null)
            {
                throw Reject("AUTHORIZATION_REQUIRED", "A signed mutation authorization is required.");
            }

            if (authorization.Kind != policy.AuthorizationKind)
            {
                throw Reject("AUTHORIZATION_LANE_MISMATCH", "The command authorization lane does not match its policy.");
            }

            var expectedSchema = authorization.Kind == MutationAuthorizationKind.ExecutionGrant
                ? ExecutionGrantSchema
                : FixPlanSchema;
            if (!string.Equals(authorization.Schema, expectedSchema, StringComparison.Ordinal))
            {
                throw Reject("AUTHORIZATION_SCHEMA_UNKNOWN", "The mutation authorization schema is not supported.");
            }

            if (!_authorizationVerifier.IsAuthentic(authorization))
            {
                throw Reject("AUTHORIZATION_SIGNATURE_INVALID", "The mutation authorization signature is invalid.");
            }

            if (!string.Equals(request.CommandId, authorization.CommandId, StringComparison.Ordinal))
            {
                throw Reject("AUTHORIZATION_COMMAND_MISMATCH", "The mutation authorization is bound to another command.");
            }

            var now = _clock.UtcNow;
            if (now < authorization.NotBeforeUtc)
            {
                throw Reject("AUTHORIZATION_NOT_ACTIVE", "The mutation authorization is not active yet.");
            }

            if (now >= authorization.ExpiresAtUtc)
            {
                throw Reject("AUTHORIZATION_EXPIRED", "The mutation authorization has expired.");
            }

            if (string.IsNullOrWhiteSpace(authorization.Jti))
            {
                throw Reject("AUTHORIZATION_JTI_REQUIRED", "The mutation authorization has no one-time identity.");
            }

            if (request.Operations == null || request.Operations.Count == 0)
            {
                throw Reject("OPERATIONS_REQUIRED", "At least one authorized operation is required.");
            }

            if (policy.MaximumRiskTier == MutationRiskTier.ReportOnly || policy.MaximumRiskTier == MutationRiskTier.Blocked)
            {
                throw Reject("COMMAND_NOT_MUTABLE", "Report-only and blocked commands cannot mutate a document.");
            }

            if (policy.RequiresPreviewConfirmation && !request.PreviewConfirmed)
            {
                throw Reject("PREVIEW_CONFIRMATION_REQUIRED", "The user must confirm the operation preview before apply.");
            }

            foreach (var operation in request.Operations)
            {
                if (operation == null || !Contains(policy.AllowedOperationTypes, operation.OperationType))
                {
                    throw Reject("OPERATION_NOT_ALLOWLISTED", "The authorization contains an unknown or disallowed operation.");
                }

                if (!Contains(policy.AllowedScopes, operation.Scope) ||
                    !string.Equals(operation.Scope, authorization.Scope, StringComparison.Ordinal))
                {
                    throw Reject("OPERATION_SCOPE_MISMATCH", "The operation scope is outside the signed authorization.");
                }

                if (operation.RiskTier == MutationRiskTier.ReportOnly ||
                    operation.RiskTier == MutationRiskTier.Blocked ||
                    operation.RiskTier > policy.MaximumRiskTier)
                {
                    throw Reject("OPERATION_RISK_NOT_ALLOWED", "The operation risk tier is not mutable under this command policy.");
                }
            }

            return policy;
        }

        private static void ValidatePreflight(DocumentMutationPreflight preflight, MutationCommandPolicy policy)
        {
            if (preflight == null || !preflight.HasActiveDocument || !preflight.HasActiveWindow)
            {
                throw Reject("ACTIVE_DOCUMENT_REQUIRED", "An active document and window are required.");
            }

            if (!preflight.IsSupportedDocumentFormat)
            {
                throw Reject("DOCUMENT_FORMAT_UNSUPPORTED", "Only saved .doc and .docx documents are supported.");
            }

            if (preflight.IsReadOnly)
            {
                throw Reject("DOCUMENT_READ_ONLY", "The active document is read-only.");
            }

            if (preflight.IsProtected)
            {
                throw Reject("DOCUMENT_PROTECTED", "The active document is protected.");
            }

            if (preflight.HasConcurrentMutation)
            {
                throw Reject("DOCUMENT_MUTATION_IN_PROGRESS", "Another mutation is already running for this document.");
            }

            if (policy.RequiresBackup && !preflight.CanCreateBackup)
            {
                throw Reject("BACKUP_UNAVAILABLE", "This command requires a backup before apply.");
            }
        }

        private static void EnsureFingerprint(string actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(actual) ||
                string.IsNullOrWhiteSpace(expected) ||
                !string.Equals(actual, expected, StringComparison.Ordinal))
            {
                throw Reject("DOCUMENT_FINGERPRINT_MISMATCH", "The document changed after authorization or preview.");
            }
        }

        private static bool Contains(IReadOnlyCollection<string> values, string candidate)
        {
            if (values == null || string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            foreach (var value in values)
            {
                if (string.Equals(value, candidate, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static void RollbackIfNeeded(IMutationDocumentAdapter document, object? backup, bool mutationStarted)
        {
            if (mutationStarted)
            {
                document.Rollback(backup);
            }
        }

        private static void EndUndoBeforeRollback(IMutationDocumentAdapter document, ref bool undoStarted)
        {
            if (!undoStarted)
            {
                return;
            }

            document.EndUndoRecord();
            undoStarted = false;
        }

        private static MutationSafetyException Reject(string code, string message)
        {
            return new MutationSafetyException(code, message);
        }
    }
}
