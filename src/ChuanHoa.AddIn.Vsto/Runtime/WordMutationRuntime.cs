using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using ChuanHoa.Client.Core.Safety;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class WordMutationRuntime
    {
        private readonly Word.Application _application;
        private readonly MutationSafetyKernel _kernel;

        public WordMutationRuntime(Word.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            var replayPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ChuanHoa",
                "Security",
                "authorization-replay.v1");
            var clock = new SystemUtcClock();
            _kernel = new MutationSafetyKernel(
                CreatePolicies(),
                new RsaSha256MutationAuthorizationVerifier(new EmptyPublicKeyProvider()),
                new PersistentAuthorizationReplayStore(replayPath, clock),
                new InProcessDocumentMutationLock(),
                clock);
        }

        public MutationExecutionResult Execute(MutationRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var document = _application.ActiveDocument;
            using (var adapter = new WordComMutationDocumentAdapter(_application, document))
            {
                return _kernel.Execute(request, adapter);
            }
        }

        public bool HasAuthorizationKeys => false;

        private static IEnumerable<MutationCommandPolicy> CreatePolicies()
        {
            yield return new MutationCommandPolicy(
                "btnGianChuNormal",
                MutationAuthorizationKind.ExecutionGrant,
                MutationRiskTier.Confirm,
                true,
                true,
                new[] { "SetCharacterSpacing" },
                new[] { "selected-text" });
        }

        private sealed class EmptyPublicKeyProvider : IMutationAuthorizationPublicKeyProvider
        {
            public bool TryGetPublicKey(string keyId, out RSAParameters publicKey)
            {
                publicKey = default(RSAParameters);
                return false;
            }
        }
    }
}
