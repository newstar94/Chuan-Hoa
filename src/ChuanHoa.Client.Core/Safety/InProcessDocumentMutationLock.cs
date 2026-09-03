using System;
using System.Collections.Generic;

namespace ChuanHoa.Client.Core.Safety
{
    public sealed class InProcessDocumentMutationLock : IDocumentMutationLock
    {
        private readonly object _sync = new object();
        private readonly HashSet<string> _activeDocuments = new HashSet<string>(StringComparer.Ordinal);

        public IDisposable? TryAcquire(string documentIdentity)
        {
            if (string.IsNullOrWhiteSpace(documentIdentity))
            {
                return null;
            }

            lock (_sync)
            {
                if (!_activeDocuments.Add(documentIdentity))
                {
                    return null;
                }
            }

            return new Lease(this, documentIdentity);
        }

        private void Release(string documentIdentity)
        {
            lock (_sync)
            {
                _activeDocuments.Remove(documentIdentity);
            }
        }

        private sealed class Lease : IDisposable
        {
            private InProcessDocumentMutationLock? _owner;
            private readonly string _documentIdentity;

            public Lease(InProcessDocumentMutationLock owner, string documentIdentity)
            {
                _owner = owner;
                _documentIdentity = documentIdentity;
            }

            public void Dispose()
            {
                var owner = _owner;
                if (owner != null)
                {
                    _owner = null;
                    owner.Release(_documentIdentity);
                }
            }
        }
    }
}
