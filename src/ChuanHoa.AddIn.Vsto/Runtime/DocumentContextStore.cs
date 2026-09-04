using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class DocumentContextStore : IDisposable
    {
        private readonly Dictionary<long, DocumentContext> _contexts =
            new Dictionary<long, DocumentContext>();
        private bool _disposed;

        public DocumentContext GetOrCreate(Word.Document document)
        {
            ThrowIfDisposed();
            var identity = GetComIdentity(document);
            DocumentContext context;
            if (!_contexts.TryGetValue(identity, out context))
            {
                context = new DocumentContext(identity);
                _contexts.Add(identity, context);
            }

            return context;
        }

        public void Remove(Word.Document document)
        {
            ThrowIfDisposed();
            _contexts.Remove(GetComIdentity(document));
        }

        public void ClearAllReadAnalysis()
        {
            ThrowIfDisposed();
            foreach (var context in _contexts.Values) context.ClearReadAnalysis();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _contexts.Clear();
            _disposed = true;
        }

        private static long GetComIdentity(Word.Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            var unknown = Marshal.GetIUnknownForObject(document);
            try
            {
                return unknown.ToInt64();
            }
            finally
            {
                Marshal.Release(unknown);
            }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }
    }
}
