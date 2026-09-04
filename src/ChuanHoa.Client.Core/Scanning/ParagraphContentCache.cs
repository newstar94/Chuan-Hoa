using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Content-addressable LRU cache for paragraph scan findings.
    /// Prevents repeated analysis of unmodified paragraphs during incremental editing.
    /// </summary>
    public sealed class ParagraphContentCache
    {
        private readonly int _capacity;
        private readonly object _syncLock = new object();
        private readonly Dictionary<string, LinkedListNode<CacheEntry>> _map;
        private readonly LinkedList<CacheEntry> _lruList;

        public ParagraphContentCache(int capacity = 500)
        {
            if (capacity <= 0) throw new ArgumentOutOfRangeException(nameof(capacity));
            _capacity = capacity;
            _map = new Dictionary<string, LinkedListNode<CacheEntry>>(capacity, StringComparer.Ordinal);
            _lruList = new LinkedList<CacheEntry>();
        }

        public int Count
        {
            get
            {
                lock (_syncLock)
                {
                    return _map.Count;
                }
            }
        }

        public static string ComputeParagraphHash(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            using (var sha = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(text);
                var hash = sha.ComputeHash(bytes);
                var builder = new StringBuilder(hash.Length * 2);
                for (var i = 0; i < hash.Length; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }

        public bool TryGet(string hash, out IReadOnlyList<AnnotationFinding> findings)
        {
            if (string.IsNullOrEmpty(hash))
            {
                findings = Array.Empty<AnnotationFinding>();
                return false;
            }

            lock (_syncLock)
            {
                if (_map.TryGetValue(hash, out var node))
                {
                    // Move to head of LRU list (most recently used)
                    _lruList.Remove(node);
                    _lruList.AddFirst(node);
                    findings = node.Value.Findings;
                    return true;
                }
            }

            findings = Array.Empty<AnnotationFinding>();
            return false;
        }

        public void Set(string hash, IReadOnlyList<AnnotationFinding> findings)
        {
            if (string.IsNullOrEmpty(hash)) return;

            lock (_syncLock)
            {
                if (_map.TryGetValue(hash, out var existingNode))
                {
                    _lruList.Remove(existingNode);
                    _lruList.AddFirst(existingNode);
                    existingNode.Value = new CacheEntry(hash, findings);
                    return;
                }

                if (_map.Count >= _capacity && _lruList.Last != null)
                {
                    var oldest = _lruList.Last;
                    _lruList.RemoveLast();
                    _map.Remove(oldest.Value.Hash);
                }

                var entry = new CacheEntry(hash, findings);
                var newNode = _lruList.AddFirst(entry);
                _map[hash] = newNode;
            }
        }

        public void Clear()
        {
            lock (_syncLock)
            {
                _map.Clear();
                _lruList.Clear();
            }
        }

        private sealed class CacheEntry
        {
            public CacheEntry(string hash, IReadOnlyList<AnnotationFinding> findings)
            {
                Hash = hash;
                Findings = findings ?? Array.Empty<AnnotationFinding>();
            }

            public string Hash { get; }
            public IReadOnlyList<AnnotationFinding> Findings { get; }
        }
    }
}
