using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ChuanHoa.Client.Core.Lexicon
{
    /// <summary>
    /// Manages local user dictionary and document-specific ignored word lists.
    /// Operates completely offline, persisting to %LocalAppData%\ChuanHoa\Dictionaries.
    /// </summary>
    public sealed class PersonalDictionaryManager
    {
        private static readonly object SyncLock = new object();
        private static PersonalDictionaryManager? _instance;

        public static PersonalDictionaryManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (SyncLock)
                    {
                        if (_instance == null)
                        {
                            var localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                            var dir = Path.Combine(localApp, "ChuanHoa", "Dictionaries");
                            _instance = new PersonalDictionaryManager(dir);
                        }
                    }
                }
                return _instance;
            }
        }

        private readonly string _storageDirectory;
        private readonly string _userCustomFilePath;
        private readonly HashSet<string> _userWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HashSet<string>> _documentIgnoredWords =
            new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        public PersonalDictionaryManager(string storageDirectory)
        {
            _storageDirectory = storageDirectory ?? throw new ArgumentNullException(nameof(storageDirectory));
            _userCustomFilePath = Path.Combine(_storageDirectory, "user_custom_dictionary.txt");
            LoadUserWords();
        }

        public bool IsKnownOrIgnored(string? word, string? documentId = null)
        {
            if (string.IsNullOrWhiteSpace(word)) return true;
            var clean = word!.Trim();

            lock (SyncLock)
            {
                if (_userWords.Contains(clean)) return true;

                if (!string.IsNullOrEmpty(documentId) &&
                    _documentIgnoredWords.TryGetValue(documentId!, out var ignored) &&
                    ignored.Contains(clean))
                {
                    return true;
                }
            }
            return false;
        }

        public void AddUserWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return;
            var clean = word.Trim();

            lock (SyncLock)
            {
                if (_userWords.Add(clean))
                {
                    SaveUserWords();
                }
            }
        }

        public void RemoveUserWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return;
            var clean = word.Trim();

            lock (SyncLock)
            {
                if (_userWords.Remove(clean))
                {
                    SaveUserWords();
                }
            }
        }

        public void IgnoreWordForDocument(string documentId, string word)
        {
            if (string.IsNullOrWhiteSpace(documentId) || string.IsNullOrWhiteSpace(word)) return;
            var clean = word.Trim();

            lock (SyncLock)
            {
                if (!_documentIgnoredWords.TryGetValue(documentId, out var set))
                {
                    set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    _documentIgnoredWords[documentId] = set;
                }
                set.Add(clean);
            }
        }

        public void ClearDocumentIgnores(string documentId)
        {
            if (string.IsNullOrWhiteSpace(documentId)) return;
            lock (SyncLock)
            {
                _documentIgnoredWords.Remove(documentId);
            }
        }

        public IReadOnlyCollection<string> GetUserWords()
        {
            lock (SyncLock)
            {
                return _userWords.ToList();
            }
        }

        private void LoadUserWords()
        {
            try
            {
                if (File.Exists(_userCustomFilePath))
                {
                    var lines = File.ReadAllLines(_userCustomFilePath, Encoding.UTF8);
                    lock (SyncLock)
                    {
                        _userWords.Clear();
                        foreach (var line in lines)
                        {
                            var trimmed = line.Trim();
                            if (!string.IsNullOrEmpty(trimmed) && !trimmed.StartsWith("#", StringComparison.Ordinal))
                            {
                                _userWords.Add(trimmed);
                            }
                        }
                    }
                }
            }
            catch
            {
                // Fallback gracefully without throwing
            }
        }

        private void SaveUserWords()
        {
            try
            {
                if (!Directory.Exists(_storageDirectory))
                {
                    Directory.CreateDirectory(_storageDirectory);
                }

                List<string> snapshot;
                lock (SyncLock)
                {
                    snapshot = _userWords.OrderBy(w => w).ToList();
                }
                File.WriteAllLines(_userCustomFilePath, snapshot, Encoding.UTF8);
            }
            catch
            {
                // Fallback gracefully
            }
        }
    }
}
