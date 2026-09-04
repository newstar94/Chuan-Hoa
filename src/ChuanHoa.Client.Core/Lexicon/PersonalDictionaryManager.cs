using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ChuanHoa.Client.Core.Lexicon
{
    public enum PersonalDictionaryStatus
    {
        Success,
        NoChange,
        Duplicate,
        InvalidInput,
        LimitExceeded,
        IoError
    }

    public sealed class PersonalDictionaryResult
    {
        private PersonalDictionaryResult(PersonalDictionaryStatus status, string message)
        {
            Status = status;
            Message = message ?? string.Empty;
        }

        public PersonalDictionaryStatus Status { get; }
        public string Message { get; }
        public bool Succeeded => Status == PersonalDictionaryStatus.Success ||
            Status == PersonalDictionaryStatus.NoChange || Status == PersonalDictionaryStatus.Duplicate;

        public static PersonalDictionaryResult From(PersonalDictionaryStatus status, string message = "") =>
            new PersonalDictionaryResult(status, message);
    }

    public sealed class PersonalDictionarySnapshot
    {
        public PersonalDictionarySnapshot(IReadOnlyList<string> words, PersonalDictionaryResult persistence)
        {
            Words = words ?? throw new ArgumentNullException(nameof(words));
            Persistence = persistence ?? throw new ArgumentNullException(nameof(persistence));
        }

        public IReadOnlyList<string> Words { get; }
        public PersonalDictionaryResult Persistence { get; }
    }

    /// <summary>
    /// Stores the user's rule-based dictionary locally. Document ignore lists are
    /// session-only and are keyed by the document fingerprint supplied to the scanner.
    /// </summary>
    public sealed class PersonalDictionaryManager
    {
        public const int MaximumEntryLength = 120;
        public const int MaximumEntryCount = 10000;

        private static readonly Lazy<PersonalDictionaryManager> Shared =
            new Lazy<PersonalDictionaryManager>(CreateShared, true);

        private readonly object _sync = new object();
        private readonly string _storageDirectory;
        private readonly string _userCustomFilePath;
        private readonly string _lastGoodFilePath;
        private readonly HashSet<string> _userWords =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HashSet<string>> _documentIgnoredWords =
            new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        private PersonalDictionaryResult _lastPersistenceResult =
            PersonalDictionaryResult.From(PersonalDictionaryStatus.Success);

        public static PersonalDictionaryManager Instance => Shared.Value;

        public event EventHandler? Changed;

        public PersonalDictionaryManager(string storageDirectory)
        {
            if (string.IsNullOrWhiteSpace(storageDirectory))
                throw new ArgumentException("A dictionary storage directory is required.", nameof(storageDirectory));
            _storageDirectory = Path.GetFullPath(storageDirectory);
            _userCustomFilePath = Path.Combine(_storageDirectory, "user_custom_dictionary.txt");
            _lastGoodFilePath = _userCustomFilePath + ".last-good";
            LoadUserWords();
        }

        public bool IsKnownOrIgnored(string? word, string? documentId = null)
        {
            string clean;
            PersonalDictionaryResult ignoredFailure;
            if (!TryNormalizeEntry(word, out clean, out ignoredFailure)) return string.IsNullOrWhiteSpace(word);
            var cleanDocumentId = NormalizeDocumentId(documentId);
            lock (_sync)
            {
                if (_userWords.Contains(clean)) return true;
                HashSet<string> ignored;
                return cleanDocumentId != null &&
                    _documentIgnoredWords.TryGetValue(cleanDocumentId, out ignored) &&
                    ignored.Contains(clean);
            }
        }

        public PersonalDictionaryResult AddUserWord(string? word)
        {
            string clean;
            PersonalDictionaryResult invalid;
            if (!TryNormalizeEntry(word, out clean, out invalid)) return invalid;
            PersonalDictionaryResult result;
            lock (_sync)
            {
                if (_userWords.Contains(clean))
                    return PersonalDictionaryResult.From(PersonalDictionaryStatus.Duplicate,
                        "Từ hoặc cụm từ này đã có trong từ điển cá nhân.");
                if (_userWords.Count >= MaximumEntryCount)
                    return PersonalDictionaryResult.From(PersonalDictionaryStatus.LimitExceeded,
                        "Từ điển cá nhân đã đạt giới hạn số mục.");
                _userWords.Add(clean);
                result = SaveUserWordsLocked();
                if (!result.Succeeded) _userWords.Remove(clean);
            }
            if (result.Succeeded) OnChanged();
            return result;
        }

        public PersonalDictionaryResult RemoveUserWord(string? word)
        {
            string clean;
            PersonalDictionaryResult invalid;
            if (!TryNormalizeEntry(word, out clean, out invalid)) return invalid;
            PersonalDictionaryResult result;
            lock (_sync)
            {
                if (!_userWords.Remove(clean))
                    return PersonalDictionaryResult.From(PersonalDictionaryStatus.NoChange);
                result = SaveUserWordsLocked();
                if (!result.Succeeded) _userWords.Add(clean);
            }
            if (result.Succeeded) OnChanged();
            return result;
        }

        public PersonalDictionaryResult IgnoreWordForDocument(string? documentId, string? word)
        {
            var cleanDocumentId = NormalizeDocumentId(documentId);
            if (cleanDocumentId == null)
                return PersonalDictionaryResult.From(PersonalDictionaryStatus.InvalidInput,
                    "Không xác định được tài liệu hiện tại.");
            string clean;
            PersonalDictionaryResult invalid;
            if (!TryNormalizeEntry(word, out clean, out invalid)) return invalid;
            bool changed;
            lock (_sync)
            {
                HashSet<string> ignored;
                if (!_documentIgnoredWords.TryGetValue(cleanDocumentId, out ignored))
                {
                    ignored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    _documentIgnoredWords.Add(cleanDocumentId, ignored);
                }
                changed = ignored.Add(clean);
            }
            if (changed) OnChanged();
            return PersonalDictionaryResult.From(changed
                ? PersonalDictionaryStatus.Success : PersonalDictionaryStatus.Duplicate);
        }

        public PersonalDictionaryResult ClearDocumentIgnores(string? documentId)
        {
            var cleanDocumentId = NormalizeDocumentId(documentId);
            if (cleanDocumentId == null)
                return PersonalDictionaryResult.From(PersonalDictionaryStatus.InvalidInput,
                    "Không xác định được tài liệu hiện tại.");
            bool changed;
            lock (_sync) changed = _documentIgnoredWords.Remove(cleanDocumentId);
            if (changed) OnChanged();
            return PersonalDictionaryResult.From(changed
                ? PersonalDictionaryStatus.Success : PersonalDictionaryStatus.NoChange);
        }

        public PersonalDictionaryResult ClearAllDocumentIgnores()
        {
            bool changed;
            lock (_sync)
            {
                changed = _documentIgnoredWords.Count > 0;
                _documentIgnoredWords.Clear();
            }
            if (changed) OnChanged();
            return PersonalDictionaryResult.From(changed
                ? PersonalDictionaryStatus.Success : PersonalDictionaryStatus.NoChange);
        }

        public PersonalDictionarySnapshot GetUserWordsResult()
        {
            lock (_sync)
            {
                return new PersonalDictionarySnapshot(
                    _userWords.OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToArray(),
                    _lastPersistenceResult);
            }
        }

        public IReadOnlyCollection<string> GetUserWords() => GetUserWordsResult().Words;

        private void LoadUserWords()
        {
            lock (_sync)
            {
                var sourcePath = File.Exists(_userCustomFilePath)
                    ? _userCustomFilePath
                    : File.Exists(_lastGoodFilePath) ? _lastGoodFilePath : null;
                if (sourcePath == null) return;
                try
                {
                    ReplaceWordsFromFileLocked(sourcePath);
                    _lastPersistenceResult = PersonalDictionaryResult.From(PersonalDictionaryStatus.Success);
                }
                catch (Exception exception) when (IsPersistenceException(exception))
                {
                    if (string.Equals(sourcePath, _userCustomFilePath, StringComparison.OrdinalIgnoreCase) &&
                        File.Exists(_lastGoodFilePath))
                    {
                        try
                        {
                            ReplaceWordsFromFileLocked(_lastGoodFilePath);
                            _lastPersistenceResult = PersonalDictionaryResult.From(PersonalDictionaryStatus.IoError,
                                "File từ điển chính bị lỗi; ứng dụng đã dùng bản sao an toàn gần nhất.");
                            return;
                        }
                        catch (Exception backupException) when (IsPersistenceException(backupException))
                        {
                        }
                    }

                    _lastPersistenceResult = PersonalDictionaryResult.From(PersonalDictionaryStatus.IoError,
                        "Không đọc được từ điển cá nhân. Hãy kiểm tra quyền truy cập thư mục dữ liệu ứng dụng.");
                }
            }
        }

        private void ReplaceWordsFromFileLocked(string path)
        {
            var loaded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var strictUtf8 = new UTF8Encoding(false, true);
            foreach (var line in File.ReadAllLines(path, strictUtf8))
            {
                string clean;
                PersonalDictionaryResult ignored;
                if (TryNormalizeEntry(line, out clean, out ignored) &&
                    !clean.StartsWith("#", StringComparison.Ordinal))
                    loaded.Add(clean);
                if (loaded.Count >= MaximumEntryCount) break;
            }
            _userWords.Clear();
            foreach (var word in loaded) _userWords.Add(word);
        }

        private PersonalDictionaryResult SaveUserWordsLocked()
        {
            string? temporaryPath = null;
            try
            {
                Directory.CreateDirectory(_storageDirectory);
                temporaryPath = Path.Combine(_storageDirectory,
                    ".user_custom_dictionary." + Guid.NewGuid().ToString("N") + ".tmp");
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write,
                    FileShare.None, 4096, FileOptions.WriteThrough))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    foreach (var word in _userWords.OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                        writer.WriteLine(word);
                    writer.Flush();
                    stream.Flush(true);
                }

                if (File.Exists(_userCustomFilePath))
                    File.Replace(temporaryPath, _userCustomFilePath, _lastGoodFilePath, true);
                else
                    File.Move(temporaryPath, _userCustomFilePath);
                temporaryPath = null;
                _lastPersistenceResult = PersonalDictionaryResult.From(PersonalDictionaryStatus.Success);
            }
            catch (Exception exception) when (IsPersistenceException(exception))
            {
                _lastPersistenceResult = PersonalDictionaryResult.From(PersonalDictionaryStatus.IoError,
                    "Không lưu được từ điển cá nhân. Hãy kiểm tra quyền truy cập hoặc dung lượng ổ đĩa.");
            }
            finally
            {
                if (temporaryPath != null)
                {
                    try { File.Delete(temporaryPath); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
            return _lastPersistenceResult;
        }

        private static bool TryNormalizeEntry(string? value, out string clean,
            out PersonalDictionaryResult failure)
        {
            clean = (value ?? string.Empty).Trim().Normalize(NormalizationForm.FormC);
            if (clean.Length == 0)
            {
                failure = PersonalDictionaryResult.From(PersonalDictionaryStatus.InvalidInput,
                    "Từ hoặc cụm từ không được để trống.");
                return false;
            }
            if (clean.Length > MaximumEntryLength)
            {
                failure = PersonalDictionaryResult.From(PersonalDictionaryStatus.LimitExceeded,
                    "Từ hoặc cụm từ không được dài quá " + MaximumEntryLength + " ký tự.");
                return false;
            }
            if (clean.Any(character => char.IsControl(character) || character == '\r' || character == '\n'))
            {
                failure = PersonalDictionaryResult.From(PersonalDictionaryStatus.InvalidInput,
                    "Từ hoặc cụm từ chứa ký tự điều khiển không hợp lệ.");
                return false;
            }
            failure = PersonalDictionaryResult.From(PersonalDictionaryStatus.Success);
            return true;
        }

        private static string? NormalizeDocumentId(string? documentId)
        {
            if (string.IsNullOrWhiteSpace(documentId)) return null;
            return documentId!.Trim().Normalize(NormalizationForm.FormC);
        }

        private static bool IsPersistenceException(Exception exception) =>
            exception is IOException || exception is UnauthorizedAccessException ||
            exception is DecoderFallbackException;

        private static PersonalDictionaryManager CreateShared()
        {
            var localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return new PersonalDictionaryManager(Path.Combine(localApp, "ChuanHoa", "Dictionaries"));
        }

        private void OnChanged()
        {
            var handler = Changed;
            if (handler != null) handler(this, EventArgs.Empty);
        }
    }
}
