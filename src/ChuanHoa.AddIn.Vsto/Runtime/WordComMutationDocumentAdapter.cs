using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using ChuanHoa.Client.Core.Safety;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class WordComMutationDocumentAdapter : IMutationDocumentAdapter, IDisposable
    {
        private const string SetCharacterSpacingOperation = "SetCharacterSpacing";
        private const string SelectedTextScope = "selected-text";
        private readonly Word.Application _application;
        private readonly Word.Document _document;
        private readonly long _documentIdentity;
        private readonly bool _supportsCustomUndoRecord;
        private int _appliedOperationCount;
        private bool _undoRecordStarted;
        private bool _disposed;

        public WordComMutationDocumentAdapter(Word.Application application, Word.Document document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _documentIdentity = ReadComIdentity(document);
            _supportsCustomUndoRecord = ReadWordMajorVersion(application) >= 15;
        }

        public string DocumentIdentity => _documentIdentity.ToString("X16", CultureInfo.InvariantCulture);

        public DocumentMutationPreflight ReadPreflight()
        {
            ThrowIfDisposed();
            var hasActiveDocument = false;
            var hasActiveWindow = false;
            try
            {
                hasActiveDocument = ReadComIdentity(_application.ActiveDocument) == _documentIdentity;
                hasActiveWindow = _application.ActiveWindow != null &&
                    ReadComIdentity(_application.ActiveWindow.Document) == _documentIdentity;
            }
            catch (COMException)
            {
            }

            var isReadOnly = _document.ReadOnly;
            var isProtected =
                _document.ProtectionType != Word.WdProtectionType.wdNoProtection ||
                _document.TrackRevisions;
            return new DocumentMutationPreflight(
                hasActiveDocument,
                hasActiveWindow,
                isReadOnly,
                isProtected,
                false,
                CanCreateBackup(),
                IsSupportedDocumentFormat());
        }

        public string CaptureFingerprint()
        {
            ThrowIfDisposed();
            using (var sha256 = SHA256.Create())
            {
                HashString(sha256, "chuanhoa-word-selection-v1");
                HashString(sha256, CaptureDocumentContentFingerprint());
                HashSelection(sha256);
                sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return "sha256:" + ToHex(sha256.Hash!);
            }
        }

        public object CaptureApplicationState()
        {
            ThrowIfDisposed();
            return new ApplicationState(
                _application.ScreenUpdating,
                _application.DisplayAlerts,
                CaptureSelection());
        }

        public object CreateBackup()
        {
            ThrowIfDisposed();
            if (!CanCreateBackup())
            {
                throw new IOException("The active document cannot be backed up safely.");
            }

            var backupPath = WordRecoveryCopyManager.Create(_application, _document, "mutation");
            if (!File.Exists(backupPath) || new FileInfo(backupPath).Length == 0)
            {
                throw new IOException("Word did not create a usable backup copy.");
            }

            return new BackupArtifact(
                backupPath,
                CaptureDocumentContentFingerprint(),
                ComputeFileSha256(backupPath));
        }

        public void BeginUndoRecord(string commandId)
        {
            ThrowIfDisposed();
            if (_undoRecordStarted)
            {
                throw new InvalidOperationException("An undo record is already active.");
            }

            if (_supportsCustomUndoRecord)
            {
                _application.UndoRecord.StartCustomRecord("Chuẩn hóa: " + commandId);
            }
            _undoRecordStarted = true;
        }

        public void Apply(MutationOperation operation)
        {
            ThrowIfDisposed();
            EnsureSupportedOperation(operation);
            var range = CaptureSelectedRange();
            try
            {
                range.Font.Spacing = 0.0f;
                _appliedOperationCount++;
            }
            finally
            {
                Marshal.ReleaseComObject(range);
            }
        }

        public bool Verify(MutationOperation operation)
        {
            ThrowIfDisposed();
            EnsureSupportedOperation(operation);
            var range = CaptureSelectedRange();
            try
            {
                return Math.Abs(range.Font.Spacing) < 0.001f;
            }
            finally
            {
                Marshal.ReleaseComObject(range);
            }
        }

        public void EndUndoRecord()
        {
            ThrowIfDisposed();
            if (_undoRecordStarted)
            {
                if (_supportsCustomUndoRecord)
                {
                    _application.UndoRecord.EndCustomRecord();
                }
                _undoRecordStarted = false;
            }
        }

        public void Rollback(object? backup)
        {
            ThrowIfDisposed();
            var artifact = backup as BackupArtifact;
            if (artifact == null || !File.Exists(artifact.Path))
            {
                throw new IOException("A verified backup is required for rollback.");
            }

            if (!string.Equals(
                ComputeFileSha256(artifact.Path),
                artifact.BackupSha256,
                StringComparison.Ordinal))
            {
                throw new IOException("The recovery copy changed after it was created.");
            }

            object times = _supportsCustomUndoRecord ? 1 : Math.Max(1, _appliedOperationCount);
            if (!_document.Undo(ref times))
            {
                throw new InvalidOperationException("Word could not undo the failed mutation.");
            }

            if (!string.Equals(
                CaptureDocumentContentFingerprint(),
                artifact.OriginalDocumentFingerprint,
                StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Word reported Undo success but the document fingerprint was not restored. " +
                    "The recovery copy is retained at " + artifact.Path + ".");
            }
        }

        public void RestoreApplicationState(object applicationState)
        {
            ThrowIfDisposed();
            var state = applicationState as ApplicationState;
            if (state == null)
            {
                throw new ArgumentException("Unknown Word application state.", nameof(applicationState));
            }

            try
            {
                _application.ScreenUpdating = state.ScreenUpdating;
                _application.DisplayAlerts = state.DisplayAlerts;
                if (state.Selection != null)
                {
                    state.Selection.Select();
                }
            }
            finally
            {
                state.Dispose();
            }
        }

        public void Dispose()
        {
            _disposed = true;
        }

        private bool CanCreateBackup()
        {
            try
            {
                return !string.IsNullOrWhiteSpace(_document.Path) &&
                    File.Exists(_document.FullName) &&
                    !_document.ReadOnly &&
                    _document.Saved;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private string ReadDocumentFullName()
        {
            try
            {
                return _document.FullName ?? string.Empty;
            }
            catch (COMException)
            {
                return string.Empty;
            }
        }

        private bool IsSupportedDocumentFormat()
        {
            try
            {
                return SupportedWordDocumentFormatPolicy.IsSupported(
                    ReadDocumentFullName(),
                    (int)_document.SaveFormat);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private string CaptureDocumentContentFingerprint()
        {
            using (var sha256 = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(_document.WordOpenXML);
                return "sha256:" + ToHex(sha256.ComputeHash(bytes));
            }
        }

        private Word.Range CaptureSelectedRange()
        {
            var selection = _application.Selection;
            if (selection == null || selection.Document == null ||
                ReadComIdentity(selection.Document) != _documentIdentity)
            {
                throw new InvalidOperationException("The active selection is not in the authorized document.");
            }

            if (selection.Start == selection.End)
            {
                throw new InvalidOperationException("A non-empty selection is required.");
            }

            return selection.Range.Duplicate;
        }

        private Word.Range? CaptureSelection()
        {
            try
            {
                var selection = _application.Selection;
                return selection == null ? null : selection.Range.Duplicate;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private void HashSelection(HashAlgorithm hash)
        {
            Word.Range? selection = null;
            try
            {
                selection = CaptureSelectedRange();
                HashString(hash, selection.StoryType.ToString());
                HashString(hash, selection.Start.ToString(CultureInfo.InvariantCulture));
                HashString(hash, selection.End.ToString(CultureInfo.InvariantCulture));
                HashString(hash, selection.Text ?? string.Empty);
            }
            finally
            {
                if (selection != null)
                {
                    Marshal.ReleaseComObject(selection);
                }
            }
        }

        private static void HashString(HashAlgorithm hash, string value)
        {
            var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            var length = BitConverter.GetBytes(bytes.Length);
            hash.TransformBlock(length, 0, length.Length, null, 0);
            hash.TransformBlock(bytes, 0, bytes.Length, null, 0);
        }

        private static string ToHex(byte[] bytes)
        {
            var builder = new StringBuilder(bytes.Length * 2);
            foreach (var value in bytes)
            {
                builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        private static string ComputeFileSha256(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var sha256 = SHA256.Create())
            {
                return ToHex(sha256.ComputeHash(stream));
            }
        }

        private static long ReadComIdentity(object comObject)
        {
            if (comObject == null)
            {
                return 0;
            }

            var unknown = Marshal.GetIUnknownForObject(comObject);
            try
            {
                return unknown.ToInt64();
            }
            finally
            {
                Marshal.Release(unknown);
            }
        }

        private static int ReadWordMajorVersion(Word.Application application)
        {
            var version = application.Version ?? string.Empty;
            var separator = version.IndexOf('.');
            var majorText = separator < 0 ? version : version.Substring(0, separator);
            int major;
            return int.TryParse(majorText, NumberStyles.None, CultureInfo.InvariantCulture, out major)
                ? major
                : 0;
        }

        private static void EnsureSupportedOperation(MutationOperation operation)
        {
            if (operation == null ||
                !string.Equals(operation.OperationType, SetCharacterSpacingOperation, StringComparison.Ordinal) ||
                !string.Equals(operation.Scope, SelectedTextScope, StringComparison.Ordinal) ||
                operation.RiskTier != MutationRiskTier.Confirm)
            {
                throw new NotSupportedException("The Word adapter operation is not allowlisted.");
            }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }

        private sealed class BackupArtifact
        {
            public BackupArtifact(
                string path,
                string originalDocumentFingerprint,
                string backupSha256)
            {
                Path = path;
                OriginalDocumentFingerprint = originalDocumentFingerprint;
                BackupSha256 = backupSha256;
            }

            public string Path { get; }

            public string OriginalDocumentFingerprint { get; }

            public string BackupSha256 { get; }
        }

        private sealed class ApplicationState : IDisposable
        {
            public ApplicationState(
                bool screenUpdating,
                Word.WdAlertLevel displayAlerts,
                Word.Range? selection)
            {
                ScreenUpdating = screenUpdating;
                DisplayAlerts = displayAlerts;
                Selection = selection;
            }

            public bool ScreenUpdating { get; }

            public Word.WdAlertLevel DisplayAlerts { get; }

            public Word.Range? Selection { get; }

            public void Dispose()
            {
                if (Selection != null)
                {
                    Marshal.ReleaseComObject(Selection);
                }
            }
        }
    }
}
