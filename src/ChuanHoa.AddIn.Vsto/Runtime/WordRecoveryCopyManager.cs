using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// Saves through Word and creates the recovery copy from a hidden Word clone. This
    /// deliberately avoids File.Exists/File.Copy on Document.FullName because cloud-backed
    /// documents can expose an https/OneDrive identity that is valid to Word but not System.IO.
    /// </summary>
    internal static class WordRecoveryCopyManager
    {
        private const int RetentionDays = 7;
        private const int MaximumBackupCount = 20;

        public static void EnsurePersistentDocument(Word.Application application, Word.Document document)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            if (document == null) throw new ArgumentNullException(nameof(document));

            if (!string.IsNullOrWhiteSpace(ReadPath(document)) && !document.Saved)
            {
                document.Save();
            }
        }

        public static string Create(Word.Application application, Word.Document document, string prefix)
        {
            EnsurePersistentDocument(application, document);

            var directory = GetBackupDirectory();
            Directory.CreateDirectory(directory);
            CleanupOwnedBackups(directory);

            var extension = ReadSupportedExtension(document);
            var shortPrefix = string.Equals(prefix, "one-click", StringComparison.Ordinal)
                ? "oc"
                : string.Equals(prefix, "tone", StringComparison.Ordinal) ? "tone" : "mutation";
            var path = Path.Combine(directory,
                "chuanhoa-" + shortPrefix + "-" +
                DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture) +
                "-" + Guid.NewGuid().ToString("N").Substring(0, 12) + extension);

            SaveThroughWordClone(application, document, path);
            if (!File.Exists(path) || new FileInfo(path).Length == 0)
                throw new IOException("Word không tạo được bản sao khôi phục.");
            CleanupOwnedBackups(directory);
            return path;
        }

        internal static string GetBackupDirectory()
        {
            return Path.Combine(Path.GetTempPath(), "ChuanHoa", "Backups");
        }

        internal static void CleanupOwnedBackups(string directory)
        {
            try
            {
                if (!Directory.Exists(directory)) return;

                var cutoff = DateTime.UtcNow.AddDays(-RetentionDays);
                foreach (var file in OwnedBackupFiles(directory))
                {
                    try
                    {
                        if (File.GetLastWriteTimeUtc(file) < cutoff) File.Delete(file);
                    }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }

                foreach (var file in OwnedBackupFiles(directory)
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .Skip(MaximumBackupCount))
                {
                    try { File.Delete(file); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static string[] OwnedBackupFiles(string directory)
        {
            try
            {
                return Directory.GetFiles(directory, "chuanhoa-*", SearchOption.TopDirectoryOnly)
                    .Where(path =>
                    {
                        var name = Path.GetFileName(path);
                        var extension = Path.GetExtension(path);
                        var ownedPrefix = name.StartsWith("chuanhoa-oc-", StringComparison.OrdinalIgnoreCase) ||
                            name.StartsWith("chuanhoa-tone-", StringComparison.OrdinalIgnoreCase) ||
                            name.StartsWith("chuanhoa-mutation-", StringComparison.OrdinalIgnoreCase);
                        return ownedPrefix &&
                            (string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase));
                    })
                    .ToArray();
            }
            catch (IOException) { return Array.Empty<string>(); }
            catch (UnauthorizedAccessException) { return Array.Empty<string>(); }
        }

        private static string ReadPath(Word.Document document)
        {
            try { return document.Path ?? string.Empty; }
            catch (COMException) { return string.Empty; }
        }

        private static void SaveThroughWordClone(Word.Application application, Word.Document source, string path)
        {
            Word.Document? copy = null;
            var wordTemporaryPath = Path.Combine(Path.GetTempPath(),
                "ch-" + Guid.NewGuid().ToString("N") + Path.GetExtension(path));
            try
            {
                var sourcePath = ReadPath(source);
                var isSaved = !string.IsNullOrWhiteSpace(sourcePath);
                if (isSaved)
                {
                    object template = source.FullName;
                    object newTemplate = false;
                    object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                    object visible = false;
                    copy = application.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                }
                else
                {
                    object missing = Type.Missing;
                    object newTemplate = false;
                    object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                    object visible = false;
                    copy = application.Documents.Add(ref missing, ref newTemplate, ref documentType, ref visible);
                    try
                    {
                        copy.Content.FormattedText = source.Content.FormattedText;
                    }
                    catch (COMException)
                    {
                        // Fallback to text copy if formatted text copy encounters transient COM error
                        copy.Content.Text = source.Content.Text;
                    }
                }

                object backupFileName = wordTemporaryPath;
                object backupFileFormat = (int)source.SaveFormat == 0
                    ? Word.WdSaveFormat.wdFormatDocument
                    : Word.WdSaveFormat.wdFormatXMLDocument;
                try
                {
                    copy.SaveAs(ref backupFileName, ref backupFileFormat);
                }
                catch (COMException exception)
                {
                    throw new IOException("Word không lưu được bản clone khôi phục tại '" + wordTemporaryPath +
                        "' (định dạng " + Convert.ToString(source.SaveFormat, CultureInfo.InvariantCulture) +
                        ", clone '" + copy.Name + "').", exception);
                }
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        object doNotSave = Word.WdSaveOptions.wdDoNotSaveChanges;
                        copy.Close(ref doNotSave);
                    }
                    finally { Release(copy); }
                }

                // Creating a document from the source as a hidden template can leave
                // cloud-backed Word sessions without an ActiveDocument after the clone
                // closes. Restore the user's source window explicitly so the next Ribbon
                // command does not see a false "no document" state.
                try
                {
                    if (source.Windows.Count > 0)
                    {
                        source.Activate();
                    }
                }
                catch (COMException)
                {
                    // The caller continues with its captured source reference. This can
                    // occur briefly while AutoSave is completing; it is not evidence that
                    // the document was closed.
                }
            }

            try
            {
                File.Move(wordTemporaryPath, path);
            }
            finally
            {
                try { if (File.Exists(wordTemporaryPath)) File.Delete(wordTemporaryPath); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        private static string ReadSupportedExtension(Word.Document document)
        {
            string extension;
            try { extension = Path.GetExtension(document.Name ?? string.Empty); }
            catch (ArgumentException) { extension = string.Empty; }

            if (string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase)) return ".doc";
            if (string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase)) return ".docx";

            switch ((int)document.SaveFormat)
            {
                case (int)Word.WdSaveFormat.wdFormatDocument:
                    return ".doc";
                case (int)Word.WdSaveFormat.wdFormatXMLDocument:
                case (int)Word.WdSaveFormat.wdFormatDocumentDefault:
                case (int)Word.WdSaveFormat.wdFormatStrictOpenXMLDocument:
                    return ".docx";
                default:
                    throw new InvalidOperationException("Chuẩn hóa chỉ xử lý trực tiếp tài liệu .doc và .docx.");
            }
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
