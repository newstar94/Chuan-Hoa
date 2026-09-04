#nullable disable
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using ChuanHoa.AddIn.Vsto.Runtime;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.SnapshotSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.Error.WriteLine("SNAPSHOT_WORD_SMOKE_FAIL: At least one .doc or .docx path is required.");
                return 2;
            }

            Word.Application application = null;
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                foreach (var path in args)
                {
                    VerifySnapshot(application, Path.GetFullPath(path));
                }

                VerifyUnsavedDocumentIsSupported(application);
                Console.WriteLine("SNAPSHOT_WORD_SMOKE_PASS");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("SNAPSHOT_WORD_SMOKE_FAIL: " + exception);
                return 1;
            }
            finally
            {
                if (application != null)
                {
                    application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                }

                Release(application);
            }
        }

        private static void VerifySnapshot(Word.Application application, string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("Golden document was not found.", path);
            }

            var beforeHash = ComputeFileSha256(path);
            var beforeWrite = File.GetLastWriteTimeUtc(path);
            Word.Document document = null;
            try
            {
                object fileName = path;
                object readOnly = false;
                object visible = false;
                document = application.Documents.Open(ref fileName, ReadOnly: ref readOnly, Visible: ref visible);
                document.Activate();

                var capability = new WordDocumentCapabilityProvider(application).Evaluate();
                Assert(capability.CanReadDocument, "Expected a readable supported document: " + capability.ReasonCode);
                var context = new DocumentContext(path.GetHashCode())
                {
                    RegimeCode = "ND30",
                    DocumentTypeCode = "1"
                };
                var snapshot = new WordDocumentSnapshotBuilder().Build(document, context, capability);
                Assert(snapshot.SchemaVersion == 2, "Unexpected snapshot schema.");
                Assert(snapshot.DocumentFingerprint.StartsWith("sha256:", StringComparison.Ordinal),
                    "Snapshot fingerprint is missing.");
                Assert(snapshot.Revision == 1, "First snapshot revision must be one.");
                Assert(snapshot.Sections.Count > 0, "Document sections were not captured.");
                Assert(snapshot.Paragraphs.Count > 0, "Document paragraphs were not captured.");
                VerifyMainStoryCoordinates(document, snapshot);
                Assert(string.Equals(snapshot.Preflight.FileFormat, Path.GetExtension(path),
                    StringComparison.OrdinalIgnoreCase), "Snapshot file format does not match the source.");
            }
            finally
            {
                if (document != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    document.Close(ref save);
                }

                Release(document);
            }

            Assert(string.Equals(beforeHash, ComputeFileSha256(path), StringComparison.Ordinal),
                "Snapshot changed the source document bytes.");
            Assert(beforeWrite == File.GetLastWriteTimeUtc(path),
                "Snapshot changed the source document timestamp.");
        }

        private static void VerifyMainStoryCoordinates(Word.Document document, WordDocumentSnapshot snapshot)
        {
            Word.Range content = null;
            try
            {
                content = document.Content;
                var mainParagraphs = snapshot.Paragraphs.Where(item =>
                    string.Equals(item.StoryType, Word.WdStoryType.wdMainTextStory.ToString(),
                        StringComparison.Ordinal)).ToArray();
                Assert(mainParagraphs.Length == document.Paragraphs.Count ||
                    mainParagraphs.Length + 1 == document.Paragraphs.Count,
                    "Snapshot main-story paragraph count does not match Word: snapshot=" +
                    mainParagraphs.Length + ", Word=" + document.Paragraphs.Count + ".");
                foreach (var paragraph in mainParagraphs)
                {
                    Assert(paragraph.AbsoluteStart >= content.Start &&
                        paragraph.AbsoluteStart + paragraph.Text.Length <= content.End,
                        "Snapshot paragraph is outside the authoritative Word main-story range: P" +
                        paragraph.Index + ".");
                }
            }
            finally
            {
                Release(content);
            }
        }

        private static void VerifyUnsavedDocumentIsSupported(Word.Application application)
        {
            Word.Document document = null;
            try
            {
                document = application.Documents.Add();
                document.Activate();
                var capability = new WordDocumentCapabilityProvider(application).Evaluate();
                Assert(capability.CanReadDocument, "Unsaved document must be readable.");
                Assert(!capability.IsSaved, "Unsaved document must report IsSaved as false.");
                Assert(string.Equals(capability.ReasonCode, "READY", StringComparison.Ordinal),
                    "Unsaved document returned an unexpected reason code.");
            }
            finally
            {
                if (document != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    document.Close(ref save);
                }

                Release(document);
            }
        }

        private static string ComputeFileSha256(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha256 = SHA256.Create())
            {
                return Convert.ToBase64String(sha256.ComputeHash(stream));
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }
    }
}
