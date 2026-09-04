#nullable disable
using System;
using System.Diagnostics;
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
            if (args.Length == 1 && string.Equals(args[0], "--synthetic-heavy",
                    StringComparison.OrdinalIgnoreCase))
            {
                return RunSyntheticHeavyMatrix();
            }
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

        private static int RunSyntheticHeavyMatrix()
        {
            var directory = Path.Combine(Path.GetTempPath(),
                "ChuanHoaHeavySnapshot-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            Word.Application application = null;
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                foreach (var targetPages in new[] { 10, 50, 100 })
                    VerifySyntheticHeavyDocument(application, directory, targetPages);
                Console.WriteLine("HEAVY_SNAPSHOT_WORD_SMOKE_PASS 10 50 100 PROGRESS CANCEL SECOND_SCAN SAVE_AS TABLES SHAPES SECTIONS");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("HEAVY_SNAPSHOT_WORD_SMOKE_FAIL: " + exception);
                return 1;
            }
            finally
            {
                if (application != null)
                    application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(application);
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        private static void VerifySyntheticHeavyDocument(Word.Application application,
            string directory, int targetPages)
        {
            Word.Document document = null;
            try
            {
                document = application.Documents.Add();
                BuildSyntheticPages(document, targetPages);
                var path = Path.Combine(directory, "heavy-" + targetPages + ".docx");
                object fileName = path;
                object format = Word.WdSaveFormat.wdFormatXMLDocument;
                document.SaveAs(ref fileName, ref format);
                document.Activate();

                var capability = new WordDocumentCapabilityProvider(application).Evaluate(document);
                Assert(capability.CanReadDocument, "Synthetic document is not readable.");
                var context = new DocumentContext(targetPages);
                var builder = new WordDocumentSnapshotBuilder();
                WordDocumentSnapshot first;
                var watch = Stopwatch.StartNew();
                using (var operation = new DocumentOperationSession(application,
                    "Synthetic " + targetPages, DocumentOperationState.Capturing))
                {
                    first = builder.Build(document, context, capability, operation);
                    Assert(operation.State == DocumentOperationState.Capturing,
                        "Progress operation left the capture state unexpectedly.");
                }
                watch.Stop();
                Assert(first.Revision == 1, "First heavy snapshot revision is invalid.");
                Assert(first.Paragraphs.Count >= targetPages * 2,
                    "Heavy snapshot lost paragraphs.");
                Assert(first.Tables.Count >= Math.Max(1, targetPages / 10),
                    "Heavy snapshot lost tables.");
                Assert(first.LineShapes.Count >= Math.Max(1, targetPages / 10),
                    "Heavy snapshot lost shapes.");
                Assert(first.Sections.Count >= Math.Max(1, targetPages / 20),
                    "Heavy snapshot lost sections.");

                WordDocumentSnapshot second;
                using (var operation = new DocumentOperationSession(application,
                    "Synthetic second scan " + targetPages, DocumentOperationState.Capturing))
                    second = builder.Build(document, context, capability, operation);
                Assert(second.Revision == 2, "Second heavy snapshot did not complete.");
                Assert(string.Equals(first.DocumentFingerprint, second.DocumentFingerprint,
                    StringComparison.Ordinal), "Read-only second scan changed snapshot identity.");

                var saveAsPath = Path.Combine(directory, "heavy-" + targetPages + "-save-as.docx");
                object saveAsName = saveAsPath;
                document.SaveAs(ref saveAsName, ref format);
                using (var operation = new DocumentOperationSession(application,
                    "Synthetic after Save As " + targetPages, DocumentOperationState.Capturing))
                {
                    var afterSaveAs = builder.Build(document, context,
                        new WordDocumentCapabilityProvider(application).Evaluate(document), operation);
                    Assert(afterSaveAs.Revision == 3, "Snapshot after Save As did not complete.");
                }

                var cancelled = false;
                using (var operation = new DocumentOperationSession(application,
                    "Synthetic cancellation " + targetPages, DocumentOperationState.Capturing))
                {
                    try
                    {
                        operation.RequestCancellation();
                        builder.Build(document, context, capability, operation);
                    }
                    catch (OperationCanceledException)
                    {
                        cancelled = true;
                    }
                }
                Assert(cancelled, "Pre-mutation cancellation was not honored.");
                Assert(document.Saved, "Read/progress/cancellation changed the document.");
                Console.WriteLine("HEAVY_SNAPSHOT_CASE_PASS pages=" + targetPages +
                    " paragraphs=" + first.Paragraphs.Count +
                    " tables=" + first.Tables.Count +
                    " shapes=" + first.LineShapes.Count +
                    " sections=" + first.Sections.Count +
                    " first_ms=" + watch.ElapsedMilliseconds);
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

        private static void BuildSyntheticPages(Word.Document document, int targetPages)
        {
            for (var page = 1; page <= targetPages; page++)
            {
                Word.Range range = null;
                try
                {
                    range = document.Range(document.Content.End - 1, document.Content.End - 1);
                    range.InsertAfter("Trang kiểm thử " + page + "\r" +
                        "Nội dung kiểm thử tài liệu lớn có bảng, hình và nhiều section.\r");
                }
                finally { Release(range); }

                if (page == 1 || page % 10 == 0)
                {
                    Word.Range tableRange = null;
                    Word.Table table = null;
                    Word.Range anchor = null;
                    Word.Shape shape = null;
                    try
                    {
                        tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                        table = document.Tables.Add(tableRange, 2, 3);
                        table.Cell(1, 1).Range.Text = "STT";
                        table.Cell(1, 2).Range.Text = "Nội dung";
                        table.Cell(1, 3).Range.Text = "Giá trị";
                        anchor = document.Range(document.Content.End - 1, document.Content.End - 1);
                        object anchorObject = anchor;
                        shape = document.Shapes.AddLine(72f, 72f, 150f, 72f, ref anchorObject);
                        shape.Name = "HEAVY_SMOKE_LINE_" + page;
                    }
                    finally
                    {
                        Release(shape);
                        Release(anchor);
                        Release(table);
                        Release(tableRange);
                    }
                }

                if (page >= targetPages) continue;
                Word.Range breakRange = null;
                try
                {
                    breakRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                    breakRange.InsertBreak(page % 20 == 0
                        ? Word.WdBreakType.wdSectionBreakNextPage
                        : Word.WdBreakType.wdPageBreak);
                }
                finally { Release(breakRange); }
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
