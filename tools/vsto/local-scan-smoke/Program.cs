#nullable disable
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Scanning;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.LocalScanSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var directory = Path.Combine(Path.GetTempPath(), "ChuanHoaLocalScanSmoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                if (args.Length > 0)
                {
                    RunExistingDocument(args[0], directory);
                    Console.WriteLine("LOCAL_SCAN_EXISTING_DOCUMENT_PASS");
                    return 0;
                }
                Run(Path.Combine(directory, "sample.docx"), Word.WdSaveFormat.wdFormatXMLDocument);
                Run(Path.Combine(directory, "sample.doc"), Word.WdSaveFormat.wdFormatDocument97);
                Console.WriteLine("LOCAL_SCAN_WORD_SMOKE_PASS DOC DOCX");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("LOCAL_SCAN_WORD_SMOKE_FAIL " + exception);
                return 1;
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        private static void RunExistingDocument(string sourcePath, string directory)
        {
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("Smoke source document was not found.", sourcePath);
            var workingPath = Path.Combine(directory, Path.GetFileName(sourcePath));
            File.Copy(sourcePath, workingPath, true);

            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Open(workingPath, ReadOnly: false, AddToRecentFiles: false);
                var context = new DocumentContext(document.GetHashCode())
                {
                    RegimeCode = "ND30",
                    DocumentTypeCode = LocalDocumentTypeCodes.Unknown,
                    RegimeWasSelectedManually = true,
                    DocumentTypeWasSelectedManually = false
                };
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    var readTimer = Stopwatch.StartNew();
                    var read = new WordDocumentReadRuntime(application, access).Read(context, document);
                    readTimer.Stop();
                    Console.WriteLine("EXISTING_READ paragraphs=" + read.Snapshot.Paragraphs.Count +
                        " format_findings=" + read.FormatScan.Findings.Count +
                        " spelling_findings=" + read.SpellingScan.Findings.Count +
                        " content_end=" + document.Content.End +
                        " elapsed_ms=" + readTimer.ElapsedMilliseconds);
                    foreach (var paragraph in read.Snapshot.Paragraphs.Where(item =>
                        string.Equals(item.StoryType, Word.WdStoryType.wdMainTextStory.ToString(), StringComparison.Ordinal) &&
                        (item.AbsoluteStart < document.Content.Start ||
                         item.AbsoluteStart + item.Text.Length > document.Content.End)))
                    {
                        Console.WriteLine("INVALID_MAIN_PARAGRAPH index=" + paragraph.Index +
                            " section=" + paragraph.SectionIndex +
                            " start=" + paragraph.AbsoluteStart +
                            " text_length=" + paragraph.Text.Length +
                            " text=" + paragraph.Text.Replace("\r", "<CR>").Replace("\a", "<CELL>"));
                    }

                    var runtime = new WordLocalScanRuntime(application, access);
                    var formatTimer = Stopwatch.StartNew();
                    runtime.ScanAndAnnotate(context, false, document);
                    formatTimer.Stop();
                    Console.WriteLine("EXISTING_FORMAT_ANNOTATION elapsed_ms=" + formatTimer.ElapsedMilliseconds);

                    var spellingTimer = Stopwatch.StartNew();
                    runtime.ScanAndAnnotate(context, true, document);
                    spellingTimer.Stop();
                    Console.WriteLine("EXISTING_SPELLING_ANNOTATION elapsed_ms=" + spellingTimer.ElapsedMilliseconds);
                }
            }
            finally
            {
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                Release(application);
            }
        }

        private static void Run(string path, Word.WdSaveFormat format)
        {
            Word.Application application = null;
            Word.Document document = null;
            Word.Range typoRange = null;
            Word.Range lineAnchor = null;
            Word.Shape lineShape = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Add();
                document.Content.Text = "Đây  là văn bản sát nhập , cần kiểm tra.\r" +
                    "QUYẾT ĐỊNH\rVề việc kiểm tra nhận diện tự động\r" +
                    "Độc lập - Tự do - Hạnh phúc\r\r";
                document.Content.Font.Name = "Arial";
                document.Paragraphs[4].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                document.Sections[1].PageSetup.LeftMargin = 10f;
                object fileName = path;
                object saveFormat = format;
                document.SaveAs(ref fileName, ref saveFormat);

                var context = new DocumentContext(document.GetHashCode())
                {
                    RegimeCode = "ND30",
                    DocumentTypeCode = LocalDocumentTypeCodes.Unknown,
                    RegimeWasSelectedManually = true,
                    DocumentTypeWasSelectedManually = false
                };
                var capability = new WordDocumentCapabilityProvider(application).Evaluate();
                var snapshotWithoutLine = new WordDocumentSnapshotBuilder().Build(document, context, capability);
                var motto = snapshotWithoutLine.Paragraphs.Single(item =>
                    item.Text.IndexOf("Độc lập", StringComparison.Ordinal) >= 0);
                Assert(motto.PageLeftPoints.HasValue && motto.TextWidthPoints.HasValue,
                    "Motto rendered geometry was not captured.");

                lineAnchor = document.Paragraphs[5].Range.Duplicate;
                lineAnchor.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                object anchor = lineAnchor;
                lineShape = document.Shapes.AddLine(72f, 18f, 180f, 18f, ref anchor);
                lineShape.LayoutInCell = 0;
                lineShape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                lineShape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                lineShape.Left = (float)motto.PageLeftPoints.Value;
                lineShape.Top = (float)(motto.PageTopPoints.Value + 18d);
                lineShape.Width = (float)motto.TextWidthPoints.Value;
                document.Save();

                var snapshot = new WordDocumentSnapshotBuilder().Build(document, context, capability);
                Assert(snapshot.LineShapes.Count == 1, "Line Shape was not captured.");
                Assert(snapshot.LineShapes[0].ShapeType == 9 && snapshot.LineShapes[0].LineVisible,
                    "Captured shape is not a visible msoLine.");
                Assert(snapshot.LineShapes[0].AnchorParagraphIndex.HasValue,
                    "Line Shape anchor was not mapped to a paragraph.");
                var positionedSnapshot = snapshot;
                var positionedMotto = positionedSnapshot.Paragraphs.Single(item =>
                    item.Text.IndexOf("Độc lập", StringComparison.Ordinal) >= 0);
                var positionedLine = positionedSnapshot.LineShapes.Single();
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    var runtime = new WordLocalScanRuntime(application, access);
                    new WordDocumentReadRuntime(application, access).Read(context, document);
                    var firstRunTimer = Stopwatch.StartNew();
                    var formatResult = runtime.ScanAndAnnotate(context, false);
                    firstRunTimer.Stop();

                    // Reproduce the real user sequence: fix one reported problem,
                    // save, then scan the same document again.
                    document.Content.Font.Name = "Times New Roman";
                    document.Save();
                    new WordDocumentReadRuntime(application, access).Read(context, document);
                    var rerunTimer = Stopwatch.StartNew();
                    var rerunFormatResult = runtime.ScanAndAnnotate(context, false);
                    rerunTimer.Stop();

                    var spellingResult = runtime.ScanAndAnnotate(context, true);
                    Assert(context.DocumentTypeCode == LocalDocumentTypeCodes.Decision &&
                        !context.DocumentTypeWasSelectedManually,
                        "Document type was not detected automatically as Decision.");
                    Assert(formatResult.Findings.Any(item => item.RuleCode == "ND30-PL1-M1-K3"), "Margin finding missing.");
                    Assert(formatResult.Findings.Any(item => item.RuleCode == "ND30-PL1-M1-K4-FONT"), "Font finding missing.");
                    Assert(!rerunFormatResult.Findings.Any(item => item.RuleCode == "ND30-PL1-M1-K4-FONT"),
                        "The second scan did not observe the saved font correction.");
                    Assert(rerunTimer.Elapsed < TimeSpan.FromSeconds(15),
                        "The second format scan took too long: " + rerunTimer.Elapsed + ".");
                    Console.WriteLine(
                        "LOCAL_SCAN_RERUN_TIMING format=" + format +
                        " findings_before=" + formatResult.Findings.Count +
                        " findings_after=" + rerunFormatResult.Findings.Count +
                        " first_ms=" + firstRunTimer.ElapsedMilliseconds +
                        " elapsed_ms=" + rerunTimer.ElapsedMilliseconds);
                    var lineFinding = formatResult.Findings.FirstOrDefault(item =>
                        item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
                    Assert(lineFinding == null,
                        "A real Line Shape anchored to the following blank paragraph was rejected: " +
                        (lineFinding == null ? string.Empty : lineFinding.CurrentIssue) +
                        " Motto(index/page/left/top/width)=" + positionedMotto.Index + "/" +
                        positionedMotto.PageNumber + "/" + positionedMotto.PageLeftPoints + "/" +
                        positionedMotto.PageTopPoints + "/" + positionedMotto.TextWidthPoints +
                        " Line(anchor/page/left/top/width/height/relative)=" +
                        positionedLine.AnchorParagraphIndex + "/" + positionedLine.AnchorPageNumber + "/" +
                        positionedLine.PageLeftPoints + "/" + positionedLine.PageTopPoints + "/" +
                        positionedLine.WidthPoints + "/" + positionedLine.HeightPoints + "/" +
                        positionedLine.RelativeHorizontalPosition + ":" + positionedLine.RelativeVerticalPosition);
                    Assert(spellingResult.Findings.Any(item => item.RuleCode == "LOCAL-TYPO-DICT"), "Dictionary finding missing.");
                    Assert(spellingResult.Findings.Any(item => item.RuleCode == "LOCAL-TYPO-SPACE"), "Whitespace finding missing.");
                    Assert(spellingResult.Findings.Any(item => item.RuleCode == "LOCAL-TYPO-PUNCT"), "Punctuation finding missing.");
                }

                var ownedVariables = Enumerable.Range(1, document.Variables.Count)
                    .Select(index => document.Variables[index].Name ?? string.Empty).ToArray();
                Assert(ownedVariables.Any(name => name.StartsWith("CHCOM_FORMAT_", StringComparison.Ordinal)),
                    "Format comments missing. Variables=" + string.Join(",", ownedVariables));
                Assert(ownedVariables.Any(name => name.StartsWith("CHCOM_SPELLING_", StringComparison.Ordinal)),
                    "Spelling comments missing. Variables=" + string.Join(",", ownedVariables));
                var start = document.Content.Text.IndexOf("sát nhập", StringComparison.Ordinal);
                typoRange = document.Range(start, start + "sát nhập".Length);
                Assert(typoRange.Font.Color == Word.WdColor.wdColorRed, "Exact spelling error is not red.");
            }
            finally
            {
                Release(lineShape);
                Release(lineAnchor);
                Release(typoRange);
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) application.Quit();
                Release(document);
                Release(application);
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
