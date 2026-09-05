#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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
                    if (string.Equals(args[0], "--performance", StringComparison.OrdinalIgnoreCase))
                    {
                        RunPerformanceBenchmark();
                        Console.WriteLine("WORD_PERFORMANCE_SMOKE_PASS WARMUP=1 MEASURED=10");
                        return 0;
                    }
                    if (string.Equals(args[0], "--academic-on", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(args[0], "--academic-off", StringComparison.OrdinalIgnoreCase))
                    {
                        var expectedEnabled = string.Equals(args[0], "--academic-on",
                            StringComparison.OrdinalIgnoreCase);
                        RunAcademicTypographyPolicy(expectedEnabled);
                        Console.WriteLine("ACADEMIC_TYPOGRAPHY_WORD_SMOKE_PASS " +
                            (expectedEnabled ? "ENABLED" : "DISABLED"));
                        return 0;
                    }
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

        private static void RunPerformanceBenchmark()
        {
            const int measuredRounds = 10;
            Word.Application application = null;
            Word.Document document = null;
            Word.Table table = null;
            Process wordProcess = null;
            var wordProcessId = 0;
            var existingWordProcessIds = Process.GetProcessesByName("WINWORD")
                .Select(item => item.Id).ToArray();
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                var newWordProcesses = Process.GetProcessesByName("WINWORD")
                    .Where(item => !existingWordProcessIds.Contains(item.Id)).ToArray();
                Assert(newWordProcesses.Length == 1,
                    "The performance smoke could not identify exactly one isolated WINWORD process.");
                wordProcess = newWordProcesses[0];
                wordProcessId = wordProcess.Id;

                document = application.Documents.Add();
                var content = new StringBuilder();
                content.Append("CƠ QUAN BAN HÀNH\r")
                    .Append("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r")
                    .Append("Độc lập - Tự do - Hạnh phúc\r")
                    .Append("Số: 01/QĐ-CQ\r")
                    .Append("Hà Nội, ngày 05 tháng 09 năm 2026\r")
                    .Append("QUYẾT ĐỊNH\r")
                    .Append("Về việc kiểm tra hiệu năng\r")
                    .Append("Căn cứ quy định hiện hành.\r")
                    .Append("Điều 1. Tổ chức thực hiện.\r");
                for (var index = 0; index < 60; index++)
                {
                    content.Append("Khoản ").Append(index + 1)
                        .Append(" có nội dung kiểm tra ổn định, không tải tài liệu lên máy chủ.\r");
                }
                document.Content.Text = content.ToString();
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 13f;
                Word.Range tableInsertion = null;
                try
                {
                    tableInsertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                    table = document.Tables.Add(tableInsertion, 12, 4);
                    for (var column = 1; column <= 4; column++)
                        table.Cell(1, column).Range.Text = "Cột " + column;
                    for (var row = 2; row <= 12; row++)
                        for (var column = 1; column <= 4; column++)
                            table.Cell(row, column).Range.Text = row + "." + column;
                }
                finally { Release(tableInsertion); }

                var context = new DocumentContext(document.GetHashCode())
                {
                    RegimeCode = "ND30",
                    DocumentTypeCode = LocalDocumentTypeCodes.Decision,
                    RegimeWasSelectedManually = true,
                    DocumentTypeWasSelectedManually = false
                };
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    var currentRound = -1;
                    var snapshotBuilder = new WordDocumentSnapshotBuilder((stage, elapsedMilliseconds) =>
                        Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                            "WORD_PERF_STAGE round={0} stage={1} elapsed_ms={2}",
                            currentRound, stage, elapsedMilliseconds)));
                    var reader = new WordDocumentReadRuntime(application, access, snapshotBuilder);
                    var commands = new WordLocalCommandRuntime(application, access);
                    long warmupCapture;
                    long warmupCommand;
                    currentRound = 0;
                    ExecutePerformanceRound(document, context, reader, commands, 0,
                        out warmupCapture, out warmupCommand);

                    var durations = new List<long>();
                    var workingSets = new List<long>();
                    var privateBytes = new List<long>();
                    var handles = new List<int>();
                    var gdiHandles = new List<int>();
                    var userHandles = new List<int>();
                    for (var round = 1; round <= measuredRounds; round++)
                    {
                        currentRound = round;
                        var timer = Stopwatch.StartNew();
                        long captureMilliseconds;
                        long commandMilliseconds;
                        ExecutePerformanceRound(document, context, reader, commands, round,
                            out captureMilliseconds, out commandMilliseconds);
                        timer.Stop();
                        wordProcess.Refresh();
                        durations.Add(timer.ElapsedMilliseconds);
                        workingSets.Add(wordProcess.WorkingSet64);
                        privateBytes.Add(wordProcess.PrivateMemorySize64);
                        handles.Add(wordProcess.HandleCount);
                        gdiHandles.Add(GetGuiResources(wordProcess.Handle, 0));
                        userHandles.Add(GetGuiResources(wordProcess.Handle, 1));
                        Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                            "WORD_PERF_ROUND round={0} elapsed_ms={1} capture_scan_ms={2} command_ms={3} working_set={4} private_bytes={5} handles={6} gdi={7} user={8}",
                            round, timer.ElapsedMilliseconds, captureMilliseconds, commandMilliseconds,
                            wordProcess.WorkingSet64,
                            wordProcess.PrivateMemorySize64, wordProcess.HandleCount,
                            gdiHandles[gdiHandles.Count - 1], userHandles[userHandles.Count - 1]));
                    }

                    Assert(!IsStrictlyIncreasing(workingSets),
                        "Word working set increased strictly in every measured round.");
                    Assert(!IsStrictlyIncreasing(privateBytes),
                        "Word private bytes increased strictly in every measured round.");
                    Assert(!IsStrictlyIncreasing(handles.Select(value => (long)value).ToList()),
                        "Word handle count increased strictly in every measured round.");
                    Assert(!IsStrictlyIncreasing(gdiHandles.Select(value => (long)value).ToList()),
                        "Word GDI handles increased strictly in every measured round.");
                    Assert(!IsStrictlyIncreasing(userHandles.Select(value => (long)value).ToList()),
                        "Word USER handles increased strictly in every measured round.");
                    var ordered = durations.OrderBy(value => value).ToArray();
                    var median = (ordered[4] + ordered[5]) / 2d;
                    var p95 = ordered[ordered.Length - 1];
                    Assert(workingSets[workingSets.Count - 1] - workingSets.Take(3).Min() <= 64L * 1024L * 1024L,
                        "Word working set did not plateau within the 64 MiB repeated-scan budget.");
                    Assert(privateBytes[privateBytes.Count - 1] - privateBytes.Take(3).Min() <= 64L * 1024L * 1024L,
                        "Word private bytes did not plateau within the 64 MiB repeated-scan budget.");
                    Assert(handles[handles.Count - 1] - handles.Take(3).Min() <= 100,
                        "Word handle count did not plateau within the repeated-scan budget.");
                    Assert(gdiHandles[gdiHandles.Count - 1] - gdiHandles.Take(3).Min() <= 10,
                        "Word GDI handles did not plateau within the repeated-scan budget.");
                    Assert(userHandles[userHandles.Count - 1] - userHandles.Take(3).Min() <= 10,
                        "Word USER handles did not plateau within the repeated-scan budget.");
                    Assert(median <= 15000d,
                        "Word repeated-scan median exceeded the 15 second current-machine baseline.");
                    Assert(p95 <= 20000L,
                        "Word repeated-scan p95 exceeded the 20 second current-machine baseline.");
                    Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                        "WORD_PERF_SUMMARY rounds={0} median_ms={1:F1} p95_ms={2} baseline=current_machine_corpus_v1 " +
                        "working_set_first={3} working_set_last={4} private_bytes_first={5} private_bytes_last={6} " +
                        "handles_first={7} handles_last={8} gdi_first={9} gdi_last={10} user_first={11} user_last={12}",
                        measuredRounds, median, p95,
                        workingSets[0], workingSets[workingSets.Count - 1],
                        privateBytes[0], privateBytes[privateBytes.Count - 1],
                        handles[0], handles[handles.Count - 1],
                        gdiHandles[0], gdiHandles[gdiHandles.Count - 1],
                        userHandles[0], userHandles[userHandles.Count - 1]));
                }
            }
            finally
            {
                Release(table);
                if (document != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    document.Close(ref save);
                }
                if (application != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    application.Quit(ref save);
                }
                Release(document);
                Release(application);
                if (wordProcess != null)
                {
                    if (!wordProcess.HasExited)
                        Assert(wordProcess.WaitForExit(10000),
                            "WINWORD remained running after the performance smoke.");
                    wordProcess.Dispose();
                }
                else if (wordProcessId != 0)
                {
                    try
                    {
                        using (var remaining = Process.GetProcessById(wordProcessId))
                            Assert(remaining.WaitForExit(10000),
                                "WINWORD remained running after the performance smoke.");
                    }
                    catch (ArgumentException) { }
                }
            }
        }

        private static void ExecutePerformanceRound(Word.Document document,
            DocumentContext context, WordDocumentReadRuntime reader,
            WordLocalCommandRuntime commands, int round,
            out long captureMilliseconds, out long commandMilliseconds)
        {
            var timer = Stopwatch.StartNew();
            reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
            timer.Stop();
            captureMilliseconds = timer.ElapsedMilliseconds;
            Assert(context.LastFormatScan != null && context.LastSpellingScan != null,
                "A performance round did not complete both format and spelling analysis.");
            Word.Range selection = null;
            try
            {
                selection = document.Range(0, Math.Min(5, document.Content.End - 1));
                selection.Select();
                var reset = round % 2 == 0;
                timer.Restart();
                var result = commands.SetCharacterSpacing(reset ? 0f : 0.1f, reset);
                timer.Stop();
                commandMilliseconds = timer.ElapsedMilliseconds;
                Assert(string.IsNullOrWhiteSpace(result),
                    "A local character-spacing command unexpectedly created a backup.");
            }
            finally { Release(selection); }
        }

        private static bool IsStrictlyIncreasing(IReadOnlyList<long> values)
        {
            for (var index = 1; index < values.Count; index++)
                if (values[index] <= values[index - 1]) return false;
            return values.Count > 1;
        }

        [DllImport("user32.dll")]
        private static extern int GetGuiResources(IntPtr process, int flags);

        private static void RunAcademicTypographyPolicy(bool expectedEnabled)
        {
            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                document = application.Documents.Add();
                document.Content.Text =
                    "1. TỔNG QUAN\r" +
                    "Nội dung phân tích học thuật đủ dài để kiểm tra kiểm soát dòng mồ côi trong một đoạn văn bản thông thường.\r";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 13f;
                document.Paragraphs[1].Range.ParagraphFormat.KeepWithNext = 0;
                document.Paragraphs[2].Range.ParagraphFormat.WidowControl = 0;

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
                    var reader = new WordDocumentReadRuntime(application, access);
                    reader.Prepare(context, DocumentAnalysisScope.Format, document, false);
                    var scan = context.LastFormatScan;
                    Assert(scan != null, "AcademicTypography format scan was not created.");
                    Assert(scan.AcademicTypographyEnabled == expectedEnabled,
                        "Signed AcademicTypography policy state did not match the expected state.");

                    var academic = scan.Findings.Where(item =>
                        item.RuleCode.StartsWith("LATEX-", StringComparison.Ordinal)).ToArray();
                    if (!expectedEnabled)
                    {
                        Assert(academic.Length == 0,
                            "Disabled signed AcademicTypography policy still emitted LATEX findings.");
                        return;
                    }

                    var styleFinding = academic.FirstOrDefault(item =>
                        string.Equals(item.RuleCode, "LATEX-SEC-STYLE", StringComparison.Ordinal));
                    var keepFinding = academic.FirstOrDefault(item =>
                        string.Equals(item.RuleCode, "LATEX-PAGINATION-KEEP", StringComparison.Ordinal));
                    Assert(styleFinding != null, "Enabled policy did not emit the academic heading style suggestion.");
                    Assert(keepFinding != null, "Enabled policy did not emit the signed KeepWithNext suggestion.");

                    var snapshot = context.LastLocalSnapshot;
                    Assert(snapshot != null, "AcademicTypography snapshot was not retained for selected-fix smoke.");
                    var heading = snapshot.Paragraphs.Single(item => item.Index == 1);
                    var runtime = new WordOneClickRuntime(application, access);
                    var commentOnlyRejected = false;
                    try
                    {
                        runtime.ExecuteSelectedFinding(context, "format", styleFinding.FindingId,
                            heading.StoryType, heading.AbsoluteStart, heading.AbsoluteEnd, document);
                    }
                    catch (InvalidOperationException exception)
                    {
                        commentOnlyRejected = exception.Message.IndexOf("Comment được giữ nguyên",
                            StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    Assert(commentOnlyRejected,
                        "LATEX-SEC-STYLE was not kept comment-only by selected fix.");
                    Assert(document.Paragraphs[1].Range.ParagraphFormat.KeepWithNext == 0,
                        "A comment-only academic suggestion mutated the heading.");

                    reader.Prepare(context, DocumentAnalysisScope.Format, document, false);
                    scan = context.LastFormatScan;
                    keepFinding = scan.Findings.FirstOrDefault(item =>
                        string.Equals(item.RuleCode, "LATEX-PAGINATION-KEEP", StringComparison.Ordinal));
                    Assert(keepFinding != null, "KeepWithNext finding disappeared before its safe-fix smoke.");
                    snapshot = context.LastLocalSnapshot;
                    heading = snapshot.Paragraphs.Single(item => item.Index == 1);
                    var fixedResult = runtime.ExecuteSelectedFinding(context, "format", keepFinding.FindingId,
                        heading.StoryType, heading.AbsoluteStart, heading.AbsoluteEnd, document);
                    Assert(fixedResult.Resolved,
                        "Signed allowlisted KeepWithNext selected fix did not pass post-scan verification.");
                    Assert(document.Paragraphs[1].Range.ParagraphFormat.KeepWithNext != 0,
                        "Signed allowlisted KeepWithNext fix did not mutate the Word paragraph.");
                }
            }
            finally
            {
                if (document != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    document.Close(ref save);
                }
                if (application != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    application.Quit(ref save);
                }
                Release(document);
                Release(application);
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
