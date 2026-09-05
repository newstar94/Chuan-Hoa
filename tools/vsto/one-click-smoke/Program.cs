#nullable disable
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Scanning;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.OneClickSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var directory = Path.Combine(Path.GetTempPath(), "ChuanHoaOneClickSmoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                if (args.Length > 0)
                {
                    if (string.Equals(args[0], "--multiple", StringComparison.OrdinalIgnoreCase))
                    {
                        RunMultipleEmbeddedDocuments(Path.Combine(directory, "multiple-documents.docx"));
                        Console.WriteLine("ONE_CLICK_MULTIPLE_DOCUMENTS_PASS");
                        return 0;
                    }
                    if (string.Equals(args[0], "--dashed-line", StringComparison.OrdinalIgnoreCase))
                    {
                        RunDashedMottoLineDetection(Path.Combine(directory, "dashed-motto-line.docx"));
                        Console.WriteLine("ONE_CLICK_DASHED_MOTTO_LINE_PASS");
                        return 0;
                    }
                    if (string.Equals(args[0], "--quick-spelling", StringComparison.OrdinalIgnoreCase))
                    {
                        RunQuickSpellingTypography(Path.Combine(directory, "quick-spelling.docx"),
                            Word.WdSaveFormat.wdFormatXMLDocument);
                        Console.WriteLine("QUICK_SPELLING_TYPOGRAPHY_PASS");
                        return 0;
                    }
                    RunExistingDocument(args[0], directory);
                    Console.WriteLine("ONE_CLICK_EXISTING_DOCUMENT_PASS");
                    return 0;
                }
                Run(Path.Combine(directory, "decision.docx"), Word.WdSaveFormat.wdFormatXMLDocument);
                Run(Path.Combine(directory, "decision.doc"), Word.WdSaveFormat.wdFormatDocument97);
                RunMultipleEmbeddedDocuments(Path.Combine(directory, "multiple-documents.docx"));
                RunDashedMottoLineDetection(Path.Combine(directory, "dashed-motto-line.docx"));
                RunQuickSpellingTypography(Path.Combine(directory, "quick-spelling.docx"),
                    Word.WdSaveFormat.wdFormatXMLDocument);
                RunQuickSpellingTypography(Path.Combine(directory, "quick-spelling.doc"),
                    Word.WdSaveFormat.wdFormatDocument97);
                RunUnsavedDocument();
                Console.WriteLine("ONE_CLICK_WORD_SMOKE_PASS DOC DOCX UNSAVED MULTIPLE_DOCUMENTS DASHED_MOTTO_LINE QUICK_SPELLING_TYPOGRAPHY");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("ONE_CLICK_WORD_SMOKE_FAIL " + exception);
                return 1;
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        private static void RunUnsavedDocument()
        {
            Word.Application application = null;
            Word.Document document = null;
            string backupPath = null;
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                document = application.Documents.Add();
                document.Content.Text =
                    "CƠ QUAN BAN HÀNH\r" +
                    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r" +
                    "Độc lập - Tự do - Hạnh phúc\r" +
                    "Số: 01/QĐ-CQ\r" +
                    "Hà Nội, ngày 05 tháng 09 năm 2026\r" +
                    "QUYẾT ĐỊNH\r" +
                    "Về việc kiểm tra văn bản mới chưa lưu\r" +
                    "Căn cứ quy định hiện hành.\r" +
                    "Điều 1. Tổ chức thực hiện.\r";
                document.Content.Font.Name = "Arial";
                document.Content.Font.Size = 10f;
                var originalName = document.Name;
                Assert(string.IsNullOrWhiteSpace(document.Path),
                    "The unsaved smoke fixture unexpectedly has a persistent path.");

                var capability = new WordDocumentCapabilityProvider(application).Evaluate(document);
                Assert(capability.CanReadDocument && !capability.IsSaved,
                    "A new unsaved Word document was not reported as usable.");

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
                    reader.PrepareForOneClick(context, document);
                    context.RequireFullAnalysis();
                    var result = new WordOneClickRuntime(application, access)
                        .Execute(context, document);
                    backupPath = result.BackupPath;

                    Assert(File.Exists(backupPath),
                        "1-Click did not create a recovery copy for the unsaved document.");
                    Assert(string.Equals(Path.GetExtension(backupPath), ".docx",
                            StringComparison.OrdinalIgnoreCase),
                        "A new unsaved document did not receive a DOCX recovery copy.");
                    Assert(string.IsNullOrWhiteSpace(document.Path) &&
                           string.Equals(document.Name, originalName, StringComparison.Ordinal),
                        "1-Click forced Save As or replaced the user's unsaved document.");

                    var commandRuntime = new WordLocalCommandRuntime(application, access);
                    Assert(string.IsNullOrWhiteSpace(commandRuntime.FormatPage()),
                        "A local command unexpectedly created a recovery copy.");
                    Assert(string.IsNullOrWhiteSpace(document.Path),
                        "A local command forced the unsaved document to disk.");

                    reader.Prepare(context, DocumentAnalysisScope.Spelling, document, false);
                    new WordLocalScanRuntime(application, access)
                        .ScanAndAnnotate(context, true, document);
                    Assert(string.IsNullOrWhiteSpace(document.Path),
                        "A scan forced the unsaved document to disk.");
                }
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(backupPath))
                    try { if (File.Exists(backupPath)) File.Delete(backupPath); } catch { }
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

        private static void RunQuickSpellingTypography(string path, Word.WdSaveFormat format)
        {
            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Add();
                document.Content.Text =
                    "Văn bản  có khoảng trắng ,dấu câu;liền ( nội dung ) [ thử ] và \" trích dẫn \".\r";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 13f;
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
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    var reader = new WordDocumentReadRuntime(application, access);
                    reader.Prepare(context, DocumentAnalysisScope.Spelling, document, false);
                    var runtime = new WordOneClickRuntime(application, access);
                    var fixedCount = runtime.FixAllSpellingFindings(context, document);
                    Assert(fixedCount > 0,
                        "Quick spelling did not report its merged typography cleanup.");
                }

                var actual = document.Content.Text ?? string.Empty;
                const string expected =
                    "Văn bản có khoảng trắng, dấu câu; liền (nội dung) [thử] và “trích dẫn”.";
                Assert(string.Equals(actual.TrimEnd('\r', '\a'), expected, StringComparison.Ordinal),
                    "Quick spelling did not merge whitespace, punctuation and bracket/quotation cleanup. " +
                    "Expected=" + expected + " Actual=" + actual.Replace("\r", "|"));
            }
            finally
            {
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; application.Quit(ref save); }
                Release(document);
                Release(application);
            }
        }

        private static void RunDashedMottoLineDetection(string path)
        {
            Word.Application application = null;
            Word.Document document = null;
            Word.Paragraph mottoParagraph = null;
            Word.Range mottoRange = null;
            Word.Shape dashedLine = null;
            string backupPath = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Add();
                document.Content.Text =
                    "CƠ QUAN BAN HÀNH\r" +
                    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r" +
                    "Độc lập - Tự do - Hạnh phúc\r" +
                    "--------------\r" +
                    "Số: 01/QĐ-CQ\r" +
                    "Hà Nội, ngày 04 tháng 09 năm 2026\r" +
                    "QUYẾT ĐỊNH\r" +
                    "Về việc kiểm tra Line Shape\r" +
                    "Căn cứ Luật Ban hành văn bản quy phạm pháp luật.\r" +
                    "Điều 1. Tổ chức thực hiện.\r";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 13f;
                object fileName = path;
                object format = Word.WdSaveFormat.wdFormatXMLDocument;
                document.SaveAs(ref fileName, ref format);

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
                    var reader = new WordDocumentReadRuntime(application, access);
                    reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
                    var motto = context.LastLocalSnapshot.Paragraphs.Single(item =>
                        item.Text.IndexOf("Độc lập - Tự do - Hạnh phúc", StringComparison.Ordinal) >= 0);
                    Assert(motto.PageLeftPoints.HasValue && motto.PageTopPoints.HasValue &&
                           motto.TextWidthPoints.HasValue && motto.TextWidthPoints.Value > 0,
                        "The smoke document did not expose rendered motto geometry.");

                    mottoParagraph = document.Paragraphs[motto.Index];
                    mottoRange = mottoParagraph.Range.Duplicate;
                    var left = (float)motto.PageLeftPoints.Value;
                    var top = (float)(motto.PageTopPoints.Value + 20d);
                    var width = (float)motto.TextWidthPoints.Value;
                    object anchor = mottoRange;
                    dashedLine = document.Shapes.AddLine(left, top, left + width, top, ref anchor);
                    dashedLine.Name = "SMOKE_DASHED_MOTTO_LINE";
                    dashedLine.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    dashedLine.RelativeVerticalPosition =
                        Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    dashedLine.Left = left;
                    dashedLine.Top = top;
                    dashedLine.Line.Visible = Office.MsoTriState.msoTrue;
                    dashedLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
                    dashedLine.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone;
                    dashedLine.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone;
                    document.Save();

                    context.ClearReadAnalysis();
                    reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
                    var captured = context.LastLocalSnapshot.LineShapes.Single(item =>
                        string.Equals(item.Name, "SMOKE_DASHED_MOTTO_LINE", StringComparison.Ordinal));
                    Assert(captured.DashStyle == (int)Office.MsoLineDashStyle.msoLineDash,
                        "Word DashStyle was not preserved by the snapshot; actual=" +
                        (captured.DashStyle.HasValue ? captured.DashStyle.Value.ToString() : "null") + ".");
                    var dashedFinding = context.LastFormatScan.Findings.Single(item =>
                        item.RuleCode == "ND30-PL1-M2-K1-TN-LINE");
                    Assert(dashedFinding.CurrentIssue.IndexOf("không phải nét liền", StringComparison.Ordinal) >= 0 ||
                           dashedFinding.CurrentIssue.IndexOf("chuỗi dấu gạch/chấm", StringComparison.Ordinal) >= 0,
                        "The dashed motto separator was not reported as an invalid line style/substitute: " +
                        dashedFinding.CurrentIssue);

                    var runtime = new WordOneClickRuntime(application, access);
                    var result = runtime.Execute(context, document);
                    backupPath = result.BackupPath;
                    context.ClearReadAnalysis();
                    reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
                    Assert(!context.LastFormatScan.Findings.Any(item =>
                            item.RuleCode == "ND30-PL1-M2-K1-TN-LINE"),
                        "1-Click did not replace the dashed motto line with a valid solid Line Shape.");
                    var postText = document.Content.Text ?? string.Empty;
                    Assert(!postText.Contains("--------------"),
                        "1-Click retained a textual dash run under the motto: " +
                        postText.Replace("\r", "|").Replace("\a", "#"));
                    var normalized = context.LastLocalSnapshot.LineShapes.Single(item =>
                        item.Name.StartsWith("CHUANHOA2_MOTTO_", StringComparison.Ordinal));
                    Assert(normalized.DashStyle == (int)Office.MsoLineDashStyle.msoLineSolid,
                        "The normalized motto Line Shape is not solid.");
                }
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(backupPath)) try { if (File.Exists(backupPath)) File.Delete(backupPath); } catch { }
                Release(dashedLine);
                Release(mottoRange);
                Release(mottoParagraph);
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; application.Quit(ref save); }
                Release(document);
                Release(application);
            }
        }

        private static void RunMultipleEmbeddedDocuments(string path)
        {
            Word.Application application = null;
            Word.Document document = null;
            string firstBackup = null;
            string secondBackup = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Add();
                document.Content.Text =
                    "ỦY BAN NHÂN DÂN HUYỆN MẪU\r" +
                    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r" +
                    "Độc Lập – Tự do – Hạnh phúc\r" +
                    "Số: 01/QĐ-UBND\r" +
                    "Huyện Mẫu, ngày 01 tháng 09 năm 2026\r" +
                    "QUYẾT ĐỊNH\r" +
                    "Về việc phê duyệt kế hoạch\r" +
                    "Căn cứ Luật Đấu thầu số 22/2023/QH15;\r" +
                    "Theo đề nghị của cơ quan chuyên môn.\r" +
                    "Điều 1. Phê duyệt kế hoạch.\r" +
                    "Nơi nhận:\r- Như trên;\r- Lưu: VT.\r\r" +
                    "ỦY BAN NHÂN DÂN HUYỆN MẪU\r" +
                    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r" +
                    "Độc Lập – Tự do – Hạnh phúc\r" +
                    "Số: 129/2026/QĐ-TTĐ.BMC\r" +
                    "Huyện Mẫu, ngày 02 tháng 09 năm 2026\r" +
                    "THÔNG BÁO\r" +
                    "Về việc triển khai nhiệm vụ\r" +
                    "Căn cứ Nghị định số 30/2020/NĐ-CP;\r" +
                    "Theo đề nghị của cơ quan chuyên môn.\r" +
                    "Nội dung thông báo.\r";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 13f;
                object fileName = path;
                object format = Word.WdSaveFormat.wdFormatXMLDocument;
                document.SaveAs(ref fileName, ref format);
                document.Save();

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
                    reader.PrepareForOneClick(context, document);
                    var initialBlocks = new DocumentRoleDetector().DetectBlocks(context.LastLocalSnapshot);
                    Assert(initialBlocks.Count == 2,
                        "The Word snapshot was not separated into two logical documents.");
                    Assert(initialBlocks[0].DocumentTypeCode == LocalDocumentTypeCodes.Decision &&
                        initialBlocks[1].DocumentTypeCode == LocalDocumentTypeCodes.Notice,
                        "Each logical Word document did not retain its own type.");
                    Assert(context.LastFormatScan.Findings.Any(item =>
                        item.RuleCode == "ND30-PL1-M2-K3-ABBR" &&
                        item.Anchor.ParagraphIndex.HasValue &&
                        initialBlocks[1].ContainsParagraph(item.Anchor.ParagraphIndex.Value)),
                        "The second document's mismatched type abbreviation was not detected.");

                    var runtime = new WordOneClickRuntime(application, access);
                    var first = runtime.Execute(context, document);
                    firstBackup = first.BackupPath;
                    Console.WriteLine("MULTI_RESULT changed=" + first.ChangedParagraphs +
                        " lines=" + first.InsertedLines + " spelling=" + first.CorrectedSpellingItems);
                    reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
                    var postBlocks = new DocumentRoleDetector().DetectBlocks(context.LastLocalSnapshot);
                    Assert(postBlocks.Count == 2,
                        "1-Click lost a logical document boundary.");
                    var secondCode = context.LastLocalSnapshot.Paragraphs.Single(item =>
                        postBlocks[1].Roles.TryGetValue(item.Index, out var role) && role == "codeNumber");
                    var actualSecondCode = secondCode.Text.TrimEnd('\r', '\a');
                    Assert(actualSecondCode == "Số: 129/TB-TTĐ-BMC",
                        "1-Click did not normalize the second document's number/type notation: " +
                        actualSecondCode);
                    Assert(context.LastLocalSnapshot.Paragraphs.Count(item =>
                        postBlocks.Any(block => block.Roles.TryGetValue(item.Index, out var role) &&
                            role == "nationalMotto") &&
                        item.Text.TrimEnd('\r', '\a') == LocalAdministrativeTextNormalizer.NationalMotto) == 2,
                        "1-Click did not normalize every embedded document motto.");
                    var remainingCorrectable = context.LastFormatScan.Findings.Where(item =>
                        item.RuleCode == "ND30-PL1-M2-K1-TN-SEP" ||
                        item.RuleCode == "ND30-PL1-M2-K3-ABBR" ||
                        item.RuleCode == "ND30-PL1-M2-K3-SEP" ||
                        item.RuleCode == "ND30-PL1-M2-K1-TN-LINE")
                        .Select(item => item.RuleCode + "@" +
                            item.Anchor.ParagraphIndex.GetValueOrDefault().ToString())
                        .ToArray();
                    Assert(remainingCorrectable.Length == 0,
                        "Correctable findings remained after multi-document 1-Click: " +
                        string.Join(",", remainingCorrectable) + " | lines=" +
                        string.Join(";", context.LastLocalSnapshot.LineShapes.Select(line =>
                            line.Name + "@P" + line.AnchorParagraphIndex.GetValueOrDefault() +
                            "/page" + line.AnchorPageNumber + "/left" +
                            line.PageLeftPoints.GetValueOrDefault().ToString("0.0") +
                            "/top" + line.PageTopPoints.GetValueOrDefault().ToString("0.0") +
                            "/width" + line.WidthPoints.ToString("0.0"))));

                    var ownedLinesAfterFirst = CountOwnedLines(document);
                    reader.PrepareForOneClick(context, document);
                    var second = runtime.Execute(context, document);
                    secondBackup = second.BackupPath;
                    Assert(CountOwnedLines(document) == ownedLinesAfterFirst,
                        "A second 1-Click duplicated managed Line Shapes.");
                }
            }
            finally
            {
                foreach (var backup in new[] { firstBackup, secondBackup })
                    if (!string.IsNullOrWhiteSpace(backup)) try { if (File.Exists(backup)) File.Delete(backup); } catch { }
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; application.Quit(ref save); }
                Release(document);
                Release(application);
            }
        }

        private static int CountOwnedLines(Word.Document document)
        {
            var count = 0;
            for (var index = 1; index <= document.Shapes.Count; index++)
            {
                Word.Shape shape = null;
                try
                {
                    shape = document.Shapes[index];
                    if ((shape.Name ?? string.Empty).StartsWith("CHUANHOA2_", StringComparison.Ordinal)) count++;
                }
                finally { Release(shape); }
            }
            return count;
        }

        private static void RunExistingDocument(string sourcePath, string directory)
        {
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("Smoke source document was not found.", sourcePath);
            var workingPath = Path.Combine(directory, Path.GetFileName(sourcePath));
            File.Copy(sourcePath, workingPath, true);
            Word.Application application = null;
            Word.Document document = null;
            string backupPath = null;
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
                    new WordDocumentReadRuntime(application, access).Read(context, document);
                    var result = new WordOneClickRuntime(application, access).Execute(context, document);
                    backupPath = result.BackupPath;
                    Assert(File.Exists(backupPath), "1-Click did not create its recovery backup.");
                    Console.WriteLine("ONE_CLICK_EXISTING_RESULT changed_paragraphs=" + result.ChangedParagraphs +
                        " corrected_spelling=" + result.CorrectedSpellingItems +
                        " normalized_tables=" + result.NormalizedTables +
                        " inserted_lines=" + result.InsertedLines);
                }
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(backupPath)) try { if (File.Exists(backupPath)) File.Delete(backupPath); } catch { }
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; application.Quit(ref save); }
                Release(document);
                Release(application);
            }
        }

        private static void Run(string path, Word.WdSaveFormat format)
        {
            Word.Application application = null;
            Word.Document document = null;
            Word.Table identityTable = null;
            Word.Table dataTable = null;
            Word.Range insertion = null;
            Word.Range emphasized = null;
            string backupPath = null;
            string selectedFixBackupPath = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Add();
                document.Content.Text = string.Empty;
                identityTable = document.Tables.Add(document.Range(0, 0), 3, 2);
                identityTable.Borders.Enable = 0;
                identityTable.Cell(1, 1).Range.Text = "ỦY BAN NHÂN DÂN HUYỆN MẪU";
                identityTable.Cell(1, 2).Range.Text = "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                identityTable.Cell(2, 2).Range.Text = "Độc lập - Tự do - Hạnh phúc";
                identityTable.Cell(3, 1).Range.Text = "Số: 01/QĐ-UBND";
                identityTable.Cell(3, 2).Range.Text = "Huyện Mẫu, ngày 02 tháng 9 năm 2026";

                // Insert the body after the identity layout table. Content.End - 1
                // points at the final cell marker and would incorrectly make every
                // body paragraph part of that table in the smoke fixture.
                identityTable.Range.InsertParagraphAfter();
                insertion = document.Range(identityTable.Range.End, identityTable.Range.End);
                insertion.InsertAfter("QUYẾT ĐỊNH\rVề việc phê duyệt báo cáo kinh tế - kỹ thuật\r" +
                    "Căn cứ Nghị định 214/2025/NĐ-CP ngày 05 tháng 3 năm 2025 của Chính phủ;\r" +
                    "Căn cứ hồ sơ và đề nghị của cơ quan chuyên môn.\r" +
                    "Điều 1. Nội dung điều này không được in đậm toàn bộ.\r" +
                    "2. NỘI DUNG THẨM ĐỊNH\r" +
                    "b) Ý kiến thẩm định về cơ sở pháp lý:\r" +
                    "Căn cứ các tài liệu được cung cấp, kết quả thẩm định được tổng hợp tại Bảng số 01.\r" +
                    "nội dung quyết định này  có cụm từ cần giữ in đậm , có cách viết sát nhập và hạn 05/03/2020. " +
                    "Quyết định xố, ự án cần kiểm tra.\r" +
                    "-\tNguồn vốn: Nguồn vốn của các công đoàn và đơn vị thành viên được trình bày trên nhiều dòng để kiểm tra thụt lề treo.\r");
                Release(insertion); insertion = null;
                ConfigureLegacyDashListTabs(document);

                var bodyText = document.Content.Text;
                var emphasizedStart = bodyText.IndexOf("cần giữ in đậm", StringComparison.Ordinal);
                emphasized = document.Range(emphasizedStart, emphasizedStart + "cần giữ in đậm".Length);
                emphasized.Font.Bold = -1;
                Release(emphasized); emphasized = null;
                var pointParagraph = FindParagraph(document, "b)");
                pointParagraph.Range.Font.Bold = -1;
                pointParagraph.Range.Font.Italic = -1;
                Release(pointParagraph);

                insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                dataTable = document.Tables.Add(insertion, 2, 2);
                dataTable.Cell(1, 1).Range.Text = "Nội dung";
                dataTable.Cell(1, 2).Range.Text = "Kết quả";
                dataTable.Cell(1, 1).Range.Font.Bold = -1;
                dataTable.Cell(1, 2).Range.Font.Bold = -1;
                dataTable.Cell(2, 1).Range.Text = "Mục kiểm tra";
                dataTable.Cell(2, 2).Range.Text = "Đạt";

                document.Content.Font.Name = "Arial";
                document.Content.Font.Size = 10f;
                emphasized = document.Range(document.Content.Text.IndexOf("cần giữ in đậm", StringComparison.Ordinal),
                    document.Content.Text.IndexOf("cần giữ in đậm", StringComparison.Ordinal) + "cần giữ in đậm".Length);
                Assert(emphasized.Font.Bold != 0, "Smoke setup lost intentional inline bold before saving.");
                Release(emphasized); emphasized = null;
                document.Sections[1].PageSetup.LeftMargin = 10f;
                object fileName = path;
                object saveFormat = format;
                document.SaveAs(ref fileName, ref saveFormat);
                document.Save();
                Assert(document.Saved, "Smoke document was not saved before 1-Click.");
                emphasized = document.Range(document.Content.Text.IndexOf("cần giữ in đậm", StringComparison.Ordinal),
                    document.Content.Text.IndexOf("cần giữ in đậm", StringComparison.Ordinal) + "cần giữ in đậm".Length);
                Assert(emphasized.Font.Bold != 0, "Smoke setup lost intentional inline bold while saving.");
                Release(emphasized); emphasized = null;

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
                    var runtime = new WordOneClickRuntime(application, access);
                    var readRuntime = new WordDocumentReadRuntime(application, access);
                    readRuntime.Read(context, document);
                    Assert(!context.LastFormatScan.Findings.Any(item =>
                            item.RuleCode.StartsWith("ND30-PL1-M2-K6A", StringComparison.Ordinal) &&
                            item.Anchor.ExpectedText != null &&
                            item.Anchor.ExpectedText.IndexOf("tổng hợp tại Bảng", StringComparison.OrdinalIgnoreCase) >= 0),
                        "A body sentence beginning with Căn cứ was misclassified as the formal legal-basis component.");
                    Assert(context.LastFormatScan.Findings.Any(item =>
                            item.RuleCode == "ND30-PL1-M2-K6D-POINT" &&
                            item.Anchor.ParagraphIndex.HasValue &&
                            context.LastLocalSnapshot.Paragraphs.Any(paragraph =>
                                paragraph.Index == item.Anchor.ParagraphIndex.Value &&
                                paragraph.Text.StartsWith("b)", StringComparison.Ordinal))),
                        "The deliberately bold/italic legal point was not reported before 1-Click.");
                    var spellingScan = new WordLocalScanRuntime(application, access)
                        .ScanAndAnnotate(context, true, document);
                    var spellingCommentCountBeforeFix = Enumerable.Range(1, document.Variables.Count)
                        .Count(index => (document.Variables[index].Name ?? string.Empty)
                            .StartsWith("CHCOM_SPELLING_", StringComparison.Ordinal));
                    Assert(spellingCommentCountBeforeFix > 1,
                        "The selected-finding fixture needs multiple spelling comments.");
                    Assert(spellingScan.Findings.Any(item => item.RuleCode == "LOCAL-TYPO-DICT" &&
                        item.Anchor.ExpectedText == "sát nhập"),
                        "The selected-finding fixture was not annotated.");
                    var selectedTypo = document.Content.Duplicate;
                    selectedTypo.Find.ClearFormatting();
                    Assert(selectedTypo.Find.Execute(FindText: "sát nhập", Wrap: Word.WdFindWrap.wdFindStop),
                        "The selected typo could not be found through Word Range.Find.");
                    selectedTypo.Select();
                    Release(selectedTypo);
                    var selectedFix = runtime.ExecuteSelectedFinding(context, document);
                    selectedFixBackupPath = selectedFix.BackupPath;
                    Assert(string.IsNullOrWhiteSpace(selectedFixBackupPath),
                        "Sửa lỗi đang chọn unexpectedly created a backup.");
                    Assert(selectedFix.Resolved && document.Content.Text.Contains("sáp nhập"),
                        "Sửa lỗi đang chọn did not fix and clear the selected dictionary finding.");
                    var spellingCommentCountAfterFix = Enumerable.Range(1, document.Variables.Count)
                        .Count(index => (document.Variables[index].Name ?? string.Empty)
                            .StartsWith("CHCOM_SPELLING_", StringComparison.Ordinal));
                    Assert(spellingCommentCountAfterFix == spellingCommentCountBeforeFix - 1,
                        "Sửa lỗi đang chọn removed unrelated spelling comments.");
                    Assert(!document.Saved,
                        "The selected fix must leave an unsaved document for the save-invalidation regression.");
                    readRuntime.Prepare(context, DocumentAnalysisScope.Full, document, false);
                    var missingSo = context.LastFormatScan.Findings.Single(item =>
                        item.RuleCode == "ND30-PL1-M2-K6B-SO");
                    var annotationSnapshot = new AnnotationDocumentSnapshot(
                        context.LastSnapshot.DocumentFingerprint,
                        context.LastSnapshot.Revision,
                        context.LastSnapshot.Paragraphs.Select(item =>
                            new AnnotationParagraphSnapshot(item.StoryType, item.Index,
                                item.SectionIndex, item.AbsoluteStart, item.Text,
                                item.TableIndex, item.RowIndex, item.CellIndex)).ToArray(),
                        context.LastSnapshot.ProtectedSpans.Select(item =>
                            new AnnotationProtectedSpan(item.StoryType, item.AbsoluteStart,
                                item.Length)).ToArray());
                    var missingSoPlan = new AnnotationPlanner().CreatePlan(
                        context.LastFormatScan.Lane,
                        context.LastFormatScan.ScanId,
                        context.LastFormatScan.DocumentFingerprint,
                        context.LastFormatScan.Revision,
                        annotationSnapshot,
                        new[] { missingSo });
                    Assert(missingSoPlan.Unresolved.Count == 0,
                        "The missing-so regression comment could not be anchored.");
                    new WordFindingAnnotationAdapter(application, document).Apply(missingSoPlan);
                    var missingSoParagraph = context.LastLocalSnapshot.Paragraphs.Single(item =>
                        item.Index == missingSo.Anchor.ParagraphIndex.Value);
                    var missingSoStart = missingSoParagraph.AbsoluteStart +
                        missingSo.Anchor.StartOffset.GetValueOrDefault();
                    var missingSoEnd = missingSoStart + missingSo.Anchor.Length.GetValueOrDefault();
                    var missingSoSelection = document.Range(missingSoStart, missingSoEnd);
                    missingSoSelection.Select();
                    Release(missingSoSelection);
                    var focusAdapter = new WordFindingAnnotationAdapter(application, document);
                    Assert(focusAdapter.TryFocusDocumentSelection(true),
                        "Modern Comments focus could not return to the document.");
                    Assert(application.Selection.Range.Start == application.Selection.Range.End,
                        "Modern Comments focus left the finding text selected and vulnerable to replacement.");
                    var selectedFormatFix = runtime.ExecuteSelectedFinding(
                        context, "format", missingSo.FindingId, missingSoParagraph.StoryType,
                        missingSoStart, missingSoEnd, document);
                    var citedTextAfterSelectedFix = document.Content.Text;
                    var citedTextIndex = citedTextAfterSelectedFix.IndexOf("Nghị định", StringComparison.Ordinal);
                    Console.WriteLine("SELECTED_MISSING_SO_RESULT=" +
                        (citedTextIndex < 0 ? "<not-found>" : citedTextAfterSelectedFix.Substring(
                            citedTextIndex, Math.Min(45, citedTextAfterSelectedFix.Length - citedTextIndex))
                            .Replace("\r", "<CR>")));
                    Assert(selectedFormatFix.Resolved &&
                        document.Content.Text.Contains("Nghị định số 214/2025/NĐ-CP"),
                        "Sửa lỗi đang chọn removed the missing-so comment without inserting 'số'.");
                    var repairedCitationText = "Nghị định số 214/2025/NĐ-CP";
                    var repairedCitationStart = document.Content.Text.IndexOf(
                        repairedCitationText, StringComparison.Ordinal);
                    var repairedCitationRange = document.Range(repairedCitationStart,
                        repairedCitationStart + repairedCitationText.Length);
                    Assert((int)repairedCitationRange.Font.Color != AnnotationOwnershipPolicy.WordRedColor,
                        "Sửa lỗi đang chọn left add-in red formatting after a longer replacement.");
                    Release(repairedCitationRange);
                    var sourceSaveInvalidations = 0;
                    Word.ApplicationEvents4_DocumentBeforeSaveEventHandler beforeSave =
                        delegate(Word.Document savingDocument, ref bool saveAsUi, ref bool cancel)
                        {
                            if (!SameComObject(savingDocument, document)) return;
                            sourceSaveInvalidations++;
                            context.ClearReadAnalysis();
                        };
                    application.DocumentBeforeSave += beforeSave;
                    OneClickResult result;
                    try
                    {
                        // Regression: the Ribbon receives an unsaved document after a
                        // scan/selected fix. The save event clears analysis, so 1-Click
                        // must persist first and only then build its full snapshot.
                        readRuntime.PrepareForOneClick(context, document);
                        Assert(sourceSaveInvalidations == 1,
                            "The regression setup did not invalidate analysis at the source save boundary.");
                        context.RequireFullAnalysis();
                        result = runtime.Execute(context, document);
                        readRuntime.Prepare(context, DocumentAnalysisScope.Full, document, false);
                        Assert(!context.LastFormatScan.Findings.Any(item =>
                                item.RuleCode.EndsWith("-LINE", StringComparison.Ordinal)),
                            "A required Line Shape remained missing or offset after 1-Click.");
                    }
                    finally
                    {
                        application.DocumentBeforeSave -= beforeSave;
                    }
                    Assert(context.DocumentTypeCode == LocalDocumentTypeCodes.Decision &&
                        !context.DocumentTypeWasSelectedManually,
                        "1-Click did not automatically detect the Decision document type.");
                    backupPath = result.BackupPath;
                    Assert(File.Exists(backupPath), "Recovery backup was not created.");
                    var expectedBackupDirectory = Path.GetFullPath(
                        Path.Combine(Path.GetTempPath(), "ChuanHoa", "Backups"));
                    Assert(string.Equals(Path.GetDirectoryName(Path.GetFullPath(backupPath)),
                            expectedBackupDirectory, StringComparison.OrdinalIgnoreCase),
                        "1-Click backup was not stored under Windows Temp.");
                    Assert(string.Equals(Path.GetExtension(backupPath), Path.GetExtension(path), StringComparison.OrdinalIgnoreCase),
                        "Recovery backup did not preserve the source format.");
                    Assert(result.InsertedLines >= 3, "Required ND30 Line Shapes were not inserted.");
                    Console.WriteLine("CORRECTED_SPELLING_ITEMS=" + result.CorrectedSpellingItems);
                    Assert(result.CorrectedSpellingItems >= 6,
                        "1-Click did not apply deterministic spelling fixes; actual=" +
                        result.CorrectedSpellingItems + ".");
                    Assert(result.NormalizedTables == 1, "Only the real data table should be normalized; actual=" +
                        result.NormalizedTables + ".");
                }

                var normalized = document.Content.Text ?? string.Empty;
                Assert(normalized.Contains("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM"), "HOÀ was rewritten or national title changed.");
                Assert(normalized.Contains("VIỆT NAM"), "VIỆT NAM did not remain uppercase.");
                Assert(!normalized.Contains("Việt Nam"), "National title was incorrectly title-cased.");
                Assert(normalized.Contains("Nội dung quyết định này có cụm từ cần giữ in đậm, có cách viết sáp nhập và hạn ngày 05/03/2020."),
                    "1-Click did not fix capitalization, spacing, dictionary spelling and the missing ngày prefix.");
                Assert(normalized.Contains("Quyết định số, dự án cần kiểm tra."),
                    "1-Click did not fix contextual Vietnamese spelling mistakes.");
                Assert(normalized.Contains("Nghị định số 214/2025/NĐ-CP"),
                    "1-Click removed the missing-so finding without inserting the required word.");
                var subject = FindParagraph(document, "Về việc phê duyệt");
                Assert(subject.Range.Bold != 0, "Decision subject was not bold after 1-Click.");
                Release(subject);
                var article = FindParagraph(document, "Điều 1.");
                var articleText = article.Range.Text.TrimEnd('\r', '\a');
                var articleMarkerLength = articleText.IndexOf('.') + 1;
                var articleMarker = document.Range(article.Range.Start,
                    article.Range.Start + articleMarkerLength);
                var articleContent = document.Range(article.Range.Start + articleMarkerLength,
                    article.Range.Start + articleText.Length);
                Assert(articleMarker.Font.Bold != 0 && articleContent.Font.Bold == 0,
                    "1-Click bolded the whole Điều paragraph instead of only its marker.");
                Assert(Math.Abs(article.Range.ParagraphFormat.FirstLineIndent - 10f * 72f / 25.4f) < .6f,
                    "1-Click did not apply the ND30 10 mm article indent.");
                Release(articleContent);
                Release(articleMarker);
                Release(article);
                emphasized = document.Range(normalized.IndexOf("cần giữ in đậm", StringComparison.Ordinal),
                    normalized.IndexOf("cần giữ in đậm", StringComparison.Ordinal) + "cần giữ in đậm".Length);
                Assert(emphasized.Font.Bold != 0, "Intentional inline bold formatting was removed from body text.");
                pointParagraph = FindParagraph(document, "b)");
                Assert(pointParagraph.Range.Font.Bold == 0 && pointParagraph.Range.Font.Italic == 0,
                    "1-Click did not normalize the legal point to upright, non-bold text.");
                Release(pointParagraph);
                AssertNormalizedDashListIndent(document);
                Assert(Enumerable.Range(1, document.Shapes.Count)
                    .Select(index => document.Shapes[index].Name)
                    .Count(name => name.StartsWith("CHUANHOA2_", StringComparison.Ordinal)) >= 3,
                    "Owned Line Shape markers are missing.");
                AssertRenderedOwnedLine(document, "CHUANHOA2_ORG_", 28f, 95f);
                AssertRenderedOwnedLine(document, "CHUANHOA2_MOTTO_", 120f, 190f);
                Assert(!Enumerable.Range(1, document.Shapes.Count)
                        .Select(index => document.Shapes[index].Name)
                        .Any(name => name.StartsWith("CHUANHOA_", StringComparison.Ordinal)),
                    "A legacy owned line remained and can disappear after Word serializes the document.");
                Assert(dataTable.Rows[1].HeadingFormat != 0, "The real data table header was not repeated.");
                Assert(identityTable.Rows[1].HeadingFormat == 0, "The identity layout table was mistaken for a data table.");
                Assert(document.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Count > 0,
                    "Page-number field was not inserted.");
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(backupPath)) try { if (File.Exists(backupPath)) File.Delete(backupPath); } catch { }
                if (!string.IsNullOrWhiteSpace(selectedFixBackupPath)) try { if (File.Exists(selectedFixBackupPath)) File.Delete(selectedFixBackupPath); } catch { }
                Release(emphasized);
                Release(insertion);
                Release(dataTable);
                Release(identityTable);
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; application.Quit(ref save); }
                Release(document);
                Release(application);
            }
        }

        private static Word.Paragraph FindParagraph(Word.Document document, string prefix)
        {
            foreach (Word.Paragraph paragraph in document.Paragraphs)
            {
                if ((paragraph.Range.Text ?? string.Empty).Trim().StartsWith(prefix, StringComparison.Ordinal)) return paragraph;
                Release(paragraph);
            }
            throw new InvalidOperationException("Paragraph not found: " + prefix);
        }

        private static void ConfigureLegacyDashListTabs(Word.Document document)
        {
            const float pointsPerMillimeter = 72f / 25.4f;
            Word.Paragraph paragraph = null;
            Word.Range range = null;
            Word.ParagraphFormat format = null;
            Word.TabStops tabStops = null;
            try
            {
                paragraph = FindParagraph(document, "-");
                range = paragraph.Range.Duplicate;
                format = range.ParagraphFormat;
                format.LeftIndent = 0f;
                format.FirstLineIndent = 10f * pointsPerMillimeter;
                tabStops = format.TabStops;
                tabStops.ClearAll();
                tabStops.Add(80f * pointsPerMillimeter, Word.WdTabAlignment.wdAlignTabLeft,
                    Word.WdTabLeader.wdTabLeaderSpaces);
            }
            finally
            {
                Release(tabStops);
                Release(format);
                Release(range);
                Release(paragraph);
            }
        }

        private static void AssertNormalizedDashListIndent(Word.Document document)
        {
            const float pointsPerMillimeter = 72f / 25.4f;
            Word.Paragraph paragraph = null;
            Word.Range range = null;
            Word.ParagraphFormat format = null;
            Word.TabStops tabStops = null;
            Word.TabStop tabStop = null;
            try
            {
                paragraph = FindParagraph(document, "-");
                range = paragraph.Range.Duplicate;
                format = range.ParagraphFormat;
                Assert(Math.Abs(format.LeftIndent - 15f * pointsPerMillimeter) < .5f,
                    "Dash-list continuation does not align at 15 mm.");
                Assert(Math.Abs(format.FirstLineIndent + 5f * pointsPerMillimeter) < .5f,
                    "Dash-list marker does not use the expected 5 mm hanging indent.");
                Assert(Math.Abs(format.RightIndent) < .1f,
                    "Dash-list retained a stale right indent.");
                tabStops = format.TabStops;
                var expectedPosition = 15f * pointsPerMillimeter;
                var customPosition = 80f * pointsPerMillimeter;
                var foundExpected = false;
                var foundCustom = false;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    Release(tabStop);
                    tabStop = tabStops[index];
                    if (Math.Abs(tabStop.Position - expectedPosition) < .5f) foundExpected = true;
                    if (Math.Abs(tabStop.Position - customPosition) < .5f) foundCustom = true;
                }
                Assert(foundExpected,
                    "Dash-list tab stop is not aligned with continuation text.");
                Assert(foundCustom,
                    "Dash-list removed the user's pre-existing 80 mm custom tab stop.");
            }
            finally
            {
                Release(tabStop);
                Release(tabStops);
                Release(format);
                Release(range);
                Release(paragraph);
            }
        }

        private static void AssertRenderedOwnedLine(Word.Document document, string namePrefix,
            float minimumWidth, float maximumWidth)
        {
            Word.Shape matched = null;
            try
            {
                for (var index = 1; index <= document.Shapes.Count; index++)
                {
                    Word.Shape candidate = null;
                    try
                    {
                        candidate = document.Shapes[index];
                        if (!(candidate.Name ?? string.Empty).StartsWith(namePrefix, StringComparison.Ordinal))
                            continue;
                        matched = candidate;
                        candidate = null;
                        break;
                    }
                    finally { Release(candidate); }
                }

                Assert(matched != null, "The expected owned Line Shape was not found: " + namePrefix);
                Assert((int)matched.Type == 9, "The owned object is not a Line Shape: " + namePrefix);
                Assert(matched.Line.Visible == Office.MsoTriState.msoTrue,
                    "The owned Line Shape is hidden: " + namePrefix);
                Assert(matched.Line.DashStyle == Office.MsoLineDashStyle.msoLineSolid,
                    "The owned Line Shape is not solid: " + namePrefix);
                Assert(matched.Line.BeginArrowheadStyle == Office.MsoArrowheadStyle.msoArrowheadNone &&
                       matched.Line.EndArrowheadStyle == Office.MsoArrowheadStyle.msoArrowheadNone,
                    "The owned Line Shape unexpectedly has an arrowhead: " + namePrefix);
                Assert(matched.RelativeHorizontalPosition ==
                       Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage &&
                       matched.RelativeVerticalPosition ==
                       Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage,
                    "The owned Line Shape is not positioned relative to the page: " + namePrefix);
                // Word can coerce LayoutInCell back to true for a table-anchored line
                // even though the serialized floating shape uses layoutInCell="0".
                // WrapNone and the page-relative coordinates are the stable COM
                // invariants that prevent displacement across Word 2010+.
                Assert(matched.WrapFormat.Type == Word.WdWrapType.wdWrapNone,
                    "The owned Line Shape can be displaced or hidden by wrapping: " + namePrefix +
                    "; wrapType=" + matched.WrapFormat.Type + ".");
                Assert(Math.Abs(matched.WrapFormat.DistanceTop) < .01f &&
                       Math.Abs(matched.WrapFormat.DistanceBottom) < .01f &&
                       Math.Abs(matched.WrapFormat.DistanceLeft) < .01f &&
                       Math.Abs(matched.WrapFormat.DistanceRight) < .01f,
                    "The owned Line Shape retained an invalid wrap distance: " + namePrefix);
                Assert(matched.Width >= minimumWidth && matched.Width <= maximumWidth &&
                       Math.Abs(matched.Height) < .01f,
                    "The owned Line Shape has implausible rendered geometry: " + namePrefix);
            }
            finally { Release(matched); }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static bool SameComObject(object left, object right)
        {
            if (left == null || right == null) return false;
            var leftIdentity = Marshal.GetIUnknownForObject(left);
            var rightIdentity = Marshal.GetIUnknownForObject(right);
            try { return leftIdentity == rightIdentity; }
            finally
            {
                Marshal.Release(leftIdentity);
                Marshal.Release(rightIdentity);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
