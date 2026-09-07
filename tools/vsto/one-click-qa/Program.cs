#nullable disable
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Scanning;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.OneClickQa
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length != 2 && !(args.Length == 3 && args[2] == "--twice"))
            {
                Console.Error.WriteLine("Usage: ChuanHoa.OneClickQa <source.doc|docx> <qa-output-directory> [--twice]");
                return 2;
            }

            var source = Path.GetFullPath(args[0]);
            var outputDirectory = Path.GetFullPath(args[1]);
            if (!File.Exists(source)) throw new FileNotFoundException("Source document was not found.", source);
            Directory.CreateDirectory(outputDirectory);
            var qaDocument = Path.Combine(outputDirectory, "one-click-result" + Path.GetExtension(source));
            var qaPdf = Path.Combine(outputDirectory, "one-click-result.pdf");
            var sourceHash = Hash(source);
            File.Copy(source, qaDocument, true);

            Word.Application application = null;
            Word.Document document = null;
            string backupPath = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                object fileName = qaDocument;
                object readOnly = false;
                object visible = false;
                document = application.Documents.Open(ref fileName, ReadOnly: ref readOnly, Visible: ref visible);
                document.Activate();
                document.Save();
                var context = new DocumentContext(document.GetHashCode())
                {
                    RegimeCode = "ND30",
                    DocumentTypeCode = LocalDocumentTypeCodes.Unknown,
                    RegimeWasSelectedManually = true,
                    DocumentTypeWasSelectedManually = false
                };
                OneClickResult result;
                LocalScanResult postFormat;
                LocalScanResult postSpelling;
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    var reader = new WordDocumentReadRuntime(application, access);
                    var scanner = new WordLocalScanRuntime(application, access);
                    reader.PrepareForOneClick(context, document);
                    if (context.LastLocalSnapshot.Paragraphs.Count > 8)
                    {
                        var p9 = context.LastLocalSnapshot.Paragraphs[8];
                        Console.WriteLine("P9: [" + p9.Text + "] len=" + p9.Text.Length + " hex=" + BitConverter.ToString(System.Text.Encoding.UTF8.GetBytes(p9.Text)));
                    }
                    Console.WriteLine("DETECTED_TYPE=" + context.DocumentTypeCode);
                    WriteSnapshotDiagnostics(context.LastLocalSnapshot);
                    result = new WordOneClickRuntime(application, access).Execute(context, document);
                    reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
                    VerifyOwnedLineAnchors(context.LastLocalSnapshot);
                    if (args.Length == 3)
                    {
                        var firstNames = context.LastLocalSnapshot.LineShapes.Where(l => LineShapeOwnership.IsOwned(l.Name))
                            .Select(l => l.Name).OrderBy(n => n, StringComparer.Ordinal).ToArray();
                        reader.PrepareForOneClick(context, document);
                        result = new WordOneClickRuntime(application, access).Execute(context, document);
                        reader.Prepare(context, DocumentAnalysisScope.Full, document, false);
                        VerifyOwnedLineAnchors(context.LastLocalSnapshot);
                        var secondNames = context.LastLocalSnapshot.LineShapes.Where(l => LineShapeOwnership.IsOwned(l.Name))
                            .Select(l => l.Name).OrderBy(n => n, StringComparer.Ordinal).ToArray();
                        if (!firstNames.SequenceEqual(secondNames))
                            throw new InvalidOperationException("Repeated normalization changed component-line ownership or count.");
                        Console.WriteLine("REPEATED_LINE_OWNERSHIP_PASS COUNT=" + secondNames.Length);
                    }
                    postFormat = scanner.ScanAndAnnotate(context, false, document);
                    postSpelling = scanner.ScanAndAnnotate(context, true, document);
                }
                backupPath = result.BackupPath;
                document.Save();
                document.ExportAsFixedFormat(qaPdf, Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false, OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: Word.WdExportRange.wdExportAllDocument, Item: Word.WdExportItem.wdExportDocumentContent,
                    IncludeDocProps: false, KeepIRM: false, CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                    DocStructureTags: true, BitmapMissingFonts: true, UseISO19005_1: false);
                if (!string.Equals(sourceHash, Hash(source), StringComparison.Ordinal))
                    throw new InvalidOperationException("The original source document changed during QA.");
                Console.WriteLine("QA_DOCUMENT=" + qaDocument);
                Console.WriteLine("QA_PDF=" + qaPdf);
                Console.WriteLine("CHANGED_PARAGRAPHS=" + result.ChangedParagraphs);
                Console.WriteLine("INSERTED_LINES=" + result.InsertedLines);
                Console.WriteLine("NORMALIZED_SECTIONS=" + result.NormalizedSections);
                Console.WriteLine("NORMALIZED_TABLES=" + result.NormalizedTables);
                Console.WriteLine("REMAINING_FINDINGS=" + result.RemainingFindings);
                Console.WriteLine("POST_FORMAT_FINDINGS=" + postFormat.Findings.Count);
                Console.WriteLine("POST_SPELLING_FINDINGS=" + postSpelling.Findings.Count);
                Console.WriteLine("POST_COMMENTS=" + document.Comments.Count);
                foreach (var finding in result.RemainingFindingItems)
                    Console.WriteLine("FINDING=" + finding.RuleCode + "|" + finding.CurrentIssue + "|" + finding.Expected);
                var remainingLine = postFormat.Findings.FirstOrDefault(item =>
                    item.RuleCode.EndsWith("-LINE", StringComparison.Ordinal));
                if (remainingLine != null)
                    throw new InvalidOperationException(
                        "1-Click left a required Line Shape missing or offset: " +
                        remainingLine.RuleCode + ".");
                Console.WriteLine("ONE_CLICK_QA_PASS");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("ONE_CLICK_QA_FAIL " + exception);
                return 1;
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

        private static string Hash(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", string.Empty);
        }

        private static void VerifyOwnedLineAnchors(LocalScanSnapshot snapshot)
        {
            var blocks = new DocumentRoleDetector().DetectBlocks(snapshot);
            var owned = snapshot.LineShapes.Where(l => LineShapeOwnership.IsOwned(l.Name)).ToArray();
            if (owned.GroupBy(l => l.Name).Any(g => g.Count() > 1))
                throw new InvalidOperationException("Duplicate generated component lines.");
            foreach (var line in owned)
            {
                var owner = snapshot.Paragraphs.FirstOrDefault(p => p.StoryType == line.AnchorStoryType &&
                    LineShapeOwnership.IsOwnedForParagraph(line.Name, p.Index));
                var block = owner == null ? null : blocks.FirstOrDefault(b => b.ContainsParagraph(owner.Index));
                var anchor = snapshot.Paragraphs.FirstOrDefault(p => p.Index == line.AnchorParagraphIndex &&
                    p.StoryType == line.AnchorStoryType);
                if (owner == null || anchor == null || block == null ||
                    !ComponentLineAnchorPolicy.CanAnchor(owner.PageNumber, line.AnchorPageNumber,
                        owner.SectionIndex, line.AnchorSectionIndex, anchor.Index,
                        block.StartParagraphIndex, block.EndParagraphIndex))
                    throw new InvalidOperationException("Generated line escaped its physical page/section/document: " + line.Name);
            }
            Console.WriteLine("OWNED_LINE_ANCHORS_PASS BLOCKS=" + blocks.Count + " LINES=" + owned.Length);
        }

        private static void WriteSnapshotDiagnostics(LocalScanSnapshot snapshot)
        {
            var roles = new DocumentRoleDetector().Detect(snapshot);
            foreach (var paragraph in snapshot.Paragraphs)
            {
                string role;
                roles.TryGetValue(paragraph.Index, out role);
                Console.WriteLine("PARAGRAPH=P" + paragraph.Index + "|role=" + (role ?? string.Empty) +
                    "|table=" + paragraph.TableIndex + "/" + paragraph.RowIndex + "/" + paragraph.CellIndex +
                    "|size=" + paragraph.FontSizePoints + "|bold=" + paragraph.Bold +
                    "|align=" + paragraph.Alignment + "|left=" + paragraph.PageLeftPoints +
                    "|top=" + paragraph.PageTopPoints + "|width=" + paragraph.TextWidthPoints +
                    "|text=" + paragraph.Text.Replace("\r", " ").Replace("\n", " "));
                if (ChuanHoa.Client.Core.Text.CombinedNationalHeader.GetBreakOffsets(paragraph.Text).Length > 0)
                    Console.WriteLine("COMBINED_FLAGS=" + paragraph.Index + ":" + paragraph.HasField + ":" +
                        paragraph.HasContentControl + ":" + paragraph.HasHyperlink + ":" + paragraph.HasMathObject +
                        ":protected=" + snapshot.IntersectsProtectedSpan(paragraph));
            }
            foreach (var line in snapshot.LineShapes)
                Console.WriteLine("LINE=L" + line.Index + "|name=" + line.Name + "|anchor=P" +
                    line.AnchorParagraphIndex + "@" + line.AnchorAbsoluteStart + "|page=" + line.AnchorPageNumber +
                    "|left=" + line.PageLeftPoints + "|top=" + line.PageTopPoints +
                    "|width=" + line.WidthPoints + "|height=" + line.HeightPoints +
                    "|relative=" + line.RelativeHorizontalPosition + "/" + line.RelativeVerticalPosition);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
