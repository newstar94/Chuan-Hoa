#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.LocalCommandSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main()
        {
            var directory = Path.Combine(Path.GetTempPath(), "ChuanHoaLocalCommandSmoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                TestBackupCleanupPolicy(directory);
                Run(Path.Combine(directory, "sample.docx"), Word.WdSaveFormat.wdFormatXMLDocument);
                Run(Path.Combine(directory, "sample.doc"), Word.WdSaveFormat.wdFormatDocument97);
                Console.WriteLine("LOCAL_COMMAND_WORD_SMOKE_PASS DOC DOCX");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("LOCAL_COMMAND_WORD_SMOKE_FAIL " + exception);
                return 1;
            }
            finally { try { Directory.Delete(directory, true); } catch { } }
        }

        private static void Run(string path, Word.WdSaveFormat format)
        {
            Word.Application application = null;
            Word.Document document = null;
            Word.Table table = null;
            Word.Table mergedHeaderTable = null;
            var backups = new List<string>();
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Add();
                document.Content.Text = "Đoạn thứ nhất cần giữ với đoạn sau.\rGiá trị 1,234,567.89\r\r\r";
                var end = document.Range(document.Content.End - 1, document.Content.End - 1);
                table = document.Tables.Add(end, 2, 2);
                table.Cell(1, 1).Range.Text = "Tiêu đề";
                table.Cell(2, 1).Range.Text = "A\vB";
                document.Content.InsertAfter("\r");
                var mergedInsertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                mergedHeaderTable = document.Tables.Add(mergedInsertion, 5, 4);
                Release(mergedInsertion);
                mergedHeaderTable.Cell(1, 1).Range.Text = "STT";
                mergedHeaderTable.Cell(1, 2).Range.Text = "Đánh giá";
                mergedHeaderTable.Cell(1, 3).Range.Text = string.Empty;
                mergedHeaderTable.Cell(1, 4).Range.Text = "Kết luận";
                mergedHeaderTable.Cell(2, 2).Range.Text = "Năng lực";
                mergedHeaderTable.Cell(2, 3).Range.Text = "Kỹ thuật";
                mergedHeaderTable.Cell(1, 1).Merge(mergedHeaderTable.Cell(2, 1));
                mergedHeaderTable.Cell(1, 2).Merge(mergedHeaderTable.Cell(1, 3));
                mergedHeaderTable.Cell(1, 3).Merge(mergedHeaderTable.Cell(2, 4));
                SetVisualRowsBold(mergedHeaderTable, 2);
                mergedHeaderTable.Cell(3, 1).Range.Text = "1";
                mergedHeaderTable.Cell(4, 1).Range.Text = "2";
                mergedHeaderTable.Cell(5, 1).Range.Text = "3";
                object fileName = path;
                object saveFormat = format;
                document.SaveAs(ref fileName, ref saveFormat);

                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    var commands = new WordLocalCommandRuntime(application, access);

                    document.Content.InsertAfter("\r\r\r");
                    var trailingLength = document.Content.End;
                    var deleteTimer = Stopwatch.StartNew();
                    AssertNoBackup(commands.RemoveTrailingBlankParagraphs(), "RemoveTrailingBlankParagraphs");
                    deleteTimer.Stop();
                    Assert(deleteTimer.Elapsed < TimeSpan.FromSeconds(5),
                        "Removing trailing blank paragraphs took too long and may have entered a retry loop.");
                    Assert(document.Content.End < trailingLength,
                        "Trailing blank paragraphs were not removed.");
                    Assert(document.Tables.Count == 2 &&
                        (mergedHeaderTable.Cell(1, 1).Range.Text ?? string.Empty).Contains("STT"),
                        "Removing trailing blank paragraphs changed the final table.");

                    // Word always keeps a final paragraph after a table. Running the
                    // command again must leave that sentinel alone and return promptly.
                    deleteTimer.Restart();
                    AssertNoBackup(commands.RemoveTrailingBlankParagraphs(),
                        "RemoveTrailingBlankParagraphsAfterTable");
                    deleteTimer.Stop();
                    Assert(deleteTimer.Elapsed < TimeSpan.FromSeconds(5),
                        "The required paragraph after the final table caused a retry loop.");

                    document.Content.InsertAfter("Thay đổi chưa lưu.");
                    Assert(!document.Saved, "The smoke document should have an unsaved edit before the command.");
                    AssertNoBackup(commands.FormatPage(), "FormatPage");
                    AssertActiveDocument(application, document, "FormatPage");
                    Assert(new WordDocumentCapabilityProvider(application).Evaluate(document).CanReadDocument,
                        "The captured source document became unavailable after FormatPage.");
                    Assert(!document.Saved, "A local command unexpectedly saved the document.");
                    Assert(Math.Abs(document.Sections[1].PageSetup.PageWidth - 210f * 72f / 25.4f) < 2f, "A4 width was not applied.");
                    document.Save();

                    document.Paragraphs[1].Range.Select();
                    AssertNoBackup(commands.KeepWithNext(), "KeepWithNext");
                    AssertActiveDocument(application, document, "KeepWithNext after FormatPage");
                    Assert(document.Paragraphs[1].Format.KeepWithNext != 0, "KeepWithNext was not applied.");
                    document.Save();

                    AssertNoBackup(commands.InsertPageNumbers(), "InsertPageNumbers");
                    Assert(document.Bookmarks.Exists("CHUANHOA_PAGE_NUMBER_S1"), "Owned page-number marker is missing.");
                    Assert(document.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Count > 0,
                        "Page-number field is missing.");
                    document.Save();

                    var spacingText = document.Content.Text;
                    var spacingStart = spacingText.IndexOf("thứ nhất", StringComparison.Ordinal);
                    var spacingRange = document.Range(spacingStart, spacingStart + "thứ nhất".Length);
                    var unaffectedRange = document.Range(0, spacingStart);
                    spacingRange.Select();
                    AssertNoBackup(commands.SetCharacterSpacing(0.1f, false), "ExpandCharacterSpacing");
                    Assert(Math.Abs(spacingRange.Font.Spacing - 0.1f) < 0.001f,
                        "Character spacing was not applied to the selected text.");
                    Assert(Math.Abs(unaffectedRange.Font.Spacing) < 0.001f,
                        "Character spacing leaked outside the selected text.");
                    document.Range(spacingStart, spacingStart).Select();
                    AssertFailsWithoutSelection(() => commands.SetCharacterSpacing(-0.1f, false));
                    Assert(Math.Abs(spacingRange.Font.Spacing - 0.1f) < 0.001f,
                        "A collapsed selection unexpectedly changed or undid text spacing.");
                    spacingRange.Select();
                    AssertNoBackup(commands.SetCharacterSpacing(0f, true), "ResetCharacterSpacing");
                    Assert(Math.Abs(spacingRange.Font.Spacing) < 0.001f, "Character spacing was not reset.");
                    Release(unaffectedRange);
                    Release(spacingRange);
                    document.Save();

                    AssertNoBackup(commands.RepeatTableHeaders(), "RepeatTableHeaders");
                    Assert(table.Rows[1].HeadingFormat != 0, "Table header was not repeated.");
                    AssertVisualHeaderState(mergedHeaderTable, 2);
                    document.Save();
                    AssertNoBackup(commands.CenterTables(), "CenterTables");
                    Assert(table.Rows.Alignment == Word.WdRowAlignment.wdAlignRowCenter, "Table was not centered.");
                    document.Save();

                    table.Cell(2, 1).Range.Select();
                    AssertNoBackup(commands.AlignCurrentCells(true), "AlignCurrentCells");
                    Assert(table.Cell(2, 1).VerticalAlignment == Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                        "Cell was not vertically centered.");
                    document.Save();
                    AssertNoBackup(commands.CleanExcelTableCharacters(), "CleanExcelTableCharacters");
                    Assert(!(table.Cell(2, 1).Range.Text ?? string.Empty).Contains("\v"), "Excel line-break character remains.");
                    document.Save();

                    var content = document.Content.Text;
                    var numberStart = content.IndexOf("1,234,567.89", StringComparison.Ordinal);
                    document.Range(numberStart, numberStart + "1,234,567.89".Length).Select();
                    AssertNoBackup(commands.ConvertDecimalSeparators(), "ConvertDecimalSeparators");
                    Assert(document.Content.Text.Contains("1.234.567,89"), "Decimal separators were not converted.");
                    document.Save();

                    var legacy = document.Range(document.Content.End - 1, document.Content.End - 1);
                    legacy.Text = "\rVI" + (char)0xD6 + "T NAM";
                    legacy.Font.Name = ".VnTimeH";
                    document.Save();
                    AssertNoBackup(commands.ConvertLegacyEncodingToUnicode(), "ConvertLegacyEncodingToUnicode");
                    Assert(document.Content.Text.Contains("VIỆT NAM"), "TCVN3 uppercase text was not converted to Unicode.");
                    Assert(string.Equals(legacy.Font.Name, "Times New Roman", StringComparison.OrdinalIgnoreCase),
                        "Converted text did not receive Times New Roman.");
                    document.Save();

                    var tone = document.Range(document.Content.End - 1, document.Content.End - 1);
                    tone.Text = "\rhòa thủy khóa";
                    document.Save();
                    var mainVowelBackup = commands.NormalizeTonePlacement(VietnameseTonePlacementStyle.MainVowel);
                    AssertBackup(mainVowelBackup, path, "MainVowelTonePlacement");
                    backups.Add(mainVowelBackup);
                    Assert(document.Content.Text.Contains("hoà thuỷ khoá"),
                        "Main-vowel tone placement was not applied. Content=" + document.Content.Text);
                    document.Save();
                    var firstVowelBackup = commands.NormalizeTonePlacement(VietnameseTonePlacementStyle.FirstVowel);
                    AssertBackup(firstVowelBackup, path, "FirstVowelTonePlacement");
                    backups.Add(firstVowelBackup);
                    Assert(document.Content.Text.Contains("hòa thủy khóa"), "First-vowel tone placement was not restored.");
                    document.Save();

                    var context = new DocumentContext(1) { RegimeCode = "ND30", DocumentTypeCode = "DECISION" };
                    new WordDocumentReadRuntime(application, access).Read(context, document);
                    AssertNoBackup(commands.BuildStyleSet(context, 14f), "BuildStyleSet");
                    Assert(Math.Abs(document.Styles[Word.WdBuiltinStyle.wdStyleNormal].Font.Size - 14f) < .1f,
                        "The 14-point style set was not built.");
                    document.Save();
                    new WordDocumentReadRuntime(application, access).Read(context, document);
                    AssertNoBackup(commands.ApplyFontSizeSet(context, 15f), "ApplyFontSizeSet");
                    Assert(Math.Abs(document.Paragraphs[1].Range.Font.Size - 15f) < .1f,
                        "The 15-point role-aware font-size set was not applied.");
                    document.Save();

                    document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                    var imageCount = document.InlineShapes.Count;
                    AssertNoBackup(commands.InsertQrCode("https://ngoctien.id.vn/chuan-hoa-the-thuc/feedback"), "InsertQrCode");
                    Assert(document.InlineShapes.Count == imageCount + 1, "The QR image was not inserted.");
                    var qr = document.InlineShapes[document.InlineShapes.Count];
                    Assert(Math.Abs(qr.Width - 50f * 72f / 25.4f) < 2f && Math.Abs(qr.Height - 50f * 72f / 25.4f) < 2f,
                        "The QR image is not 5 x 5 cm.");
                    Release(qr);
                    document.Save();

                    document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                    var sectionCount = document.Sections.Count;
                    AssertNoBackup(commands.InsertSection(true), "InsertSection");
                    Assert(document.Sections.Count == sectionCount + 1, "Landscape section was not inserted.");
                    Assert(document.Sections[document.Sections.Count].PageSetup.Orientation == Word.WdOrientation.wdOrientLandscape,
                        "Inserted section is not landscape.");
                }
            }
            finally
            {
                foreach (var backup in backups) try { if (File.Exists(backup)) File.Delete(backup); } catch { }
                if (document != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; document.Close(ref save); }
                if (application != null) { object save = Word.WdSaveOptions.wdDoNotSaveChanges; application.Quit(ref save); }
                Release(mergedHeaderTable); Release(table); Release(document); Release(application);
            }
        }

        private static void SetVisualRowsBold(Word.Table table, int headerRowCount)
        {
            foreach (Word.Cell cell in table.Range.Cells)
            {
                try
                {
                    if (cell.RowIndex <= headerRowCount) cell.Range.Font.Bold = -1;
                }
                finally { Release(cell); }
            }
        }

        private static void AssertVisualHeaderState(Word.Table table, int headerRowCount)
        {
            foreach (Word.Cell cell in table.Range.Cells)
            {
                try
                {
                    var heading = cell.Range.Rows.HeadingFormat != 0;
                    if (cell.RowIndex <= headerRowCount)
                        Assert(heading, "Merged visual header row " + cell.RowIndex + " was not repeated.");
                    else
                        Assert(!heading, "A data row was incorrectly marked as a repeating header.");
                }
                finally { Release(cell); }
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void AssertNoBackup(string result, string operation)
        {
            Assert(string.IsNullOrWhiteSpace(result), operation + " unexpectedly created a backup.");
        }

        private static void TestBackupCleanupPolicy(string parentDirectory)
        {
            var directory = Path.Combine(parentDirectory, "cleanup-policy");
            Directory.CreateDirectory(directory);
            var nonOwnedFile = Path.Combine(directory, "keep-user-file.docx");
            File.WriteAllText(nonOwnedFile, "not owned by Chuẩn hóa");
            string oldOwnedFile = null;
            for (var index = 0; index < 22; index++)
            {
                var file = Path.Combine(directory,
                    "chuanhoa-tone-20260903000000000-" + index.ToString("D12") + ".docx");
                File.WriteAllText(file, "owned backup " + index);
                File.SetLastWriteTimeUtc(file, DateTime.UtcNow.AddMinutes(-index));
                if (index == 21)
                {
                    oldOwnedFile = file;
                    File.SetLastWriteTimeUtc(file, DateTime.UtcNow.AddDays(-8));
                }
            }

            var manager = typeof(WordLocalCommandRuntime).Assembly.GetType(
                "ChuanHoa.AddIn.Vsto.Runtime.WordRecoveryCopyManager", true);
            var cleanup = manager.GetMethod("CleanupOwnedBackups",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(cleanup != null, "Backup cleanup entry point was not found.");
            cleanup.Invoke(null, new object[] { directory });

            Assert(!File.Exists(oldOwnedFile), "An owned backup older than seven days was not deleted.");
            Assert(File.Exists(nonOwnedFile), "Cleanup deleted a file it does not own.");
            Assert(Directory.GetFiles(directory, "chuanhoa-tone-*.docx").Length <= 20,
                "Backup cleanup retained more than 20 owned backups.");
        }

        private static void AssertBackup(string result, string sourcePath, string operation)
        {
            Assert(!string.IsNullOrWhiteSpace(result) && File.Exists(result),
                operation + " did not create a recovery backup.");
            var expectedDirectory = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "ChuanHoa", "Backups"));
            Assert(string.Equals(Path.GetDirectoryName(Path.GetFullPath(result)), expectedDirectory,
                    StringComparison.OrdinalIgnoreCase),
                operation + " created its backup outside Windows Temp.");
            Assert(string.Equals(Path.GetExtension(result), Path.GetExtension(sourcePath),
                    StringComparison.OrdinalIgnoreCase),
                operation + " did not preserve the source format.");
        }

        private static void AssertFailsWithoutSelection(Func<string> command)
        {
            try
            {
                command();
                throw new InvalidOperationException("A collapsed selection was accepted.");
            }
            catch (InvalidOperationException exception)
            {
                Assert(exception.Message.Contains("chọn phần chữ"),
                    "Collapsed-selection error was not actionable: " + exception.Message);
            }
        }

        private static void AssertActiveDocument(
            Word.Application application,
            Word.Document expected,
            string operation)
        {
            var active = application.ActiveDocument;
            Assert(active != null, operation + " left Word without an active document.");
            Assert(string.Equals(active.FullName, expected.FullName, StringComparison.OrdinalIgnoreCase),
                operation + " did not restore the user's source document as active.");
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
