#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.PassiveStartupSmoke
{
    internal static class Program
    {
        private const string AddInProgId = "ChuanHoa.AddIn.Vsto";

        [STAThread]
        private static int Main()
        {
            var testDirectory = Path.Combine(Path.GetTempPath(),
                "ChuanHoaPassiveStartup-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(testDirectory);
            var documentPath = Path.Combine(testDirectory, "passive-open.docx");
            Word.Application application = null;
            Word.Document document = null;
            Office.COMAddIn addIn = null;
            try
            {
                var cacheBefore = CaptureFiles(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ChuanHoa", "Cache"));
                var backupsBefore = CaptureFiles(Path.Combine(
                    Path.GetTempPath(), "ChuanHoa", "Backups"));

                Console.WriteLine("PASSIVE_STEP=create-word");
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                addIn = application.COMAddIns.Item(AddInProgId);
                Assert(addIn != null && addIn.Connect,
                    "The installed Chuẩn hóa add-in is not connected.");
                Console.WriteLine("PASSIVE_ADDIN_CONNECTED=True");

                Console.WriteLine("PASSIVE_STEP=create-saved-document");
                document = application.Documents.Add();
                document.Content.Text = string.Join("\r", Enumerable.Repeat(
                    "Tài liệu kiểm tra trạng thái thụ động.", 80)) + "\r";
                document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;

                Console.WriteLine("PASSIVE_STEP=open-saved-document");
                document = application.Documents.Open(documentPath, ReadOnly: true,
                    AddToRecentFiles: false, Visible: false);
                var wordProcess = Process.GetProcessesByName("WINWORD")
                    .OrderByDescending(item => item.StartTime).FirstOrDefault();
                Assert(wordProcess != null, "WINWORD process was not found.");
                Thread.Sleep(750);
                wordProcess.Refresh();
                var cpuBefore = wordProcess.TotalProcessorTime.TotalMilliseconds;
                Thread.Sleep(TimeSpan.FromSeconds(5));
                wordProcess.Refresh();
                var cpuDelta = wordProcess.TotalProcessorTime.TotalMilliseconds - cpuBefore;

                var cacheAfter = CaptureFiles(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ChuanHoa", "Cache"));
                var backupsAfter = CaptureFiles(Path.Combine(
                    Path.GetTempPath(), "ChuanHoa", "Backups"));
                var cacheChanges = ChangedFiles(cacheBefore, cacheAfter);
                var newBackups = backupsAfter.Keys.Except(backupsBefore.Keys,
                    StringComparer.OrdinalIgnoreCase).ToArray();

                Console.WriteLine("PASSIVE_IDLE_CPU_DELTA_MS_5S=" +
                    Math.Round(cpuDelta, 2).ToString("0.00"));
                Console.WriteLine("PASSIVE_BACKGROUND_SAVING=" +
                    application.BackgroundSavingStatus);
                Console.WriteLine("PASSIVE_BACKGROUND_PRINTING=" +
                    application.BackgroundPrintingStatus);
                Console.WriteLine("PASSIVE_CACHE_CHANGES=" + cacheChanges.Length);
                Console.WriteLine("PASSIVE_NEW_BACKUPS=" + newBackups.Length);
                Assert(application.BackgroundSavingStatus == 0,
                    "Word still has background saving work after the passive wait.");
                Assert(application.BackgroundPrintingStatus == 0,
                    "Word still has background printing work after the passive wait.");
                Assert(cacheChanges.Length == 0,
                    "Opening a document without a Ribbon action changed add-in cache files: " +
                    string.Join(", ", cacheChanges));
                Assert(newBackups.Length == 0,
                    "Opening a document without a Ribbon action created a recovery backup.");

                Console.WriteLine("PASSIVE_STARTUP_WORD_SMOKE_PASS");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("PASSIVE_STARTUP_WORD_SMOKE_FAIL " + exception);
                return 1;
            }
            finally
            {
                if (document != null)
                {
                    try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                    catch { }
                }
                if (application != null)
                {
                    try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); }
                    catch { }
                }
                Release(document);
                Release(addIn);
                Release(application);
                try { Directory.Delete(testDirectory, true); }
                catch { }
            }
        }

        private static Dictionary<string, long> CaptureFiles(string directory)
        {
            if (!Directory.Exists(directory))
                return new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            return Directory.GetFiles(directory, "*", SearchOption.AllDirectories)
                .ToDictionary(path => path, path => File.GetLastWriteTimeUtc(path).Ticks,
                    StringComparer.OrdinalIgnoreCase);
        }

        private static string[] ChangedFiles(IReadOnlyDictionary<string, long> before,
            IReadOnlyDictionary<string, long> after)
        {
            return after.Where(item => !before.TryGetValue(item.Key, out var oldTicks) ||
                    oldTicks != item.Value)
                .Select(item => item.Key).ToArray();
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.FinalReleaseComObject(value);
        }
    }
}
