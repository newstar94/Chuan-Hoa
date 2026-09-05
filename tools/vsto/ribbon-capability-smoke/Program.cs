#nullable disable
using System;
using System.IO;
using System.Runtime.InteropServices;
using ChuanHoa.AddIn.Vsto.Runtime;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.RibbonCapabilitySmoke
{
    internal static class Program
    {
        private static readonly string[] ControlIds =
        {
            "btnAutoFixAll2026", "ddQuyDinh", "ddLoaiVanBan", "btnKiemTra",
            "btnKiemTraChinhTa", "btnChuyenDoiUnicode", "btnDinhDangTrangGiay",
            "btnChenTrangNgang", "btnChenTrangDoc", "btnXoaTrangThua",
            "mnuDungBoStyle", "btnDungBoStyleCo15", "btnDungBoStyleCo14",
            "btnDungBoStyleCo13", "btnCoChu15", "btnCoChu14", "btnCoChu13",
            "btnKeepWithNext", "btnChenSoTrang", "btnCoChu", "btnGianChuNormal",
            "btnGianChuRa", "btnLapDongTieuDe", "btnChuanHoaBang", "btnChuanHoaAnh",
            "btnCanDinhO", "btnCanGiuaO", "btnXoaKyTuThuaBangExcel", "mnuBoDau",
            "btnKieuOaUy", "btnKieuOaUy2", "btnDoiDauThapPhan",
            "mnuThongTinTienIch", "btnKiemTraPhienBanMoi", "btnGuiPhanHoi",
            "btnGioiThieu", "btnSuaLoiDangChon", "btnSuaTatCaChinhTa",
            "btnTuDienCaNhan"
        };

        [STAThread]
        private static int Main()
        {
            var directory = Path.Combine(Path.GetTempPath(),
                "ChuanHoaRibbonCapabilitySmoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                Verify(directory, false);
                Verify(directory, true);
                Console.WriteLine("RIBBON_CAPABILITY_WORD_SMOKE_PASS SAVED UNSAVED CONTROLS=" +
                    ControlIds.Length);
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("RIBBON_CAPABILITY_WORD_SMOKE_FAIL " + exception);
                return 1;
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        private static void Verify(string directory, bool saveDocument)
        {
            Word.Application application = null;
            Word.Document document = null;
            DocumentContextStore contexts = null;
            RibbonRuntime runtime = null;
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                document = application.Documents.Add();
                document.Content.Text = "Văn bản kiểm tra trạng thái Ribbon.\r";
                if (saveDocument)
                {
                    object fileName = Path.Combine(directory, "saved.docx");
                    object format = Word.WdSaveFormat.wdFormatXMLDocument;
                    document.SaveAs(ref fileName, ref format);
                }

                contexts = new DocumentContextStore();
                runtime = new RibbonRuntime(application, contexts,
                    new WordMutationRuntime(application));
                foreach (var controlId in ControlIds)
                    if (!runtime.IsEnabled(controlId))
                        throw new InvalidOperationException(
                            (saveDocument ? "Saved" : "Unsaved") +
                            " document disabled Ribbon control " + controlId + ".");

                var capability = new WordDocumentCapabilityProvider(application).Evaluate(document);
                if (!capability.CanReadDocument || capability.IsSaved != saveDocument)
                    throw new InvalidOperationException(
                        "Unexpected document capability: " + capability.ReasonCode +
                        ", IsSaved=" + capability.IsSaved + ".");
            }
            finally
            {
                if (runtime != null) runtime.Dispose();
                if (contexts != null) contexts.Dispose();
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

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
