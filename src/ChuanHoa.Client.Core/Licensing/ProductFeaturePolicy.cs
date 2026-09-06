using System;
using System.Collections.Generic;

namespace ChuanHoa.Client.Core.Licensing
{
    public static class ProductFeaturePolicy
    {
        public const string TableImageTools = "TABLE_IMAGE_TOOLS";
        public static IReadOnlyList<string> FreeFeatures { get; } = Array.AsReadOnly(new[]
            { "FORMAT_SCAN", "SPELLING_SCAN", TableImageTools });

        public static string RequiredFeature(string commandId)
        {
            switch (commandId)
            {
                case "btnKiemTra": return "FORMAT_SCAN";
                case "btnKiemTraChinhTa": return "SPELLING_SCAN";
                case "btnLapDongTieuDe": case "btnChuanHoaBang": case "btnChuanHoaAnh":
                case "btnCanDinhO": case "btnCanGiuaO": case "btnXoaKyTuThuaBangExcel":
                    return TableImageTools;
                case "btnAutoFixAll2026": case "btnSuaLoiDangChon": case "btnSuaTatCaChinhTa":
                    return "AUTOFIX";
                case "btnThietLap": case "btnGioiThieu": case "btnKiemTraPhienBanMoi":
                case "ddQuyDinh": case "ddLoaiVanBan": case "mnuThongTinTienIch":
                    return string.Empty;
                default: return "DOCUMENT_TOOLS";
            }
        }
    }
}
