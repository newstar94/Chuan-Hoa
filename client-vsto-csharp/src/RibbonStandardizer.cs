using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VietDocStandardizer
{
    [ComVisible(true)]
    public class RibbonStandardizer : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this._ribbon = ribbonUI;
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("VietDocStandardizer.src.RibbonStandardizer.xml");
        }

        public void OnAutoFixClick(Office.IRibbonControl control)
        {
            WordInteropEngine.Perform1ClickAutoFix();
        }

        public void OnToggleSidebarClick(Office.IRibbonControl control)
        {
            ThisAddIn.Instance?.ToggleTaskPane();
        }

        public void OnRegimeND30Click(Office.IRibbonControl control)
        {
            WordInteropEngine.CurrentRegime = "ND30";
            MessageBox.Show("Đã chuyển sang chế độ: Nghị định 30/2020/NĐ-CP (Hành chính Nhà nước)", "Chuẩn Hóa Thể Thức", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnRegimeDangClick(Office.IRibbonControl control)
        {
            WordInteropEngine.CurrentRegime = "DANG_HD05";
            MessageBox.Show("Đã chuyển sang chế độ: Hướng dẫn 05-HD/VPTW (Văn bản Đảng)", "Chuẩn Hóa Thể Thức", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnRegimeViettelClick(Office.IRibbonControl control)
        {
            WordInteropEngine.CurrentRegime = "VIETTEL";
            MessageBox.Show("Đã chuyển sang chế độ: Quy chế Doanh nghiệp Viettel (QĐ 11095)", "Chuẩn Hóa Thể Thức", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnFixToneClick(Office.IRibbonControl control)
        {
            WordInteropEngine.NormalizeToneMarks();
        }

        public void OnFixIyClick(Office.IRibbonControl control)
        {
            WordInteropEngine.NormalizeIySpellings();
        }

        public void OnCleanSpacesClick(Office.IRibbonControl control)
        {
            WordInteropEngine.CleanSpacesAndPunctuation();
        }

        public void OnRepeatHeaderClick(Office.IRibbonControl control)
        {
            WordInteropEngine.RepeatTableHeaders();
        }

        public void OnPreventSplitClick(Office.IRibbonControl control)
        {
            WordInteropEngine.PreventTableRowSplitting();
        }

        public void OnRemoveBlankPagesClick(Office.IRibbonControl control)
        {
            WordInteropEngine.RemoveTrailingEmptyPages();
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            // Fallback direct file read
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "src", "RibbonStandardizer.xml");
            if (File.Exists(path)) return File.ReadAllText(path);
            return null;
        }
    }
}
