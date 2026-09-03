using System;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace VietDocStandardizer
{
    public partial class ThisAddIn
    {
        public static ThisAddIn Instance { get; private set; }
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Instance = this;
            _taskPaneControl = new TaskPaneControl();
            _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, "Chuẩn Hóa Thể Thức (Word 2013/2016)");
            _taskPane.Width = 360;
            _taskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _taskPane = null;
            _taskPaneControl = null;
        }

        public void ToggleTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonStandardizer();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
