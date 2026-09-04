using Office = Microsoft.Office.Core;

namespace ChuanHoa.AddIn.Vsto.Ribbon
{
    public interface IChuanHoaRibbonRuntime
    {
        void AttachRibbon(Office.IRibbonUI ribbonUi);

        bool IsEnabled(string controlId);

        int GetSelectedItemIndex(string controlId);

        void SelectDropDownItem(string controlId, string selectedId, int selectedIndex);

        int GetItemCount(string controlId);

        string GetItemLabel(string controlId, int index);

        object GetImage(string controlId);

        void ExecuteButton(string controlId);
    }
}
