using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ChuanHoa.AddIn.Vsto.Ribbon
{
    [ComVisible(true)]
    public sealed partial class ChuanHoaRibbon : Office.IRibbonExtensibility
    {
        private const string ResourceName = "ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml";
        private IChuanHoaRibbonRuntime? _runtime;
        private Office.IRibbonUI? _ribbonUi;

        public ChuanHoaRibbon()
        {
        }

        internal event EventHandler? RibbonLoaded;

        private IChuanHoaRibbonRuntime Runtime => _runtime ?? UnavailableRibbonRuntime.Instance;

        internal void AttachRuntime(IChuanHoaRibbonRuntime runtime)
        {
            _runtime = runtime ?? throw new ArgumentNullException("runtime");
            if (_ribbonUi != null)
            {
                try
                {
                    runtime.AttachRibbon(_ribbonUi);
                }
                catch (Exception exception)
                {
                    Trace.TraceError("ChuanHoa could not attach its Ribbon runtime: {0}", exception);
                    return;
                }

                try
                {
                    _ribbonUi.Invalidate();
                }
                catch (Exception exception)
                {
                    // Word can transiently reject Ribbon invalidation while it is
                    // still opening Home/Recent/OneDrive documents. The tab is
                    // already loaded, so this must never escape into Office.
                    Trace.TraceError("ChuanHoa could not invalidate its Ribbon: {0}", exception);
                }
            }
        }

        internal void DetachRuntime()
        {
            _runtime = null;
        }

        private void CompleteRibbonLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
            try
            {
                Runtime.AttachRibbon(ribbonUi);
            }
            catch (Exception exception)
            {
                Trace.TraceError("ChuanHoa RibbonOnLoad runtime callback failed: {0}", exception);
            }

            try
            {
                RibbonLoaded?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception exception)
            {
                // Office treats an exception escaping RibbonOnLoad as a CustomUI
                // load failure. Keep the static tab available even if a lifecycle
                // subscriber fails.
                Trace.TraceError("ChuanHoa RibbonOnLoad handshake failed: {0}", exception);
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(ResourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "Embedded Ribbon resource was not found: " + ResourceName);
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static string RequireControlId(Office.IRibbonControl control)
        {
            if (control == null || string.IsNullOrWhiteSpace(control.Id))
            {
                throw new ArgumentException("Ribbon callback requires a control with a stable id.", "control");
            }

            return control.Id;
        }

        private sealed class UnavailableRibbonRuntime : IChuanHoaRibbonRuntime
        {
            internal static readonly UnavailableRibbonRuntime Instance = new UnavailableRibbonRuntime();

            private UnavailableRibbonRuntime()
            {
            }

            public void AttachRibbon(Office.IRibbonUI ribbonUi)
            {
            }

            public bool IsEnabled(string controlId) => false;

            public bool GetPressed(string controlId) => false;

            public void SetPressed(string controlId, bool pressed)
            {
            }

            public int GetSelectedItemIndex(string controlId) => 0;

            public void SelectDropDownItem(string controlId, string selectedId, int selectedIndex)
            {
            }

            public int GetItemCount(string controlId) => 0;

            public string GetItemLabel(string controlId, int index) => string.Empty;

            public object GetImage(string controlId) => null!;

            public void ExecuteButton(string controlId)
            {
            }
        }
    }
}
