using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ChuanHoa.AddIn.Vsto.Runtime;
using Office = Microsoft.Office.Core;

namespace ChuanHoa.AddIn.Vsto.Ribbon
{
    [ComVisible(true)]
    public sealed partial class ChuanHoaRibbon : Office.IRibbonExtensibility
    {
        private const string ResourceName = "ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml";
        private readonly DeferredRibbonRuntime _deferredRuntime;
        private IChuanHoaRibbonRuntime? _runtime;
        private Func<IChuanHoaRibbonRuntime>? _runtimeFactory;
        private Office.IRibbonUI? _ribbonUi;

        public ChuanHoaRibbon()
        {
            _deferredRuntime = new DeferredRibbonRuntime(this);
        }

        private IChuanHoaRibbonRuntime Runtime => _runtime ?? _deferredRuntime;

        internal void ConfigureRuntimeFactory(Func<IChuanHoaRibbonRuntime> runtimeFactory)
        {
            if (runtimeFactory == null) throw new ArgumentNullException("runtimeFactory");
            if (_runtimeFactory != null) throw new InvalidOperationException("Ribbon runtime factory is already configured.");
            _runtimeFactory = runtimeFactory;
        }

        internal void AttachRuntime(IChuanHoaRibbonRuntime runtime)
        {
            AttachRuntime(runtime, true);
        }

        private void AttachRuntime(IChuanHoaRibbonRuntime runtime, bool invalidateRibbon)
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

                if (!invalidateRibbon) return;
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
            _runtimeFactory = null;
            _deferredRuntime.Dispose();
        }

        private IChuanHoaRibbonRuntime GetOrCreateRuntimeForUserAction()
        {
            if (_runtime != null) return _runtime;
            var factory = _runtimeFactory;
            if (factory == null)
                throw new InvalidOperationException("Add-in chưa sẵn sàng nhận lệnh.");

            var runtime = factory();
            // This path runs inside an explicit onAction callback. Do not invalidate
            // Office while that callback is still on the stack; command completion
            // will invalidate the affected controls when needed.
            AttachRuntime(runtime, false);
            return runtime;
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

        private sealed class DeferredRibbonRuntime : IChuanHoaRibbonRuntime, IDisposable
        {
            private readonly ChuanHoaRibbon _owner;
            private readonly RibbonImageProvider _imageProvider = new RibbonImageProvider();
            private bool _disposed;

            internal DeferredRibbonRuntime(ChuanHoaRibbon owner)
            {
                _owner = owner;
            }

            public void AttachRibbon(Office.IRibbonUI ribbonUi)
            {
            }

            public bool IsEnabled(string controlId) => !_disposed;

            public int GetSelectedItemIndex(string controlId) => 0;

            public void SelectDropDownItem(string controlId, string selectedId, int selectedIndex)
            {
                Invoke(runtime => runtime.SelectDropDownItem(controlId, selectedId, selectedIndex));
            }

            public int GetItemCount(string controlId) =>
                string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal)
                    ? RibbonRuntime.DocumentTypeItemCount
                    : 0;

            public string GetItemLabel(string controlId, int index) =>
                string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal)
                    ? RibbonRuntime.GetDocumentTypeItemLabel(index)
                    : string.Empty;

            public object GetImage(string controlId) => _imageProvider.GetImage(controlId);

            public void ExecuteButton(string controlId)
            {
                Invoke(runtime => runtime.ExecuteButton(controlId));
            }

            private void Invoke(Action<IChuanHoaRibbonRuntime> action)
            {
                if (_disposed) return;
                try
                {
                    action(_owner.GetOrCreateRuntimeForUserAction());
                }
                catch (Exception exception)
                {
                    Trace.TraceError("ChuanHoa could not initialize its command runtime: {0}", exception);
                    MessageBox.Show(
                        "Không thể khởi tạo chức năng Chuẩn hóa.\n\n" + exception.Message,
                        "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public void Dispose()
            {
                if (_disposed) return;
                _imageProvider.Dispose();
                _disposed = true;
            }
        }
    }
}
