using ChuanHoa.AddIn.Vsto.Ribbon;
using ChuanHoa.AddIn.Vsto.Runtime;
using System;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto
{
    public partial class ThisAddIn
    {
        private DocumentContextStore? _documentContexts;
        private WordMutationRuntime? _wordMutationRuntime;
        private RibbonRuntime? _ribbonRuntime;
        private ChuanHoaRibbon? _ribbon;
        private bool _applicationEventsSubscribed;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            StartupIntegrityGuard.VerifyOrThrow();
            // Startup is deliberately inert. Do not even subscribe to document
            // lifecycle events until the first explicit Ribbon action creates the
            // runtime. Opening Word or opening/activating a document must not cause
            // this add-in to inspect Word state, refresh a lease, start a timer or
            // invalidate CustomUI.
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_applicationEventsSubscribed)
            {
                try
                {
                    Application.DocumentBeforeClose -= OnDocumentBeforeClose;
                    Application.DocumentBeforeSave -= OnDocumentBeforeSave;
                }
                catch (Exception exception)
                {
                    TraceLifecycleFailure("Shutdown event removal", exception);
                }
                _applicationEventsSubscribed = false;
            }

            if (_ribbonRuntime != null)
            {
                try
                {
                    _ribbonRuntime.Dispose();
                }
                catch (Exception exception)
                {
                    TraceLifecycleFailure("Runtime shutdown", exception);
                }
                _ribbonRuntime = null;
            }

            if (_documentContexts != null)
            {
                _documentContexts.Dispose();
                _documentContexts = null;
            }

            if (_ribbon != null)
            {
                _ribbon.DetachRuntime();
            }

            _wordMutationRuntime = null;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            StartupIntegrityGuard.VerifyOrThrow();
            // Office requests IRibbonExtensibility before the VSTO Startup event on
            // some Word launches. Keep this path limited to constructing the Ribbon
            // adapter: touching Word COM, the local access cache, timers, or scanners
            // here can fail transiently while leaving COMAddIn.Connect=true and no tab.
            if (_ribbon == null)
            {
                _ribbon = new ChuanHoaRibbon();
                _ribbon.ConfigureRuntimeFactory(EnsureRuntime);
            }

            return _ribbon;
        }

        private RibbonRuntime EnsureRuntime()
        {
            if (_ribbonRuntime != null)
            {
                return _ribbonRuntime;
            }

            var documentContexts = new DocumentContextStore();
            try
            {
                var wordMutationRuntime = new WordMutationRuntime(Application);
                var ribbonRuntime = new RibbonRuntime(Application, documentContexts, wordMutationRuntime);
                _documentContexts = documentContexts;
                _wordMutationRuntime = wordMutationRuntime;
                _ribbonRuntime = ribbonRuntime;
                SubscribeApplicationEvents();
                return ribbonRuntime;
            }
            catch
            {
                if (_ribbonRuntime != null)
                {
                    try { _ribbonRuntime.Dispose(); }
                    catch (Exception exception) { TraceLifecycleFailure("Runtime initialization rollback", exception); }
                }
                _ribbonRuntime = null;
                _wordMutationRuntime = null;
                _documentContexts = null;
                documentContexts.Dispose();
                throw;
            }
        }

        private void SubscribeApplicationEvents()
        {
            if (_applicationEventsSubscribed) return;

            Application.DocumentBeforeClose += OnDocumentBeforeClose;
            Application.DocumentBeforeSave += OnDocumentBeforeSave;
            _applicationEventsSubscribed = true;
        }

        private static void TraceLifecycleFailure(string stage, Exception exception)
        {
            Trace.TraceError("ChuanHoa VSTO lifecycle failure at {0}: {1}", stage, exception);
        }

        private void OnDocumentBeforeClose(Word.Document document, ref bool cancel)
        {
            try
            {
                if (!cancel)
                {
                    var runtime = _ribbonRuntime;
                    if (runtime != null)
                    {
                        runtime.OnDocumentBeforeClose(document);
                    }
                }
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("DocumentBeforeClose", exception);
            }
        }

        private void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUi, ref bool cancel)
        {
            try
            {
                var runtime = _ribbonRuntime;
                if (runtime != null)
                {
                    runtime.OnDocumentBeforeSave(document);
                }
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("DocumentBeforeSave", exception);
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
