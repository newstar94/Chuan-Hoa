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
        private System.Windows.Forms.Timer? _runtimeStartupTimer;
        private bool _startupCompleted;
        private bool _ribbonLoaded;
        private bool _applicationEventsSubscribed;
        private bool _shutdownStarted;
        private int _runtimeInitializationAttempts;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                SubscribeApplicationEvents();
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("Startup event subscription", exception);
            }

            // Word does not guarantee whether VSTO Startup or Ribbon discovery
            // happens first. Startup must therefore remain failure-contained and
            // must never construct services, read ActiveDocument, or invalidate
            // CustomUI. The runtime is scheduled only after RibbonOnLoad has
            // returned control to Office's message loop.
            _startupCompleted = true;
            TryScheduleRuntimeInitialization(250);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _shutdownStarted = true;
            StopRuntimeStartupTimer();

            if (_applicationEventsSubscribed)
            {
                try
                {
                    Application.WindowActivate -= OnWindowActivate;
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
                _ribbon.RibbonLoaded -= OnRibbonLoaded;
                _ribbon.DetachRuntime();
            }

            _wordMutationRuntime = null;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            // Office requests IRibbonExtensibility before the VSTO Startup event on
            // some Word launches. Keep this path limited to constructing the Ribbon
            // adapter: touching Word COM, the local access cache, timers, or scanners
            // here can fail transiently while leaving COMAddIn.Connect=true and no tab.
            if (_ribbon == null)
            {
                _ribbon = new ChuanHoaRibbon();
                _ribbon.RibbonLoaded += OnRibbonLoaded;
                if (_ribbonRuntime != null)
                {
                    _ribbon.AttachRuntime(_ribbonRuntime);
                }
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
                return ribbonRuntime;
            }
            catch
            {
                documentContexts.Dispose();
                throw;
            }
        }

        private void SubscribeApplicationEvents()
        {
            if (_applicationEventsSubscribed) return;

            Application.WindowActivate += OnWindowActivate;
            Application.DocumentBeforeClose += OnDocumentBeforeClose;
            Application.DocumentBeforeSave += OnDocumentBeforeSave;
            _applicationEventsSubscribed = true;
        }

        private void OnRibbonLoaded(object? sender, EventArgs eventArgs)
        {
            try
            {
                _ribbonLoaded = true;
                TryScheduleRuntimeInitialization(250);
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("Ribbon load handshake", exception);
            }
        }

        private void TryScheduleRuntimeInitialization(int intervalMilliseconds)
        {
            if (_shutdownStarted || !_startupCompleted || !_ribbonLoaded ||
                _ribbonRuntime != null || _runtimeStartupTimer != null)
            {
                return;
            }

            try
            {
                var timer = new System.Windows.Forms.Timer
                {
                    Interval = Math.Max(1, intervalMilliseconds)
                };
                timer.Tick += OnRuntimeStartupTimerTick;
                _runtimeStartupTimer = timer;
                timer.Start();
            }
            catch (Exception exception)
            {
                StopRuntimeStartupTimer();
                TraceLifecycleFailure("Deferred runtime scheduling", exception);
            }
        }

        private void OnRuntimeStartupTimerTick(object? sender, EventArgs eventArgs)
        {
            StopRuntimeStartupTimer();
            if (_shutdownStarted || _ribbonRuntime != null) return;

            _runtimeInitializationAttempts++;
            try
            {
                var runtime = EnsureRuntime();
                _ribbon?.AttachRuntime(runtime);
                CaptureActiveDocumentMetadata(runtime);
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("Deferred runtime initialization", exception);

                // A document/template may still be opening, especially through
                // Word Home/Recent or OneDrive. Keep the static tab visible and
                // retry later without reading document content.
                if (_runtimeInitializationAttempts < 3)
                {
                    TryScheduleRuntimeInitialization(1500);
                }
            }
        }

        private void CaptureActiveDocumentMetadata(RibbonRuntime runtime)
        {
            try
            {
                var activeDocument = Application.ActiveDocument;
                if (activeDocument != null)
                {
                    runtime.OnDocumentWindowActivated(activeDocument);
                }
            }
            catch (Exception exception)
            {
                // This is metadata-only capability capture. Full content reads and
                // scans remain command-scoped. A transient COM rejection must not
                // remove the Ribbon or fail the add-in startup.
                TraceLifecycleFailure("Initial document metadata", exception);
            }
        }

        private void StopRuntimeStartupTimer()
        {
            var timer = _runtimeStartupTimer;
            _runtimeStartupTimer = null;
            if (timer == null) return;

            try
            {
                timer.Stop();
                timer.Tick -= OnRuntimeStartupTimerTick;
                timer.Dispose();
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("Deferred runtime timer cleanup", exception);
            }
        }

        private static void TraceLifecycleFailure(string stage, Exception exception)
        {
            Trace.TraceError("ChuanHoa VSTO lifecycle failure at {0}: {1}", stage, exception);
        }

        private void OnWindowActivate(Word.Document document, Word.Window window)
        {
            try
            {
                var runtime = _ribbonRuntime;
                if (runtime != null)
                {
                    runtime.OnDocumentWindowActivated(document);
                }
                else
                {
                    TryScheduleRuntimeInitialization(250);
                }
            }
            catch (Exception exception)
            {
                TraceLifecycleFailure("WindowActivate", exception);
            }
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
