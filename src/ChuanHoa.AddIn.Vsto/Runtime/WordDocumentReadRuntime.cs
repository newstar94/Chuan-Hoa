using System;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    [Flags]
    public enum DocumentAnalysisScope
    {
        SnapshotOnly = 0,
        Format = 1,
        Spelling = 2,
        Full = Format | Spelling
    }

    public sealed class WordDocumentReadResult
    {
        public WordDocumentReadResult(WordDocumentSnapshot snapshot, string documentTypeCode,
            LocalScanResult formatScan, LocalScanResult spellingScan)
        {
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            DocumentTypeCode = documentTypeCode ?? LocalDocumentTypeCodes.Unknown;
            FormatScan = formatScan ?? throw new ArgumentNullException(nameof(formatScan));
            SpellingScan = spellingScan ?? throw new ArgumentNullException(nameof(spellingScan));
        }

        public WordDocumentSnapshot Snapshot { get; }
        public string DocumentTypeCode { get; }
        public LocalScanResult FormatScan { get; }
        public LocalScanResult SpellingScan { get; }
    }

    /// <summary>
    /// The single production boundary allowed to read and scan an entire document.
    /// It is invoked only by an explicit Ribbon command and prepares the minimum
    /// analysis lane required by that command. Window activation and startup never
    /// call this runtime.
    /// </summary>
    public sealed class WordDocumentReadRuntime
    {
        private readonly Word.Application _application;
        private readonly WordDocumentCapabilityProvider _capabilityProvider;
        private readonly WordDocumentSnapshotBuilder _snapshotBuilder = new WordDocumentSnapshotBuilder();
        private readonly LocalAccessManager _accessManager;
        private readonly LocalDocumentScanner _scanner = new LocalDocumentScanner();

        public WordDocumentReadRuntime(Word.Application application, LocalAccessManager accessManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _capabilityProvider = new WordDocumentCapabilityProvider(application);
            _accessManager = accessManager ?? throw new ArgumentNullException(nameof(accessManager));
        }

        public WordDocumentReadResult Read(DocumentContext context, Word.Document? activeDocument = null)
        {
            Prepare(context, DocumentAnalysisScope.Full, activeDocument, false);
            return new WordDocumentReadResult(
                context.LastSnapshot!,
                context.DocumentTypeCode,
                context.LastFormatScan!,
                context.LastSpellingScan!);
        }

        public bool Prepare(
            DocumentContext context,
            DocumentAnalysisScope scope,
            Word.Document? activeDocument = null,
            bool allowSavedDocumentCache = true,
            DocumentOperationSession? operation = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            var document = activeDocument ?? _application.ActiveDocument;
            var capability = _capabilityProvider.Evaluate(document);
            if (!capability.CanReadDocument) throw new InvalidOperationException(capability.Reason);

            var requireFormat = (scope & DocumentAnalysisScope.Format) != 0;
            var requireSpelling = (scope & DocumentAnalysisScope.Spelling) != 0;
            LocalRulePack? formatRules = requireFormat
                ? _accessManager.GetRulePack(LocalAccessManager.FormatFeature)
                : null;
            LocalRulePack? spellingRules = requireSpelling
                ? _accessManager.GetRulePack(LocalAccessManager.SpellingFeature)
                : null;
            var documentIsSaved = ReadSavedState(document);
            if (allowSavedDocumentCache && context.CanReuseAnalysis(
                    documentIsSaved,
                    requireFormat,
                    requireSpelling,
                    formatRules?.PackId,
                    spellingRules?.PackId))
                return true;

            var snapshot = _snapshotBuilder.Build(document, context, capability, operation);
            var initialLocal = WordLocalScanRuntime.ToLocalSnapshot(snapshot, context);
            WordDocumentTypeClassifier.DetectAndApply(context, initialLocal, false);
            var local = WordLocalScanRuntime.ToLocalSnapshot(snapshot, context);
            operation?.Transition(DocumentOperationState.Scanning,
                requireFormat && requireSpelling ? "kiểm tra thể thức và chính tả" :
                requireFormat ? "kiểm tra thể thức" :
                requireSpelling ? "kiểm tra chính tả" : "chuẩn bị dữ liệu");
            var formatScan = formatRules == null ? null :
                _scanner.ScanFormat(local, formatRules,
                    operation == null ? default(System.Threading.CancellationToken) : operation.CancellationToken);
            operation?.Checkpoint();
            var spellingScan = spellingRules == null ? null :
                _scanner.ScanSpelling(local, spellingRules,
                    operation == null ? default(System.Threading.CancellationToken) : operation.CancellationToken);
            operation?.Checkpoint();
            context.SetAnalysis(snapshot, local, formatScan, spellingScan, documentIsSaved);
            return false;
        }

        public void PrepareForOneClick(
            DocumentContext context,
            Word.Document? activeDocument = null,
            DocumentOperationSession? operation = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            var document = activeDocument ?? _application.ActiveDocument;

            // Saving an unsaved document raises DocumentBeforeSave. RibbonRuntime
            // intentionally invalidates the old snapshot in that event. Persist first,
            // then force a fresh full analysis so the save cannot erase the snapshot
            // between Prepare and WordOneClickRuntime.Execute.
            WordRecoveryCopyManager.EnsurePersistentDocument(_application, document);
            context.ClearReadAnalysis();
            Prepare(context, DocumentAnalysisScope.Full, document, false, operation);
            context.RequireFullAnalysis();
        }

        private static bool ReadSavedState(Word.Document document)
        {
            try { return document.Saved; }
            catch (System.Runtime.InteropServices.COMException) { return false; }
        }
    }
}
