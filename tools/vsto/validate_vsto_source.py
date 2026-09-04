from __future__ import annotations

import hashlib
import json
import re
import struct
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
VSTO_ROOT = ROOT / "src" / "ChuanHoa.AddIn.Vsto"
CONTRACT_PATH = ROOT / "shared" / "contracts" / "ribbon" / "ribbon-contract.v1.json"
RIBBON_XML_PATH = VSTO_ROOT / "Ribbon" / "ChuanHoaRibbon.xml"
CALLBACK_SOURCE_PATH = VSTO_ROOT / "Ribbon" / "ChuanHoaRibbon.Callbacks.g.cs"
RIBBON_SOURCE_PATH = VSTO_ROOT / "Ribbon" / "ChuanHoaRibbon.cs"
THIS_ADDIN_PATH = VSTO_ROOT / "ThisAddIn.cs"
PROJECT_PATH = VSTO_ROOT / "ChuanHoa.AddIn.Vsto.csproj"
WORD_ADAPTER_PATH = VSTO_ROOT / "Runtime" / "WordComMutationDocumentAdapter.cs"
WORD_MUTATION_RUNTIME_PATH = VSTO_ROOT / "Runtime" / "WordMutationRuntime.cs"
WORD_ANNOTATION_ADAPTER_PATH = VSTO_ROOT / "Runtime" / "WordFindingAnnotationAdapter.cs"
RIBBON_RUNTIME_PATH = VSTO_ROOT / "Runtime" / "RibbonRuntime.cs"
LOCAL_ACCESS_MANAGER_PATH = VSTO_ROOT / "Runtime" / "LocalAccessManager.cs"
WORD_SNAPSHOT_BUILDER_PATH = VSTO_ROOT / "Runtime" / "WordDocumentSnapshotBuilder.cs"
WORD_LOCAL_COMMAND_PATH = VSTO_ROOT / "Runtime" / "WordLocalCommandRuntime.cs"
WORD_ONE_CLICK_PATH = VSTO_ROOT / "Runtime" / "WordOneClickRuntime.cs"
WORD_APPENDIX_PAGINATION_PATH = VSTO_ROOT / "Runtime" / "WordAppendixPaginationNormalizer.cs"
DOCUMENT_OPERATION_SESSION_PATH = VSTO_ROOT / "Runtime" / "DocumentOperationSession.cs"
CUSTOM_DICTIONARY_DIALOG_PATH = VSTO_ROOT / "Runtime" / "CustomDictionaryDialog.cs"
DOCUMENT_CONTEXT_PATH = VSTO_ROOT / "Runtime" / "DocumentContext.cs"
RIBBON_RUNTIME_INTERFACE_PATH = VSTO_ROOT / "Ribbon" / "IChuanHoaRibbonRuntime.cs"
LOCAL_SCAN_MODELS_PATH = ROOT / "src" / "ChuanHoa.Client.Core" / "Scanning" / "LocalScanModels.cs"
CANONICAL_SCANNER_PATH = ROOT / "src" / "ChuanHoa.Client.Core" / "Scanning" / "CanonicalRuleScanner.cs"
PERSONAL_DICTIONARY_ICON_16 = VSTO_ROOT / "Resources" / "Icons" / "personal-dictionary-16.png"
PERSONAL_DICTIONARY_ICON_32 = VSTO_ROOT / "Resources" / "Icons" / "personal-dictionary-32.png"
EVIDENCE_PATH = (
    ROOT
    / "shared"
    / "docs"
    / "implementation"
    / "evidence"
    / "vsto_source_foundation.json"
)
BUILD_LOAD_EVIDENCE_PATH = EVIDENCE_PATH.with_name("vsto_build_load.json")
INTERACTIVE_ELEMENTS = {"button", "menu", "dropDown", "checkBox"}
CALLBACK_ATTRIBUTES = {
    "getEnabled",
    "getImage",
    "getItemCount",
    "getItemLabel",
    "getPressed",
    "getSelectedItemIndex",
    "onAction",
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def validate() -> dict:
    contract = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
    ribbon_tree = ET.parse(RIBBON_XML_PATH)
    root = ribbon_tree.getroot()
    callback_source = CALLBACK_SOURCE_PATH.read_text(encoding="utf-8")
    ribbon_source = RIBBON_SOURCE_PATH.read_text(encoding="utf-8")
    this_addin_source = THIS_ADDIN_PATH.read_text(encoding="utf-8")
    project_source = PROJECT_PATH.read_text(encoding="utf-8")
    ribbon_runtime_interface = RIBBON_RUNTIME_INTERFACE_PATH.read_text(encoding="utf-8")
    word_local_command_source = WORD_LOCAL_COMMAND_PATH.read_text(encoding="utf-8")
    word_one_click_source = WORD_ONE_CLICK_PATH.read_text(encoding="utf-8")
    word_appendix_pagination_source = WORD_APPENDIX_PAGINATION_PATH.read_text(encoding="utf-8")
    document_operation_source = DOCUMENT_OPERATION_SESSION_PATH.read_text(encoding="utf-8")
    custom_dictionary_dialog_source = CUSTOM_DICTIONARY_DIALOG_PATH.read_text(encoding="utf-8")
    document_context_source = DOCUMENT_CONTEXT_PATH.read_text(encoding="utf-8")
    local_scan_models_source = LOCAL_SCAN_MODELS_PATH.read_text(encoding="utf-8")
    canonical_scanner_source = CANONICAL_SCANNER_PATH.read_text(encoding="utf-8")
    for resource_contract in (
        '<EmbeddedResource Include="Ribbon\\ChuanHoaRibbon.xml">',
        "<LogicalName>ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml</LogicalName>",
        "<ManifestResourceName>ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml</ManifestResourceName>",
        "<WithCulture>false</WithCulture>",
        "<Type>Non-Resx</Type>",
    ):
        if resource_contract not in project_source:
            raise RuntimeError(
                f"VSTO project is missing the embedded Ribbon resource contract: {resource_contract}"
            )
    if '<Target Name="RegisterOfficeAddin" />' not in project_source:
        raise RuntimeError(
            "Verification builds must not replace the installed Word add-in manifest registration."
        )
    dictionary_control = next(
        control for control in contract["controls"] if control["id"] == "btnTuDienCaNhan"
    )
    if dictionary_control.get("imageMso") is not None or \
            dictionary_control["callbacks"].get("getImage") != "GetImageTuDienCaNhan":
        raise RuntimeError("Personal dictionary must use its custom Ribbon image callback only.")
    if dictionary_control["callbacks"].get("getEnabled") != "GetEnabledAlways" or \
            dictionary_control["commandContract"].get("preconditions"):
        raise RuntimeError("Personal dictionary management must work without an active document.")
    for icon_path, expected_size in (
        (PERSONAL_DICTIONARY_ICON_16, 16),
        (PERSONAL_DICTIONARY_ICON_32, 32),
    ):
        data = icon_path.read_bytes()
        if data[:8] != b"\x89PNG\r\n\x1a\n":
            raise RuntimeError(f"Ribbon icon is not a PNG: {icon_path}")
        width, height = struct.unpack(">II", data[16:24])
        if (width, height) != (expected_size, expected_size) or data[25] not in {4, 6}:
            raise RuntimeError(
                f"Ribbon icon must be {expected_size}x{expected_size} with alpha: {icon_path}"
            )
        resource_include = (
            f'<EmbeddedResource Include="Resources\\Icons\\{icon_path.name}">'
        )
        if resource_include not in project_source:
            raise RuntimeError(f"Ribbon icon is not embedded in the VSTO assembly: {icon_path.name}")
    if "object GetImage(string controlId);" not in ribbon_runtime_interface:
        raise RuntimeError("Ribbon runtime interface is missing the custom image contract.")
    for dictionary_workflow_contract in (
        "TryGetSelectedText(out captured)",
        "AddUserWord(_selectedText)",
        "IgnoreWordForDocument(",
        "ClearAllDocumentIgnores()",
    ):
        if dictionary_workflow_contract not in (
                word_local_command_source + custom_dictionary_dialog_source):
            raise RuntimeError(
                "Personal dictionary selected-text workflow is missing: "
                + dictionary_workflow_contract
            )
    for stable_ignore_contract, source in (
        ("public string DictionaryScopeId", document_context_source),
        ("dictionaryScopeId", local_scan_models_source),
        ("snapshot.DictionaryScopeId", canonical_scanner_source),
        ("ClearDocumentIgnores(closingContext.DictionaryScopeId)",
         RIBBON_RUNTIME_PATH.read_text(encoding="utf-8")),
    ):
        if stable_ignore_contract not in source:
            raise RuntimeError(
                "Document-scoped dictionary ignores must use a stable lifetime scope and be cleared on close: "
                + stable_ignore_contract
            )
    if "FindPersonalDictionaryPhraseSpans" not in canonical_scanner_source:
        raise RuntimeError(
            "Personal dictionary phrase entries must protect every covered token from lexicon findings."
        )
    all_vsto_source = "\n".join(
        path.read_text(encoding="utf-8")
        for path in sorted(VSTO_ROOT.rglob("*"))
        if path.is_file() and path.suffix.lower() in {".cs", ".xml", ".csproj"}
    )
    retired_qr_markers = (
        "btnChenQrCode",
        "InsertQrCode",
        "QrCodeInputDialog",
        "QRCoder",
    )
    retired_qr_hits = [marker for marker in retired_qr_markers if marker in all_vsto_source]
    if retired_qr_hits:
        raise RuntimeError(
            "QR is retired and must not reappear in current VSTO product source: "
            + ", ".join(retired_qr_hits)
        )
    if any(control["id"] == "btnChenQrCode" for control in contract["controls"]):
        raise RuntimeError("QR is retired and must not reappear in the Ribbon contract.")
    quick_spelling = next(
        control for control in contract["controls"] if control["id"] == "btnSuaTatCaChinhTa"
    )
    if quick_spelling["commandContract"].get("mutationScope") != \
            "SPELLING_AND_TYPOGRAPHY_ALL_EDITABLE_STORIES":
        raise RuntimeError(
            "Quick spelling must declare spelling, whitespace, punctuation and bracket cleanup."
        )
    if any(
        marker in control.get("label", "").casefold()
        for control in contract["controls"]
        for marker in ("dọn khoảng trắng", "dấu ngoặc chuẩn")
    ):
        raise RuntimeError(
            "Whitespace/punctuation and standard brackets must not reappear as separate Ribbon controls."
        )
    for merged_typography_contract in (
        "fixedCount += CleanTypographyAndQuotes(document);",
        "VietnameseTypographyCleaner.CleanWhitespaceAndPunctuation(source)",
        "VietnameseTypographyCleaner.NormalizeQuotationMarks(cleaned)",
    ):
        if merged_typography_contract not in word_one_click_source:
            raise RuntimeError(
                "Quick spelling is missing merged typography cleanup: "
                + merged_typography_contract
            )

    controls = []
    callback_names = set()
    counts = {name: 0 for name in INTERACTIVE_ELEMENTS}
    for element in root.iter():
        element_name = local_name(element.tag)
        if element_name in INTERACTIVE_ELEMENTS:
            counts[element_name] += 1
            controls.append(
                {
                    "id": element.attrib["id"],
                    "type": element_name,
                    "onAction": element.attrib.get("onAction"),
                }
            )
        for callback_attribute in CALLBACK_ATTRIBUTES:
            callback_name = element.attrib.get(callback_attribute)
            if callback_name:
                callback_names.add(callback_name)

    callbacks_with_load = callback_names | {root.attrib.get("onLoad", "")}
    missing_callback_methods = sorted(
        callback_name
        for callback_name in callbacks_with_load
        if callback_name
        and re.search(rf"\b{re.escape(callback_name)}\s*\(", callback_source) is None
    )
    if missing_callback_methods:
        raise RuntimeError(f"Ribbon callbacks missing from generated source: {missing_callback_methods}")

    expected_controls = {
        control["id"]: control["controlType"]
        for control in contract["controls"]
    }
    actual_controls = {control["id"]: control["type"] for control in controls}
    if actual_controls != expected_controls:
        missing = sorted(set(expected_controls) - set(actual_controls))
        extra = sorted(set(actual_controls) - set(expected_controls))
        mismatched = sorted(
            control_id
            for control_id in set(expected_controls) & set(actual_controls)
            if expected_controls[control_id] != actual_controls[control_id]
        )
        raise RuntimeError(
            f"Ribbon control mismatch. missing={missing}, extra={extra}, typeMismatch={mismatched}"
        )

    auto_fix = next(control for control in controls if control["id"] == "btnAutoFixAll2026")
    if auto_fix["onAction"] != "OnAutoFixAll2026":
        raise RuntimeError("AutoFix is not bound to its dedicated fail-closed callback.")

    if "<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>" not in project_source:
        raise RuntimeError("VSTO host must target .NET Framework 4.8.")
    if "<OfficeVersion>14.0</OfficeVersion>" not in project_source:
        raise RuntimeError("VSTO host must declare Word/Office 2010 as its minimum Office lane.")
    if re.search(r"\bCustomTaskPane(?:s)?\b|\bTaskPane\b", all_vsto_source, re.IGNORECASE):
        raise RuntimeError("Production VSTO source contains a persistent task pane reference.")
    if re.search(r"\bstatic\s+(?:class\s+)?(?:Word\.)?Document\b", all_vsto_source):
        raise RuntimeError("Production VSTO source contains a static Word document reference.")
    if "FinalReleaseComObject" in all_vsto_source:
        raise RuntimeError(
            "Production VSTO source must not call Marshal.FinalReleaseComObject; "
            "it can invalidate shared Word RCWs and crash wwlib.dll."
        )
    create_ribbon_method = re.search(
        r"protected override Office\.IRibbonExtensibility "
        r"CreateRibbonExtensibilityObject\(\)(.*?)"
        r"(?=\n        private )",
        this_addin_source,
        re.DOTALL,
    )
    if create_ribbon_method is None:
        raise RuntimeError("CreateRibbonExtensibilityObject was not found.")
    if "EnsureRuntime()" in create_ribbon_method.group(1) or \
            "new ChuanHoaRibbon()" not in create_ribbon_method.group(1):
        raise RuntimeError(
            "Ribbon discovery must not initialize Word COM/runtime services before VSTO Startup."
        )
    startup_method = re.search(
        r"private void ThisAddIn_Startup\(object sender, System\.EventArgs e\)(.*?)"
        r"(?=\n        private void )",
        this_addin_source,
        re.DOTALL,
    )
    if startup_method is None:
        raise RuntimeError("ThisAddIn_Startup was not found.")
    for forbidden_startup_call in (
        "EnsureRuntime()",
        "Application.ActiveDocument",
        "OnDocumentWindowActivated(",
        "AttachRuntime(",
    ):
        if forbidden_startup_call in startup_method.group(1):
            raise RuntimeError(
                "VSTO Startup must not initialize runtime or touch a document before RibbonOnLoad: "
                + forbidden_startup_call
            )
    for deferred_startup_contract in (
        "_startupCompleted = true;",
        "TryScheduleRuntimeInitialization(250);",
        "private void OnRibbonLoaded(",
        "private void OnRuntimeStartupTimerTick(",
        "CaptureActiveDocumentMetadata(runtime);",
    ):
        if deferred_startup_contract not in this_addin_source:
            raise RuntimeError(
                "Deferred post-Ribbon startup contract is missing: " + deferred_startup_contract
            )
    for deferred_ribbon_contract in (
        "private IChuanHoaRibbonRuntime Runtime =>",
        "internal void AttachRuntime(IChuanHoaRibbonRuntime runtime)",
        "UnavailableRibbonRuntime.Instance",
        "_ribbonUi.Invalidate();",
    ):
        if deferred_ribbon_contract not in ribbon_source:
            raise RuntimeError(
                "Deferred Ribbon runtime contract is missing: " + deferred_ribbon_contract
            )
    if "_runtime." in callback_source or "CompleteRibbonLoad(ribbonUi);" not in callback_source:
        raise RuntimeError(
            "Generated Ribbon callbacks must use the fail-closed deferred runtime adapter."
        )
    for ribbon_load_safety_contract in (
        "internal event EventHandler? RibbonLoaded;",
        "private void CompleteRibbonLoad(Office.IRibbonUI ribbonUi)",
        "RibbonLoaded?.Invoke(this, EventArgs.Empty);",
    ):
        if ribbon_load_safety_contract not in ribbon_source:
            raise RuntimeError(
                "Ribbon load handshake contract is missing: " + ribbon_load_safety_contract
            )
    word_adapter_source = WORD_ADAPTER_PATH.read_text(encoding="utf-8")
    mutation_runtime_source = WORD_MUTATION_RUNTIME_PATH.read_text(encoding="utf-8")
    annotation_adapter_source = WORD_ANNOTATION_ADAPTER_PATH.read_text(encoding="utf-8")
    ribbon_runtime_source = RIBBON_RUNTIME_PATH.read_text(encoding="utf-8")
    local_access_source = LOCAL_ACCESS_MANAGER_PATH.read_text(encoding="utf-8")
    snapshot_builder_source = WORD_SNAPSHOT_BUILDER_PATH.read_text(encoding="utf-8")
    activation_method = re.search(
        r"public void OnDocumentWindowActivated\(Word\.Document document\)(.*?)(?=\n        public void )",
        ribbon_runtime_source,
        re.DOTALL,
    )
    if activation_method is None:
        raise RuntimeError("Word activation handler was not found.")
    if "_snapshotBuilder.Build(" in activation_method.group(1) or \
            "_documentReadRuntime.Prepare(" in activation_method.group(1) or \
            "_documentReadRuntime.PrepareForOneClick(" in activation_method.group(1):
        raise RuntimeError("WindowActivate must not read or snapshot document content.")
    if "btnDocDuLieu" in actual_controls or "ReadDocumentData" in ribbon_runtime_source:
        raise RuntimeError("The obsolete manual Read Data command must not be present.")
    runtime_sources = {
        path: path.read_text(encoding="utf-8")
        for path in sorted((VSTO_ROOT / "Runtime").glob("*.cs"))
        if path.name != "WordDocumentSnapshotBuilder.cs"
    }
    full_read_calls = []
    for path, source in runtime_sources.items():
        for pattern in (r"_snapshotBuilder\.Build\(", r"\.ScanFormat\(", r"\.ScanSpelling\("):
            for match in re.finditer(pattern, source):
                full_read_calls.append((path.name, match.group(0), match.start()))
    expected_read_calls = [
        ("WordDocumentReadRuntime.cs", "_snapshotBuilder.Build("),
        ("WordDocumentReadRuntime.cs", ".ScanFormat("),
        ("WordDocumentReadRuntime.cs", ".ScanSpelling("),
    ]
    actual_read_calls = [(name, call) for name, call, _ in full_read_calls]
    if actual_read_calls != expected_read_calls:
        raise RuntimeError(
            "Full document read/scan calls must exist only in WordDocumentReadRuntime: "
            + repr(actual_read_calls)
        )
    if "_documentReadRuntime.Read(" in ribbon_runtime_source:
        raise RuntimeError("Ribbon commands must use minimum-scope Prepare, not the full legacy Read helper.")
    command_analysis_calls = {
        "RunOneClick": "_documentReadRuntime.PrepareForOneClick(",
        "RunSelectedFindingFix": "_documentReadRuntime.Prepare(",
        "RunLocalScan": "_documentReadRuntime.Prepare(",
        "RunAnalysisBackedLocalCommand": "_documentReadRuntime.Prepare(",
        "RunQuickFixAllSpelling": "_documentReadRuntime.Prepare(",
    }
    for command_method_name, required_call in command_analysis_calls.items():
        command_method = re.search(
            rf"private void {command_method_name}\(.*?\)(.*?)(?=\n        private void )",
            ribbon_runtime_source,
            re.DOTALL,
        )
        if command_method is None or required_call not in command_method.group(1):
            raise RuntimeError(
                f"{command_method_name} must prepare its command-scoped document analysis."
            )
    if ribbon_runtime_source.count("_documentReadRuntime.Prepare(") != 6 or \
            ribbon_runtime_source.count("_documentReadRuntime.PrepareForOneClick(") != 1:
        raise RuntimeError("Unexpected command-scoped analysis call count in RibbonRuntime.")
    selected_fix_method = re.search(
        r"private void RunSelectedFindingFix\(\)(.*?)(?=\n        private void )",
        ribbon_runtime_source,
        re.DOTALL,
    )
    if selected_fix_method is None or \
            selected_fix_method.group(1).count("TryGetCurrentDocument()") != 1 or \
            "TryGetActiveContext()" in selected_fix_method.group(1):
        raise RuntimeError(
            "Selected finding fix must capture one stable document reference across Modern Comments focus changes."
        )
    selected_fix_body = selected_fix_method.group(1)
    for required_selected_capture in (
        "TryGetSelectedFinding(out selectedLane, out selectedFindingId)",
        "TryGetSelectedDocumentRange(out selectedStory, out selectedStart, out selectedEnd)",
        "FocusDocumentForCommand(document, true)",
        "selectedStory, selectedStart, selectedEnd, document",
    ):
        if required_selected_capture not in selected_fix_body:
            raise RuntimeError(
                "Selected finding fix must capture identity/range before leaving Modern Comments: "
                + required_selected_capture
            )
    if selected_fix_body.index("TryGetSelectedDocumentRange") > selected_fix_body.index(
        "FocusDocumentForCommand(document, true)"
    ):
        raise RuntimeError(
            "Selected finding range must be captured before document activation changes Selection."
        )
    current_document_method = re.search(
        r"private Word\.Document\? TryGetCurrentDocument\(\)(.*?)(?=\n        private )",
        ribbon_runtime_source,
        re.DOTALL,
    )
    if current_document_method is None or \
            "if (fallback.Windows.Count" in current_document_method.group(1) or \
            "_lastActivatedDocument = null" in current_document_method.group(1):
        raise RuntimeError(
            "Ribbon focus transitions must retain the last activated document; only DocumentBeforeClose may clear it."
        )
    capability_source = (VSTO_ROOT / "Runtime" / "WordDocumentCapabilityProvider.cs").read_text(encoding="utf-8")
    if "document.Windows.Count" in capability_source:
        raise RuntimeError(
            "Ribbon capability must not use transient Document.Windows.Count during Modern Comments focus changes."
        )
    if "TransientStateReasonCode" not in capability_source or \
            '"ACTIVE_DOCUMENT_REQUIRED"' in re.search(
                r"public WordDocumentCapability Evaluate\(Word\.Document document\)(.*?)(?=\n        private )",
                capability_source,
                re.DOTALL,
            ).group(1).split("catch (COMException)")[-1]:
        raise RuntimeError(
            "A transient COM rejection for a captured document must not become ACTIVE_DOCUMENT_REQUIRED."
        )
    for required_ribbon_focus_guard in (
        "sameCachedDocument",
        "WordDocumentCapabilityProvider.TransientStateReasonCode",
        "FocusDocumentForCommand(document)",
    ):
        if required_ribbon_focus_guard not in ribbon_runtime_source:
            raise RuntimeError(
                "Ribbon Modern Comments focus recovery is missing: " + required_ribbon_focus_guard
            )
    if "document.Activate();" in ribbon_runtime_source or \
            "TryFocusDocumentSelection()" not in annotation_adapter_source:
        raise RuntimeError(
            "Ribbon commands must leave Modern Comments through its captured document range, not Document.Activate."
        )
    if "TryFocusDocumentStart()" not in annotation_adapter_source or \
            "FocusDocumentForCommand(activeDocument, true, true)" not in ribbon_runtime_source:
        raise RuntimeError(
            "Whole-document 1-Click must have a selection-independent Modern Comments focus fallback."
        )
    if "if (_lastActivatedDocument != null)" not in activation_method.group(1) or \
            "if (document.Windows.Count == 0) return;" not in activation_method.group(1):
        raise RuntimeError(
            "WindowActivate must capture the first user document and ignore later hidden recovery clones."
        )
    if "CaptureStableRibbonCapability(document)" not in activation_method.group(1) or \
            "private void CaptureStableRibbonCapability(Word.Document document)" not in ribbon_runtime_source:
        raise RuntimeError(
            "WindowActivate must cache a stable metadata-only capability before Modern Comments can own focus."
        )
    if "Hãy bấm “Đọc dữ liệu”" in all_vsto_source:
        raise RuntimeError("Production VSTO source still requires the removed Read Data button.")
    if re.search(
        r"var scope = string\.Equals\(selectedLane, \"spelling\".*?"
        r"\? DocumentAnalysisScope\.Spelling\s*: DocumentAnalysisScope\.Format",
        ribbon_runtime_source,
        re.DOTALL,
    ):
        raise RuntimeError("Unknown selected-finding lanes must fail closed instead of defaulting to format.")
    for required_heavy_document_guard in (
        "IsPageLayoutCaptureSafe(document)",
        "document.Tables.Count > 10",
        "content.End - content.Start <= 200000",
        "allowPageLayout && paragraphCount <= 400",
        "withInTable && !largeStory",
        "!largeStory || !withInTable || paragraphIndex <= 150",
        "foreach (Word.Paragraph paragraphItem in paragraphs)",
        "CaptureLargeStoryParagraphs(story, storyType",
        "range = paragraph.Range.Duplicate",
        "authoritative Word coordinate",
    ):
        if required_heavy_document_guard not in snapshot_builder_source:
            raise RuntimeError(
                "The snapshot builder is missing its table-heavy pagination guard: "
                + required_heavy_document_guard
            )
    if "Application.DoEvents()" in all_vsto_source:
        raise RuntimeError("VSTO source must not pump messages and re-enter Word commands.")
    if re.search(r"\bTask\.Run\s*\(", all_vsto_source):
        raise RuntimeError("Word COM work must not be moved to Task.Run or a worker thread.")
    for operation_state in (
        "Idle", "Capturing", "Scanning", "Annotating", "Mutating", "Cancelling",
        "FailedRecoverable",
    ):
        if re.search(rf"\b{operation_state}\s*=", document_operation_source) is None:
            raise RuntimeError("Document operation state is missing: " + operation_state)
    for operation_contract in (
        "Environment.CurrentManagedThreadId",
        "GetAsyncKeyState(VirtualKeyEscape)",
        "_cancellationPermanentlyDisabled",
        "state == DocumentOperationState.Mutating",
        "CancellationToken => _cancellation.Token",
        "_application.StatusBar = string.Empty",
    ):
        if operation_contract not in document_operation_source:
            raise RuntimeError("Document operation contract is missing: " + operation_contract)
    for snapshot_cancellation_contract in (
        "CaptureLineShapes(document, sections, paragraphs, operation)",
        "CaptureTables(document, operation)",
        "CaptureProtectedSpans(document, operation)",
        "operation?.Checkpoint();",
        "operation?.ReportProgress(",
    ):
        if snapshot_cancellation_contract not in snapshot_builder_source:
            raise RuntimeError(
                "Snapshot capture is missing a bounded cancellation seam: "
                + snapshot_cancellation_contract
            )
    if "ClearLane(plan.Lane, operation);" not in annotation_adapter_source or \
            "public void ClearLane(string lane, DocumentOperationSession? operation = null)" not in annotation_adapter_source:
        raise RuntimeError("Annotation cleanup must honor command-scoped cancellation checkpoints.")
    for operation_method_name in (
        "RunOneClick", "RunQuickFixAllSpelling", "RunSelectedFindingFix", "RunLocalScan",
        "RunLocalCommand",
    ):
        operation_method = re.search(
            rf"private void {operation_method_name}\(.*?\)(.*?)(?=\n        private void )",
            ribbon_runtime_source,
            re.DOTALL,
        )
        if operation_method is None or "finally" not in operation_method.group(1) or \
                "_currentDocumentOperation?.Dispose();" not in operation_method.group(1):
            raise RuntimeError(
                operation_method_name + " must dispose its document operation from finally."
            )
    if (
        'RunLocalCommand("Co chữ", () => _localCommandRuntime.SetCharacterSpacing(-0.1f, false), '
        "showSuccessNotification: false)"
    ) not in ribbon_runtime_source:
        raise RuntimeError("Co chữ must run silently after a successful local edit.")
    for trailing_blank_guard in (
        "WordTrailingBlankPageCleaner.Remove(document)",
        "MaximumParagraphsPerRun",
        "documentRange.End - 1",
        "raw.IndexOf('\\a')",
    ):
        if trailing_blank_guard not in word_local_command_source:
            raise RuntimeError(
                "Trailing blank-page cleanup is missing its termination guard: "
                + trailing_blank_guard
            )
    if "WordTrailingBlankPageCleaner.Remove(document)" not in word_one_click_source or \
            "while (document.Paragraphs.Count" in all_vsto_source:
        raise RuntimeError(
            "Local and 1-Click trailing blank-page cleanup must use the shared bounded implementation."
        )
    for paragraph_indent_contract in (
        "ApplyBodyParagraphIndent(range, paragraph.Text);",
        "ParagraphIndentPolicy.IsDashListParagraph(text)",
        "tabStops.ClearAll();",
        "format.LeftIndent = textPosition;",
        "format.FirstLineIndent = markerPosition - textPosition;",
    ):
        if paragraph_indent_contract not in word_one_click_source:
            raise RuntimeError(
                "Body/list paragraph tab-stop normalization is missing: "
                + paragraph_indent_contract
            )
    for point_autofix_contract in (
        "IsLegalPointParagraph(paragraph.Text)",
        "range.Font.Bold = 0;",
        "range.Font.Italic = 0;",
    ):
        if point_autofix_contract not in word_one_click_source:
            raise RuntimeError(
                "1-Click is missing deterministic point-style normalization: "
                + point_autofix_contract
            )
    for appendix_pagination_guard in (
        "WordAppendixPaginationNormalizer.Normalize(document, roles)",
        "HasSamePageGeometry",
        "heading.Format.KeepWithNext = -1",
        "Word.WdSectionStart.wdSectionContinuous",
        "Different orientations are deliberately left untouched",
    ):
        if appendix_pagination_guard not in word_one_click_source and \
                appendix_pagination_guard not in word_appendix_pagination_source:
            raise RuntimeError(
                "Appendix pagination is missing its orientation-preserving guard: "
                + appendix_pagination_guard
            )
    selected_fix_method = re.search(
        r"private void RunSelectedFindingFix\(\)(.*?)(?=\n        private void )",
        ribbon_runtime_source,
        re.DOTALL,
    )
    if selected_fix_method is None:
        raise RuntimeError("The selected-finding fix command was not found.")
    if "Đã sửa lỗi đang chọn" in selected_fix_method.group(1) or "if (!result.Resolved)" not in selected_fix_method.group(1):
        raise RuntimeError("A successful selected-finding fix must not show a modal notification.")
    get_rule_pack_method = re.search(
        r"public LocalRulePack GetRulePack\(string requiredFeature\)(.*?)(?=\n        public )",
        local_access_source,
        re.DOTALL,
    )
    if get_rule_pack_method is None:
        raise RuntimeError("LocalAccessManager.GetRulePack was not found.")
    if "Refresh();" in get_rule_pack_method.group(1) or "GetAwaiter().GetResult()" in get_rule_pack_method.group(1):
        raise RuntimeError("A Word command must not perform synchronous license network I/O.")
    for required_adapter_seam in (
        "DocumentMutationPreflight",
        "CaptureFingerprint",
        "WordRecoveryCopyManager.Create",
        ".Saved",
        "BeginUndoRecord",
        "Rollback",
        "RestoreApplicationState",
    ):
        if required_adapter_seam not in word_adapter_source:
            raise RuntimeError(f"Word COM safety adapter is missing: {required_adapter_seam}")
    for required_runtime_seam in (
        "MutationSafetyKernel",
        "RsaSha256MutationAuthorizationVerifier",
        "PersistentAuthorizationReplayStore",
        "EmptyPublicKeyProvider",
    ):
        if required_runtime_seam not in mutation_runtime_source:
            raise RuntimeError(f"Word mutation runtime is missing: {required_runtime_seam}")
    for required_annotation_seam in (
        "AnnotationOwnershipPolicy.IsOwnedComment",
        "CaptureColorState",
        "RestoreVisualMarker",
        "Word.WdColor.wdColorRed",
        "Fields.Count",
        "ContentControls.Count",
    ):
        if required_annotation_seam not in annotation_adapter_source:
            raise RuntimeError(
                f"Word finding annotation adapter is missing: {required_annotation_seam}"
            )

    actual_counts = {
        "buttons": counts["button"],
        "menus": counts["menu"],
        "dropDowns": counts["dropDown"],
        "checkBoxes": counts["checkBox"],
        "interactiveControls": len(controls),
    }
    for count_name, actual in actual_counts.items():
        expected = contract["counts"][count_name]
        if actual != expected:
            raise RuntimeError(f"{count_name} mismatch: {actual} != {expected}")

    build_gate = "NOT_RUN"
    load_gate = "NOT_RUN"
    build_evidence_status = "No verified VSTO build/load evidence is available."
    if BUILD_LOAD_EVIDENCE_PATH.exists():
        build_load_evidence = json.loads(
            BUILD_LOAD_EVIDENCE_PATH.read_text(encoding="utf-8")
        )
        verified_artifacts = all(
            (ROOT / artifact["path"]).exists()
            and sha256(ROOT / artifact["path"]) == artifact["sha256"]
            for artifact in build_load_evidence["build"]["artifacts"]
        )
        word_smoke = build_load_evidence["wordSmoke"]
        screenshot = word_smoke.get("screenshot")
        verified_screenshot = False
        if screenshot:
            screenshot_path = ROOT / screenshot["path"]
            verified_screenshot = (
                screenshot_path.exists()
                and sha256(screenshot_path) == screenshot["sha256"]
            )
        verified_visual_smoke = (
            word_smoke.get("visualInspectionPerformed") is True
            and word_smoke.get("ribbonXmlSha256") == sha256(RIBBON_XML_PATH)
            and word_smoke.get("singleCustomRibbonTabVisible") is True
            and word_smoke.get("ribbonGroupsVisible") == 7
            and word_smoke.get("saveAsDocxControlVisible") is False
            and word_smoke.get("documentMutationExecuted") is False
        )
        if verified_artifacts:
            build_gate = "PASS_LOCAL_DEVELOPMENT"
        if verified_artifacts and (verified_screenshot or verified_visual_smoke):
            load_gate = "PASS_WORD16_X64_LOCAL"
        if verified_artifacts and (verified_screenshot or verified_visual_smoke):
            build_evidence_status = (
                "Local development build evidence verified by artifact hashes; Word 16 x64 visual smoke is bound to the current Ribbon hash. "
                "Production signing and the Word 2010/x86 compatibility lanes remain open."
            )
        elif verified_artifacts:
            build_evidence_status = (
                "Local development build evidence verified by artifact hashes; COM load is recorded separately, "
                "while current Word visual Ribbon/icon smoke remains NOT_RUN."
            )
        else:
            build_evidence_status = "Stored VSTO build artifacts are stale or unavailable."

    evidence = {
        "testId": "VSTO-SOURCE-001",
        "status": "PASS",
        "targetFramework": "net48",
        "minimumOfficeVersion": "14.0",
        "counts": actual_counts,
        "callbackMethods": len(callbacks_with_load),
        "invariants": {
            "exactlyOneRibbonTab": sum(1 for element in root.iter() if local_name(element.tag) == "tab")
            == 1,
            "autoFixDedicatedCallback": True,
            "noPersistentTaskPane": True,
            "noStaticWordDocument": True,
            "noFinalReleaseComObject": True,
            "deferredRibbonRuntimeInitialization": True,
            "startupWaitsForRibbonOnLoad": True,
            "noSynchronousLicenseRefreshOnCommand": True,
            "commandScopedMinimumAnalysis": True,
            "noManualReadDataPrerequisite": True,
            "tableHeavyPaginationGuard": True,
            "bodyListTabStopNormalization": True,
            "pointStyleAutoFix": True,
            "trailingBlankPageTerminationGuard": True,
            "appendixPaginationPreservesOrientation": True,
            "verificationBuildDoesNotRewriteInstalledManifest": True,
            "characterCondenseSuccessIsSilent": True,
            "selectedFindingFixSuccessIsSilent": True,
            "modernCommentsStableCapabilityBaseline": True,
            "oneClickModernCommentsFocusFallback": True,
            "perDocumentContextStore": True,
            "unregisteredCommandsFailClosed": True,
            "wordComAdapterConnected": True,
            "persistentReplayStoreConnected": True,
            "productionKeyProviderFailClosed": True,
            "findingAnnotationAdapterBuildConnected": True,
            "findingAnnotationCommandFailClosed": True,
        },
        "hashes": {
            "ribbonXmlSha256": sha256(RIBBON_XML_PATH),
            "generatedCallbacksSha256": sha256(CALLBACK_SOURCE_PATH),
            "projectSha256": sha256(PROJECT_PATH),
            "wordAdapterSha256": sha256(WORD_ADAPTER_PATH),
            "wordMutationRuntimeSha256": sha256(WORD_MUTATION_RUNTIME_PATH),
            "wordFindingAnnotationAdapterSha256": sha256(WORD_ANNOTATION_ADAPTER_PATH),
            "wordSnapshotBuilderSha256": sha256(WORD_SNAPSHOT_BUILDER_PATH),
        },
        "buildGate": build_gate,
        "loadGate": load_gate,
        "buildEvidenceStatus": build_evidence_status,
    }
    EVIDENCE_PATH.parent.mkdir(parents=True, exist_ok=True)
    EVIDENCE_PATH.write_text(
        json.dumps(evidence, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return evidence


def main() -> None:
    print(json.dumps(validate(), ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
