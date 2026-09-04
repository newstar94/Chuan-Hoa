using System;
using System.IO;
using System.Runtime.InteropServices;
using ChuanHoa.Client.Core.Safety;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class WordDocumentCapability
    {
        public WordDocumentCapability(
            bool hasActiveDocument,
            bool hasActiveWindow,
            bool isSupportedFormat,
            bool isSaved,
            bool isReadOnly,
            bool isProtected,
            bool trackChangesEnabled,
            bool hasSelection,
            bool selectionIsInTable,
            string reasonCode,
            string reason)
        {
            HasActiveDocument = hasActiveDocument;
            HasActiveWindow = hasActiveWindow;
            IsSupportedFormat = isSupportedFormat;
            IsSaved = isSaved;
            IsReadOnly = isReadOnly;
            IsProtected = isProtected;
            TrackChangesEnabled = trackChangesEnabled;
            HasSelection = hasSelection;
            SelectionIsInTable = selectionIsInTable;
            ReasonCode = reasonCode ?? string.Empty;
            Reason = reason ?? string.Empty;
        }

        public bool HasActiveDocument { get; }

        public bool HasActiveWindow { get; }

        public bool IsSupportedFormat { get; }

        public bool IsSaved { get; }

        public bool IsReadOnly { get; }

        public bool IsProtected { get; }

        public bool TrackChangesEnabled { get; }

        public bool HasSelection { get; }

        public bool SelectionIsInTable { get; }

        public string ReasonCode { get; }

        public string Reason { get; }

        public bool CanReadDocument =>
            HasActiveDocument && HasActiveWindow && IsSupportedFormat;

        public bool CanAnnotateDocument =>
            CanReadDocument && !IsReadOnly && !IsProtected;

        public bool CanMutateDocument =>
            CanAnnotateDocument && !TrackChangesEnabled;
    }

    public sealed class WordDocumentCapabilityProvider
    {
        public const string TransientStateReasonCode = "DOCUMENT_STATE_TEMPORARILY_UNAVAILABLE";
        private readonly Word.Application _application;

        public WordDocumentCapabilityProvider(Word.Application application)
        {
            _application = application ?? throw new ArgumentNullException("application");
        }

        public WordDocumentCapability Evaluate()
        {
            try
            {
                return Evaluate(_application.ActiveDocument);
            }
            catch (COMException)
            {
                return CreateBlocked(false, false, false, false, false, false, false, false, false,
                    "ACTIVE_DOCUMENT_REQUIRED", "Hãy mở một tài liệu Word trước khi sử dụng Chuẩn hóa.");
            }
        }

        public WordDocumentCapability Evaluate(Word.Document document)
        {
            if (document == null)
            {
                return CreateBlocked(false, false, false, false, false, false, false, false, false,
                    "ACTIVE_DOCUMENT_REQUIRED", "Hãy mở một tài liệu Word trước khi sử dụng Chuẩn hóa.");
            }

            try
            {
                // A captured Document reference remains authoritative while focus moves
                // between a Modern Comments card and the Ribbon. During that transition
                // Word can report Windows.Count == 0 or reject the call even though the
                // visible document is still open. Do not turn that transient UI state
                // into ACTIVE_DOCUMENT_REQUIRED and disable the entire Ribbon.
                var hasWindow = true;
                // Do not hide a rejected FullName/Path call by returning an empty
                // string. While focus is leaving Modern Comments, Word can reject one
                // of these properties for a few milliseconds. That is a busy document,
                // not a missing or unsupported document.
                var fullName = document.FullName ?? string.Empty;
                var isSaved = !string.IsNullOrWhiteSpace(document.Path);
                var saveFormat = (int)document.SaveFormat;
                var extension = Path.GetExtension(fullName);
                var supported = isSaved
                    ? SupportedWordDocumentFormatPolicy.IsSupported(fullName, saveFormat)
                    : (saveFormat == 0 || saveFormat == 12 || saveFormat == 16 || saveFormat == 24 || string.IsNullOrEmpty(extension));
                var readOnly = document.ReadOnly;
                var protectedDocument = document.ProtectionType != Word.WdProtectionType.wdNoProtection;
                var trackChanges = document.TrackRevisions;
                // Selection state is intentionally not queried during Ribbon capability
                // evaluation. Office calls getEnabled in bursts, and touching Selection for
                // every burst can block while Word is repaginating or completing a save.
                // Commands that require a selection validate it only after the user clicks.
                var hasSelection = false;
                var selectionIsInTable = false;

                if (!hasWindow)
                {
                    return CreateBlocked(true, false, supported, isSaved, readOnly, protectedDocument,
                        trackChanges, hasSelection, selectionIsInTable,
                        "ACTIVE_WINDOW_REQUIRED", "Cần một cửa sổ tài liệu Word đang hoạt động.");
                }

                if (!supported)
                {
                    return CreateBlocked(true, true, false, isSaved, readOnly, protectedDocument,
                        trackChanges, hasSelection, selectionIsInTable,
                        "DOCUMENT_FORMAT_UNSUPPORTED", "Chuẩn hóa chỉ xử lý trực tiếp tài liệu .doc và .docx.");
                }

                if (readOnly)
                {
                    return CreateBlocked(true, true, true, isSaved, true, protectedDocument,
                        trackChanges, hasSelection, selectionIsInTable,
                        "DOCUMENT_READ_ONLY", "Tài liệu đang ở chế độ chỉ đọc; có thể quét nhưng không thể ghi chú hoặc sửa.");
                }

                if (protectedDocument)
                {
                    return CreateBlocked(true, true, true, isSaved, false, true,
                        trackChanges, hasSelection, selectionIsInTable,
                        "DOCUMENT_PROTECTED", "Tài liệu đang được bảo vệ; có thể quét nhưng không thể ghi chú hoặc sửa.");
                }

                if (trackChanges)
                {
                    return CreateBlocked(true, true, true, isSaved, false, false,
                        true, hasSelection, selectionIsInTable,
                        "TRACK_CHANGES_ENABLED", "Đang bật Theo dõi thay đổi; lệnh sửa tài liệu sẽ bị khóa an toàn.");
                }

                return new WordDocumentCapability(
                    true, true, true, isSaved, false, false, false,
                    hasSelection, selectionIsInTable, "READY", "Tài liệu sẵn sàng.");
            }
            catch (COMException)
            {
                return CreateBlocked(true, true, false, false, false, false, false, false, false,
                    TransientStateReasonCode,
                    "Word đang hoàn tất chuyển tiêu điểm từ khung comment. Hãy thử lại thao tác.");
            }
        }

        private static WordDocumentCapability CreateBlocked(
            bool hasActiveDocument,
            bool hasActiveWindow,
            bool isSupportedFormat,
            bool isSaved,
            bool isReadOnly,
            bool isProtected,
            bool trackChangesEnabled,
            bool hasSelection,
            bool selectionIsInTable,
            string reasonCode,
            string reason)
        {
            return new WordDocumentCapability(
                hasActiveDocument,
                hasActiveWindow,
                isSupportedFormat,
                isSaved,
                isReadOnly,
                isProtected,
                trackChangesEnabled,
                hasSelection,
                selectionIsInTable,
                reasonCode,
                reason);
        }
    }
}
