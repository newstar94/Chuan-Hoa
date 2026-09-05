using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Deployment.Application;
using ChuanHoa.AddIn.Vsto.Ribbon;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;
using ChuanHoa.Client.Core.Lexicon;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class RibbonRuntime : IChuanHoaRibbonRuntime, IDisposable
    {
        private static readonly IReadOnlyList<string> DocumentTypeCodes = new[]
        {
            LocalDocumentTypeCodes.Unknown,
            LocalDocumentTypeCodes.OfficialLetter,
            LocalDocumentTypeCodes.Resolution,
            LocalDocumentTypeCodes.Decision,
            LocalDocumentTypeCodes.Directive,
            LocalDocumentTypeCodes.Circular,
            LocalDocumentTypeCodes.Communique,
            LocalDocumentTypeCodes.Notice,
            LocalDocumentTypeCodes.Guidance,
            LocalDocumentTypeCodes.Program,
            LocalDocumentTypeCodes.Plan,
            LocalDocumentTypeCodes.Option,
            LocalDocumentTypeCodes.Scheme,
            LocalDocumentTypeCodes.Project,
            LocalDocumentTypeCodes.Report,
            LocalDocumentTypeCodes.Proposal,
            LocalDocumentTypeCodes.Regulation,
            LocalDocumentTypeCodes.Invitation,
            LocalDocumentTypeCodes.Telegram,
            LocalDocumentTypeCodes.IntroductionLetter,
            LocalDocumentTypeCodes.Minutes,
            LocalDocumentTypeCodes.LeavePermit,
            LocalDocumentTypeCodes.AuthorizationLetter,
            LocalDocumentTypeCodes.SendingSlip,
            LocalDocumentTypeCodes.TransferSlip,
            LocalDocumentTypeCodes.NotificationSlip,
            LocalDocumentTypeCodes.Conclusion
        };

        private readonly Word.Application _application;
        private readonly DocumentContextStore _contextStore;
        private readonly WordMutationRuntime _wordMutationRuntime;
        private readonly WordDocumentReadRuntime _documentReadRuntime;
        private readonly LocalAccessManager _localAccessManager;
        private readonly WordLocalScanRuntime _localScanRuntime;
        private readonly WordLocalCommandRuntime _localCommandRuntime;
        private readonly WordOneClickRuntime _oneClickRuntime;
        private readonly RibbonImageProvider _imageProvider = new RibbonImageProvider();
        private readonly Dictionary<string, Action> _buttonCommands;
        private readonly System.Windows.Forms.Timer _accessRefreshTimer;
        private Office.IRibbonUI? _ribbonUi;
        private Word.Document? _lastActivatedDocument;
        private DocumentContext? _lastUserContext;
        private int _accessRefreshCompleted;
        private int _documentOperationInProgress;
        private DocumentOperationSession? _currentDocumentOperation;
        private bool _disposed;

        public RibbonRuntime(
            Word.Application application,
            DocumentContextStore contextStore,
            WordMutationRuntime wordMutationRuntime)
        {
            _application = application ?? throw new ArgumentNullException("application");
            _contextStore = contextStore ?? throw new ArgumentNullException("contextStore");
            _wordMutationRuntime = wordMutationRuntime ?? throw new ArgumentNullException("wordMutationRuntime");
            var assemblyVersion = typeof(RibbonRuntime).Assembly.GetName().Version;
            _localAccessManager = new LocalAccessManager(assemblyVersion == null ? "1.0.0.0" : assemblyVersion.ToString());
            _documentReadRuntime = new WordDocumentReadRuntime(_application, _localAccessManager);
            _localScanRuntime = new WordLocalScanRuntime(_application, _localAccessManager);
            _localCommandRuntime = new WordLocalCommandRuntime(_application, _localAccessManager,
                TryGetCurrentDocument, TryGetCurrentDocumentFingerprint);
            _oneClickRuntime = new WordOneClickRuntime(_application, _localAccessManager);
            _localAccessManager.CacheStateChanged += OnCacheStateChanged;
            PersonalDictionaryManager.Instance.Changed += OnPersonalDictionaryChanged;
            _accessRefreshTimer = new System.Windows.Forms.Timer { Interval = 500 };
            _accessRefreshTimer.Tick += OnAccessRefreshTimerTick;
            _buttonCommands = new Dictionary<string, Action>(StringComparer.Ordinal)
            {
                { "btnAutoFixAll2026", RunOneClick },
                { "btnKiemTra", RunFormatScan },
                { "btnKiemTraChinhTa", RunSpellingScan },
                { "btnSuaTatCaChinhTa", RunQuickFixAllSpelling },
                { "btnSuaLoiDangChon", RunSelectedFindingFix },
                { "btnChuyenDoiUnicode", () => RunLocalCommand("Chuyển đổi Unicode", _localCommandRuntime.ConvertLegacyEncodingToUnicode) },
                { "btnDinhDangTrangGiay", () => RunLocalCommand("Định dạng trang giấy", _localCommandRuntime.FormatPage) },
                { "btnChenTrangNgang", () => RunLocalCommand("Chèn trang ngang", () => _localCommandRuntime.InsertSection(true)) },
                { "btnChenTrangDoc", () => RunLocalCommand("Chèn trang dọc", () => _localCommandRuntime.InsertSection(false)) },
                { "btnXoaTrangThua", () => RunLocalCommand("Xóa trang thừa", _localCommandRuntime.RemoveTrailingBlankParagraphs) },
                { "btnDungBoStyleCo15", () => RunAnalysisBackedLocalCommand("Dựng bộ Style cỡ 15", context => _localCommandRuntime.BuildStyleSet(context, 15f)) },
                { "btnDungBoStyleCo14", () => RunAnalysisBackedLocalCommand("Dựng bộ Style cỡ 14", context => _localCommandRuntime.BuildStyleSet(context, 14f)) },
                { "btnDungBoStyleCo13", () => RunAnalysisBackedLocalCommand("Dựng bộ Style cỡ 13", context => _localCommandRuntime.BuildStyleSet(context, 13f)) },
                { "btnCoChu15", () => RunAnalysisBackedLocalCommand("Bộ cỡ chữ 15", context => _localCommandRuntime.ApplyFontSizeSet(context, 15f)) },
                { "btnCoChu14", () => RunAnalysisBackedLocalCommand("Bộ cỡ chữ 14", context => _localCommandRuntime.ApplyFontSizeSet(context, 14f)) },
                { "btnCoChu13", () => RunAnalysisBackedLocalCommand("Bộ cỡ chữ 13", context => _localCommandRuntime.ApplyFontSizeSet(context, 13f)) },
                { "btnKeepWithNext", () => RunLocalCommand("Keep with next", _localCommandRuntime.KeepWithNext) },
                { "btnChenSoTrang", () => RunLocalCommand("Chèn số trang", _localCommandRuntime.InsertPageNumbers) },
                { "btnCoChu", () => RunLocalCommand("Co chữ", () => _localCommandRuntime.SetCharacterSpacing(-0.1f, false), showSuccessNotification: false) },
                { "btnGianChuNormal", () => RunLocalCommand("Giãn chữ bình thường", () => _localCommandRuntime.SetCharacterSpacing(0f, true)) },
                { "btnGianChuRa", () => RunLocalCommand("Giãn chữ", () => _localCommandRuntime.SetCharacterSpacing(0.1f, false)) },
                { "btnLapDongTieuDe", () => RunLocalCommand("Lặp tiêu đề bảng", _localCommandRuntime.RepeatTableHeaders) },
                { "btnChuanHoaBang", () => RunLocalCommand("Căn giữa bảng", _localCommandRuntime.CenterTables) },
                { "btnChuanHoaAnh", () => RunLocalCommand("Căn giữa ảnh", _localCommandRuntime.CenterImages) },
                { "btnCanDinhO", () => RunLocalCommand("Căn đỉnh ô", () => _localCommandRuntime.AlignCurrentCells(false)) },
                { "btnCanGiuaO", () => RunLocalCommand("Căn giữa ô", () => _localCommandRuntime.AlignCurrentCells(true)) },
                { "btnXoaKyTuThuaBangExcel", () => RunLocalCommand("Xóa ký tự thừa", _localCommandRuntime.CleanExcelTableCharacters) },
                { "btnKieuOaUy", () => RunLocalCommand("Kiểu oà, uý", () => _localCommandRuntime.NormalizeTonePlacement(VietnameseTonePlacementStyle.MainVowel)) },
                { "btnKieuOaUy2", () => RunLocalCommand("Kiểu òa, úy", () => _localCommandRuntime.NormalizeTonePlacement(VietnameseTonePlacementStyle.FirstVowel)) },
                { "btnDoiDauThapPhan", () => RunLocalCommand("Dấu phẩy thập phân", _localCommandRuntime.ConvertDecimalSeparators) },
                { "btnTuDienCaNhan", () => _localCommandRuntime.OpenCustomDictionaryDialog() },
                { "btnKiemTraPhienBanMoi", ShowUpdateStatus },
                { "btnGuiPhanHoi", OpenFeedback },
                { "btnGioiThieu", ShowAbout }
            };
        }

        public void AttachRibbon(Office.IRibbonUI ribbonUi)
        {
            ThrowIfDisposed();
            _ribbonUi = ribbonUi ?? throw new ArgumentNullException("ribbonUi");
        }

        public bool IsEnabled(string controlId)
        {
            ThrowIfDisposed();
            if (Volatile.Read(ref _documentOperationInProgress) != 0)
            {
                return false;
            }
            // Office invokes getEnabled repeatedly while opening a document, moving
            // focus and repainting the Ribbon. This callback must therefore be a pure
            // in-memory lookup: it must not touch ActiveDocument, Selection, a lease
            // file, the network or any Word document property. Each explicit command
            // performs its own capability and license checks after the user clicks.
            return _buttonCommands.ContainsKey(controlId) ||
                string.Equals(controlId, "ddQuyDinh", StringComparison.Ordinal) ||
                string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal) ||
                string.Equals(controlId, "mnuBoDau", StringComparison.Ordinal) ||
                string.Equals(controlId, "mnuDungBoStyle", StringComparison.Ordinal) ||
                string.Equals(controlId, "mnuThongTinTienIch", StringComparison.Ordinal);
        }

        public int GetSelectedItemIndex(string controlId)
        {
            ThrowIfDisposed();
            // getSelectedItemIndex is also called by Office without a user action.
            // Return only the last user-established selection; never resolve a Word
            // document from this callback.
            return _lastUserContext == null ? 0 : _lastUserContext.GetSelection(controlId);
        }

        public void SelectDropDownItem(string controlId, string selectedId, int selectedIndex)
        {
            ThrowIfDisposed();
            var document = TryGetCurrentDocument();
            var context = document == null ? null : _contextStore.GetOrCreate(document);
            if (context == null)
            {
                return;
            }
            _lastUserContext = context;

            context.SetSelection(controlId, selectedIndex);
            if (string.Equals(controlId, "ddQuyDinh", StringComparison.Ordinal))
            {
                context.RegimeCode = MapRegime(selectedId);
                context.RegimeWasSelectedManually = true;
                context.DocumentTypeCode = "UNKNOWN";
                context.DocumentTypeWasSelectedManually = false;
                context.SetSelection("ddLoaiVanBan", 0);
                context.ClearReadAnalysis();
                Invalidate("ddLoaiVanBan");
            }
            else if (string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal))
            {
                context.DocumentTypeCode = MapDocumentType(selectedIndex);
                context.DocumentTypeWasSelectedManually = true;
                context.ClearReadAnalysis();
            }

            Invalidate(controlId);
        }

        public int GetItemCount(string controlId)
        {
            ThrowIfDisposed();
            return string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal)
                ? DocumentTypeItemCount
                : 0;
        }

        public string GetItemLabel(string controlId, int index)
        {
            ThrowIfDisposed();
            if (!string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal))
            {
                return string.Empty;
            }

            return GetDocumentTypeItemLabel(index);
        }

        internal static int DocumentTypeItemCount => DocumentTypeCodes.Count;

        internal static string GetDocumentTypeItemLabel(int index)
        {
            return index < 0 || index >= DocumentTypeCodes.Count
                ? string.Empty
                : LocalDocumentTypeCodes.GetDisplayName(DocumentTypeCodes[index]);
        }

        public void ExecuteButton(string controlId)
        {
            ThrowIfDisposed();
            Action command;
            if (!_buttonCommands.TryGetValue(controlId, out command))
            {
                FailClosed(controlId);
                return;
            }

            command();
            // GetRulePack starts a background refresh only when the signed local
            // cache is unavailable. Poll the completion flag only for that bounded
            // refresh; a permanently running UI timer would keep waking Word after
            // the user has stopped interacting with the add-in.
            if (_localAccessManager.IsRefreshInProgress ||
                Volatile.Read(ref _accessRefreshCompleted) != 0)
                _accessRefreshTimer.Start();
        }

        public object GetImage(string controlId)
        {
            if (_disposed) return null!;
            try { return _imageProvider.GetImage(controlId); }
            catch (Exception exception)
            {
                Trace.TraceError("ChuanHoa Ribbon image callback failed: {0}", exception);
                return null!;
            }
        }

        public void OnDocumentBeforeClose(Word.Document document)
        {
            if (_disposed || document == null)
            {
                return;
            }

            // Document-scoped dictionary ignores are session-only. Remove the exact
            // stable scope while the COM identity is still valid, then discard the
            // document context so the cache cannot grow for the lifetime of Word.
            var closingContext = _contextStore.GetOrCreate(document);
            PersonalDictionaryManager.Instance.ClearDocumentIgnores(closingContext.DictionaryScopeId);
            _contextStore.Remove(document);
            if (IsSameComDocument(_lastActivatedDocument, document))
                _lastActivatedDocument = null;
            if (ReferenceEquals(_lastUserContext, closingContext))
                _lastUserContext = null;
            InvalidateAll();
        }

        public void OnDocumentBeforeSave(Word.Document document)
        {
            if (_disposed || document == null) return;
            _contextStore.GetOrCreate(document).ClearReadAnalysis();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _ribbonUi = null;
            _lastActivatedDocument = null;
            _lastUserContext = null;
            _localAccessManager.CacheStateChanged -= OnCacheStateChanged;
            PersonalDictionaryManager.Instance.Changed -= OnPersonalDictionaryChanged;
            _imageProvider.Dispose();
            _accessRefreshTimer.Stop();
            _accessRefreshTimer.Tick -= OnAccessRefreshTimerTick;
            _accessRefreshTimer.Dispose();
            _localAccessManager.Dispose();
            _disposed = true;
        }

        private void OnCacheStateChanged(object? sender, EventArgs eventArgs)
        {
            // This callback is raised by a worker thread. The WinForms timer below marshals
            // the Ribbon invalidation back to Word's UI thread.
            Interlocked.Exchange(ref _accessRefreshCompleted, 1);
        }

        private void OnPersonalDictionaryChanged(object? sender, EventArgs eventArgs)
        {
            // Dictionary changes are raised by the modal dialog on Word's UI thread.
            // Cached spelling results must never outlive the dictionary that produced them.
            _contextStore.ClearAllReadAnalysis();
            InvalidateAll();
        }

        private void OnAccessRefreshTimerTick(object? sender, EventArgs eventArgs)
        {
            if (_disposed)
            {
                _accessRefreshTimer.Stop();
                return;
            }
            if (Interlocked.Exchange(ref _accessRefreshCompleted, 0) == 0)
            {
                if (!_localAccessManager.IsRefreshInProgress)
                    _accessRefreshTimer.Stop();
                return;
            }
            // The refreshed lease can unlock every DOCUMENT_TOOLS command, not only
            // the scanners and 1-Click. Refresh the complete Ribbon so newly ported
            // style, Unicode, font-size and tone commands do not remain grey.
            _accessRefreshTimer.Stop();
            InvalidateAll();
        }

        private DocumentContext? TryGetActiveContext()
        {
            try
            {
                var document = TryGetCurrentDocument();
                return document == null ? null : _contextStore.GetOrCreate(document);
            }
            catch (COMException)
            {
                return null;
            }
        }

        private string? TryGetCurrentDocumentFingerprint()
        {
            var context = TryGetActiveContext();
            return context == null ? null : context.DictionaryScopeId;
        }

        private Word.Document? TryGetCurrentDocument()
        {
            try
            {
                var active = _application.ActiveDocument;
                if (active != null)
                {
                    _lastActivatedDocument = active;
                    return active;
                }
            }
            catch (COMException)
            {
            }

            // CaptureStableRibbonCapability holds the same user document through a
            // Modern Comments focus transition. It is an independent safe fallback
            // when Word rejects ActiveDocument and a transient close event has cleared
            // the activation pointer. DocumentBeforeClose clears both references for
            // the exact COM document, so a genuinely closed document is never reused.
            var fallback = _lastActivatedDocument;
            // Do not probe fallback.Windows.Count here. Modern Comments temporarily
            // reports zero windows while focus moves to the Ribbon, even though the
            // document remains open. DocumentBeforeClose is the authoritative event
            // that clears this reference. In particular, a Ribbon getEnabled callback
            // must never destroy the state needed by the command it is enabling.
            return fallback;
        }

        private static bool IsSameComDocument(Word.Document? left, Word.Document right)
        {
            if (left == null) return false;
            if (ReferenceEquals(left, right)) return true;
            IntPtr leftIdentity = IntPtr.Zero;
            IntPtr rightIdentity = IntPtr.Zero;
            try
            {
                leftIdentity = Marshal.GetIUnknownForObject(left);
                rightIdentity = Marshal.GetIUnknownForObject(right);
                return leftIdentity == rightIdentity;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
                if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
            }
        }

        private void FocusDocumentForCommand(Word.Document document, bool collapseDocumentSelection = false,
            bool allowDocumentStartFallback = false)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var annotations = new WordFindingAnnotationAdapter(_application, document);
            // Document.Activate itself is rejected while Modern Comments owns focus.
            // Selecting the comment's document scope is Word's supported focus hand-off
            // and does not read or scan any unrelated document content.
            if (!annotations.TryFocusDocumentSelection(collapseDocumentSelection) &&
                (!allowDocumentStartFallback || !annotations.TryFocusDocumentStart()))
                throw new InvalidOperationException(
                    "Word đang hoàn tất chuyển tiêu điểm từ khung comment. Hãy thử lại thao tác.");
            _lastActivatedDocument = document;
        }

        private DocumentContext RequireActiveContext()
        {
            var context = TryGetActiveContext();
            if (context == null) throw new InvalidOperationException("Hãy mở một tài liệu Word.");
            return context;
        }

        private static void OpenFeedback()
        {
            const string url = "https://ngoctien.id.vn/chuan-hoa-the-thuc/feedback";
            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "Không thể mở trang phản hồi.\n\n" + exception.Message,
                    "Gửi phản hồi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                "Chuẩn hóa cho Microsoft Word\nNền tảng VSTO, hỗ trợ mục tiêu Word 2010 trở lên.\n\n" +
                _localAccessManager.DescribeStatus(),
                "Thông tin Chuẩn hóa",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ShowUpdateStatus()
        {
            var assemblyVersion = typeof(RibbonRuntime).Assembly.GetName().Version;
            if (!ApplicationDeployment.IsNetworkDeployed)
            {
                MessageBox.Show(
                    "Phiên bản assembly: " + assemblyVersion + "\n" +
                    "Bản đang chạy được nạp trực tiếp từ máy phát triển, nên Word sẽ nhận code mới sau khi đóng hoàn toàn và mở lại.\n" +
                    "Kênh cập nhật ClickOnce chưa hoạt động trong phiên cài đặt này.",
                    "Cập nhật Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                var deployment = ApplicationDeployment.CurrentDeployment;
                var update = deployment.CheckForDetailedUpdate(false);
                var currentVersion = deployment.CurrentVersion;
                var message = update.UpdateAvailable
                    ? "Có phiên bản mới " + update.AvailableVersion + ". Word sẽ nhận bản cập nhật theo chính sách ClickOnce đã cài đặt."
                    : "Bạn đang dùng phiên bản mới nhất trên kênh hiện tại.";
                MessageBox.Show(
                    "Phiên bản hiện tại: " + currentVersion + "\n" + message,
                    "Cập nhật Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (DeploymentDownloadException exception)
            {
                MessageBox.Show(
                    "Không thể kiểm tra cập nhật lúc này.\n\n" + exception.Message,
                    "Cập nhật Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (InvalidDeploymentException exception)
            {
                MessageBox.Show(
                    "Cấu hình cập nhật hiện tại không hợp lệ.\n\n" + exception.Message,
                    "Cập nhật Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void RunFormatScan()
        {
            RunLocalScan(false);
        }

        private void RunOneClick()
        {
            var context = TryGetActiveContext();
            if (context == null)
            {
                MessageBox.Show("Hãy mở một tài liệu Word.", "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var regime = string.Equals(context.RegimeCode, "PARTY_HD05", StringComparison.OrdinalIgnoreCase)
                ? "Hướng dẫn 05 (văn bản Đảng)"
                : "Nghị định 30 (văn bản hành chính)";
            var answer = MessageBox.Show(
                "Chuẩn hóa toàn bộ tài liệu theo " + regime + "?\n\n" +
                "Ứng dụng sẽ tự đọc trạng thái hiện tại và nhận diện loại văn bản.\n\n" +
                "Chương trình sẽ tạo bản sao khôi phục, sửa định dạng, bổ sung Line Shape và tự sửa các lỗi chính tả có phương án xác định. Lỗi cần cân nhắc vẫn được comment và tô đỏ.",
                "Chuẩn hóa toàn bộ", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (answer != DialogResult.OK) return;

            if (Interlocked.CompareExchange(ref _documentOperationInProgress, 1, 0) != 0)
            {
                MessageBox.Show("Ứng dụng đang xử lý tài liệu. Hãy chờ thao tác hiện tại hoàn tất.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                _currentDocumentOperation = new DocumentOperationSession(
                    _application, "Chuẩn hóa toàn bộ", DocumentOperationState.Capturing);
                var activeDocument = TryGetCurrentDocument();
                if (activeDocument == null) throw new InvalidOperationException("Hãy mở một tài liệu Word.");
                // 1-Click operates on the whole document and does not depend on the
                // current text selection. If Modern Comments cannot map its transient
                // selection back to a comment scope, move to a safe collapsed range at
                // the start of this same document instead of aborting the command.
                FocusDocumentForCommand(activeDocument, true, true);
                // Resolve the context again after the confirmation dialog so this
                // command cannot carry state from a previously active document.
                context = _contextStore.GetOrCreate(activeDocument);
                _documentReadRuntime.PrepareForOneClick(context, activeDocument, _currentDocumentOperation);
                var result = _oneClickRuntime.Execute(context, activeDocument, _currentDocumentOperation);
                // Execute has already captured and rescanned the mutated document.
                // Reuse that exact post-fix evidence for annotation instead of doing a
                // second full COM capture/scan on large or table-heavy documents.
                var remainingFormat = _localScanRuntime.ScanAndAnnotate(context, false, activeDocument,
                    _currentDocumentOperation);
                var remainingSpelling = _localScanRuntime.ScanAndAnnotate(context, true, activeDocument,
                    _currentDocumentOperation);
                var remainingFindingCount = remainingFormat.Findings.Count + remainingSpelling.Findings.Count;
                SynchronizeDocumentTypeSelection(context);
                MessageBox.Show(
                    "Đã chuẩn hóa toàn bộ tại máy.\n\n" +
                    "Loại văn bản: " + LocalDocumentTypeCodes.GetDisplayName(context.DocumentTypeCode) +
                    ".\nĐoạn đã định dạng: " + result.ChangedParagraphs +
                    "\nĐường kẻ đã xử lý: " + result.InsertedLines +
                    "\nSection đã chuẩn hóa: " + result.NormalizedSections +
                    "\nBảng đã xử lý: " + result.NormalizedTables +
                    "\nLỗi chính tả đã tự sửa: " + result.CorrectedSpellingItems +
                    "\nVấn đề còn lại sau khi sửa: " + remainingFindingCount +
                    "\n\nCác lỗi còn lại đã được comment và tô đỏ theo kết quả kiểm tra mới." +
                    "\n\nBản sao khôi phục:\n" + result.BackupPath,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                context.ClearReadAnalysis();
                MessageBox.Show("Đã hủy trước khi chỉnh sửa tài liệu. Tài liệu chưa bị thay đổi.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                MessageBox.Show("Không thể chuẩn hóa toàn bộ.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _currentDocumentOperation?.Dispose();
                _currentDocumentOperation = null;
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateAll();
            }
        }

        private void RunSpellingScan()
        {
            RunLocalScan(true);
        }

        private void RunQuickFixAllSpelling()
        {
            var document = TryGetCurrentDocument();
            if (document == null)
            {
                MessageBox.Show("Hãy mở một tài liệu Word.", "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Interlocked.CompareExchange(ref _documentOperationInProgress, 1, 0) != 0)
            {
                MessageBox.Show("Ứng dụng đang xử lý tài liệu. Hãy chờ thao tác hiện tại hoàn tất.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                _currentDocumentOperation = new DocumentOperationSession(
                    _application, "Sửa nhanh chính tả", DocumentOperationState.Capturing);
                FocusDocumentForCommand(document);
                var context = _contextStore.GetOrCreate(document);
                _documentReadRuntime.Prepare(context, DocumentAnalysisScope.Spelling, document, false,
                    _currentDocumentOperation);
                _localScanRuntime.ScanAndAnnotate(context, true, document, _currentDocumentOperation);

                var fixedCount = _oneClickRuntime.FixAllSpellingFindings(context, document,
                    _currentDocumentOperation);

                _documentReadRuntime.Prepare(context, DocumentAnalysisScope.Spelling, document, false,
                    _currentDocumentOperation);
                var remainingScan = _localScanRuntime.ScanAndAnnotate(context, true, document,
                    _currentDocumentOperation);

                MessageBox.Show(
                    "Đã sửa nhanh chính tả tại máy.\n\n" +
                    "Số lỗi đã tự động sửa thành công: " + fixedCount + "\n" +
                    "Số lỗi còn lại cần người dùng cân nhắc: " + remainingScan.Findings.Count + "\n\n" +
                    "Bạn có thể hoàn tác toàn bộ bằng Ctrl+Z nếu cần.",
                    "Sửa nhanh chính tả",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                _contextStore.GetOrCreate(document).ClearReadAnalysis();
                MessageBox.Show("Đã hủy trước khi chỉnh sửa tài liệu. Tài liệu chưa bị thay đổi.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                MessageBox.Show("Không thể sửa nhanh chính tả.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _currentDocumentOperation?.Dispose();
                _currentDocumentOperation = null;
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateAll();
            }
        }

        private void RunSelectedFindingFix()
        {
            // Resolve the document exactly once. When focus leaves a Modern Comments
            // card for the Ribbon, Word can briefly reject a second ActiveDocument
            // request even though the document remains open.
            var document = TryGetCurrentDocument();
            if (document == null)
            {
                MessageBox.Show("Hãy mở một tài liệu Word.", "Chuẩn hóa",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string selectedLane;
            string selectedFindingId;
            string selectedStory;
            int selectedStart;
            int selectedEnd;
            try
            {
                // Capture every selection-dependent value before activating the
                // document window. Activation intentionally leaves the Modern Comment
                // card, so no later stage may query Application.Selection again.
                var annotations = new WordFindingAnnotationAdapter(_application, document);
                if (!annotations.TryGetSelectedFinding(out selectedLane, out selectedFindingId) ||
                    !annotations.TryGetSelectedDocumentRange(out selectedStory, out selectedStart, out selectedEnd))
                    throw new InvalidOperationException(
                        "Hãy bấm vào phần văn bản đang được comment/tô đỏ bởi Chuẩn hóa rồi bấm lại Sửa lỗi đang chọn.");
            }
            catch (COMException)
            {
                MessageBox.Show("Word đang chuyển tiêu điểm từ khung comment. Hãy bấm lại Sửa lỗi đang chọn.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            catch (InvalidOperationException exception)
            {
                MessageBox.Show("Không thể sửa lỗi đang chọn.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (Interlocked.CompareExchange(ref _documentOperationInProgress, 1, 0) != 0)
            {
                MessageBox.Show("Ứng dụng đang xử lý tài liệu. Hãy chờ thao tác hiện tại hoàn tất.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Every selection-dependent value and the matching finding were
                // captured above from the user's explicit check. Starting this
                // targeted command in Capturing state incorrectly tells Word that
                // the add-in is reading the whole document, even though no snapshot
                // or scanner is involved. Enter the mutation phase immediately.
                _currentDocumentOperation = new DocumentOperationSession(
                    _application, "Sửa lỗi đang chọn", DocumentOperationState.Mutating);
                var supportedLane = string.Equals(selectedLane, "spelling", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(selectedLane, "format", StringComparison.OrdinalIgnoreCase);
                if (!supportedLane)
                {
                    throw new InvalidOperationException(
                        "Comment đang chọn không thuộc nhóm lỗi Chuẩn hóa hỗ trợ.");
                }
                FocusDocumentForCommand(document, true);
                var context = _contextStore.GetOrCreate(document);
                var scan = string.Equals(selectedLane, "spelling", StringComparison.OrdinalIgnoreCase)
                    ? context.LastSpellingScan
                    : context.LastFormatScan;
                // A selected repair is deliberately targeted. It must use the finding
                // already produced by the user's explicit check and validate only the
                // selected Word range. Never hide a whole-document capture/scan behind
                // this button.
                var selectedFinding = scan == null ? null : scan.Findings.FirstOrDefault(item =>
                    string.Equals(item.FindingId, selectedFindingId, StringComparison.Ordinal));
                if (context.LastLocalSnapshot == null || selectedFinding == null)
                    throw new InvalidOperationException(
                        "Kết quả kiểm tra của lỗi đang chọn không còn trong phiên hiện tại. " +
                        "Hãy bấm Kiểm tra thể thức hoặc Kiểm tra chính tả rồi chọn lại lỗi.");
                var result = _oneClickRuntime.ExecuteSelectedFinding(
                    context, selectedLane, selectedFindingId,
                    selectedStory, selectedStart, selectedEnd, document,
                    _currentDocumentOperation);
                if (!result.Resolved)
                {
                    MessageBox.Show(
                        "Lỗi này cần người dùng quyết định hoặc bổ sung nội dung nên không thể tự sửa an toàn.\n\n" +
                        "Yêu cầu đúng: " + selectedFinding.Expected + "\n\n" +
                        "Comment và phần tô đỏ được giữ nguyên.",
                        "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (OperationCanceledException)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                _contextStore.GetOrCreate(document).ClearReadAnalysis();
            }
            catch (Exception exception)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                MessageBox.Show("Không thể sửa lỗi đang chọn.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                _currentDocumentOperation?.Dispose();
                _currentDocumentOperation = null;
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateAll();
            }
        }

        private void RunLocalScan(bool spelling)
        {
            var document = TryGetCurrentDocument();
            if (document == null)
            {
                MessageBox.Show("Hãy mở một tài liệu Word.", "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Interlocked.CompareExchange(ref _documentOperationInProgress, 1, 0) != 0)
            {
                MessageBox.Show(
                    "Ứng dụng đang kiểm tra tài liệu. Hãy chờ lần kiểm tra hiện tại hoàn tất.",
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                _currentDocumentOperation = new DocumentOperationSession(
                    _application,
                    spelling ? "Kiểm tra chính tả" : "Kiểm tra thể thức",
                    DocumentOperationState.Capturing);
                FocusDocumentForCommand(document);
                var context = _contextStore.GetOrCreate(document);
                _documentReadRuntime.Prepare(
                    context,
                    spelling ? DocumentAnalysisScope.Spelling : DocumentAnalysisScope.Format,
                    document,
                    true,
                    _currentDocumentOperation);
                var result = _localScanRuntime.ScanAndAnnotate(context, spelling, document,
                    _currentDocumentOperation);
                SynchronizeDocumentTypeSelection(context);
                var nd30Findings = result.Findings.Count(f =>
                    f.SourceFamily == AnnotationSourceFamily.Nd30);
                var hd05Findings = result.Findings.Count(f =>
                    f.SourceFamily == AnnotationSourceFamily.Hd05);
                var localFindings = result.Findings.Count(f =>
                    f.SourceFamily == AnnotationSourceFamily.LocalLanguage);
                var latexFindings = result.Findings.Count(f =>
                    f.SourceFamily == AnnotationSourceFamily.LatexTypst);
                var unknownFindings = result.Findings.Count(f =>
                    f.SourceFamily == AnnotationSourceFamily.Unknown ||
                    f.SourceFamily == AnnotationSourceFamily.NotEvaluated);
                var details = spelling
                    ? "Phát hiện và đánh dấu: " + result.Findings.Count + " lỗi.\n"
                    : "Phát hiện và đánh dấu: " + result.Findings.Count + " điểm cần lưu ý:\n" +
                      "  • NĐ30: " + nd30Findings + ".\n" +
                      "  • HD05: " + hd05Findings + ".\n" +
                      "  • Quy tắc ngôn ngữ/local: " + localFindings + ".\n" +
                      "  • Khuyến nghị LaTeX/Typst: " + latexFindings + ".\n" +
                      (unknownFindings > 0 ? "  • Chưa phân loại/chưa đánh giá: " + unknownFindings + ".\n" : string.Empty) +
                      "Heading học thuật nhận diện: " + result.AcademicHeadingCount +
                      " (cấp 1: " + result.HeadingLevel1Count +
                      ", cấp 2: " + result.HeadingLevel2Count +
                      ", cấp 3: " + result.HeadingLevel3Count + ").\n" +
                      (result.AcademicTypographyEnabled
                          ? "AcademicTypography: bật theo gói quy tắc có chữ ký.\n"
                          : "AcademicTypography: tắt theo gói quy tắc có chữ ký.\n") +
                      (result.NotEvaluatedRuleCodes.Count > 0
                          ? "Chưa thể đánh giá: " + string.Join(", ", result.NotEvaluatedRuleCodes) + ".\n"
                          : string.Empty);
                MessageBox.Show(
                    "Đã kiểm tra " + (spelling ? "chính tả" : "thể thức") + " hoàn toàn tại máy.\n" +
                    "Loại văn bản tự nhận diện: " + LocalDocumentTypeCodes.GetDisplayName(context.DocumentTypeCode) + ".\n" +
                    details +
                    "Gói quy tắc: " + result.RulePackId + ".\n\nNội dung tài liệu không được gửi lên máy chủ.",
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                _contextStore.GetOrCreate(document).ClearReadAnalysis();
                MessageBox.Show(
                    "Đã hủy kiểm tra. Tài liệu chưa bị thay đổi.",
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                MessageBox.Show(
                    "Không thể chạy kiểm tra. Tài liệu chưa bị thay đổi.\n\n" + exception.Message,
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _currentDocumentOperation?.Dispose();
                _currentDocumentOperation = null;
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateAll();
            }
        }

        private void RunLocalCommand(string title, Func<string> command, bool showSuccessNotification = true)
        {
            if (Interlocked.CompareExchange(ref _documentOperationInProgress, 1, 0) != 0)
            {
                MessageBox.Show("Ứng dụng đang xử lý tài liệu. Hãy chờ thao tác hiện tại hoàn tất.",
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var completed = false;
            try
            {
                _currentDocumentOperation = new DocumentOperationSession(
                    _application, title, DocumentOperationState.Mutating);
                _currentDocumentOperation.Transition(DocumentOperationState.Mutating,
                    cancellationEnabled: false);
                var document = TryGetCurrentDocument();
                if (document == null) throw new InvalidOperationException("Hãy mở một tài liệu Word.");
                FocusDocumentForCommand(document);
                var backup = command();
                completed = true;
                if (showSuccessNotification)
                {
                    var backupMessage = string.IsNullOrWhiteSpace(backup)
                        ? "Có thể hoàn tác bằng Ctrl+Z."
                        : "Có thể hoàn tác bằng Ctrl+Z. Bản sao khôi phục đã được tạo tại:\n" + backup;
                    MessageBox.Show(
                        "Đã thực hiện “" + title + "” hoàn toàn tại máy.\n\n" +
                        backupMessage,
                        "Chuẩn hóa",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                _currentDocumentOperation?.MarkFailedRecoverable();
                MessageBox.Show(
                    "Không thể thực hiện “" + title + "”.\n\n" + exception.Message,
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (completed)
                {
                    var context = TryGetActiveContext();
                    if (context != null) context.ClearReadAnalysis();
                }
                _currentDocumentOperation?.Dispose();
                _currentDocumentOperation = null;
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateAll();
            }
        }

        private void RunAnalysisBackedLocalCommand(
            string title,
            Func<DocumentContext, string> command)
        {
            if (command == null) throw new ArgumentNullException(nameof(command));
            RunLocalCommand(title, () =>
            {
                var document = TryGetCurrentDocument();
                if (document == null) throw new InvalidOperationException("Hãy mở một tài liệu Word.");
                var context = _contextStore.GetOrCreate(document);
                _currentDocumentOperation?.Transition(DocumentOperationState.Capturing,
                    "chuẩn bị dữ liệu");
                _documentReadRuntime.Prepare(
                    context,
                    DocumentAnalysisScope.SnapshotOnly,
                    document,
                    true,
                    _currentDocumentOperation);
                SynchronizeDocumentTypeSelection(context);
                _currentDocumentOperation?.Transition(DocumentOperationState.Mutating,
                    cancellationEnabled: false);
                return command(context);
            });
        }

        private static void SynchronizeDocumentTypeSelection(DocumentContext context)
        {
            context.SetSelection("ddLoaiVanBan", MapDocumentTypeIndex(context.DocumentTypeCode));
        }

        private static string MapRegime(string selectedId)
        {
            switch (selectedId)
            {
                case "iqdND30":
                    return "ND30";
                case "iqdVIETTEL":
                    return "VIETTEL";
                case "iqdDANG":
                    return "PARTY_HD05";
                default:
                    return "UNKNOWN";
            }
        }

        private static string MapDocumentType(int selectedIndex)
        {
            return selectedIndex >= 0 && selectedIndex < DocumentTypeCodes.Count
                ? DocumentTypeCodes[selectedIndex]
                : LocalDocumentTypeCodes.Unknown;
        }

        private static int MapDocumentTypeIndex(string code)
        {
            for (var index = 0; index < DocumentTypeCodes.Count; index++)
                if (string.Equals(DocumentTypeCodes[index], code, StringComparison.OrdinalIgnoreCase))
                    return index;
            return 0;
        }

        private void FailClosed(string controlId)
        {
            var reason = _wordMutationRuntime.HasAuthorizationKeys
                ? "chưa có implementation đã vượt qua exit gate"
                : "chưa có khóa xác minh authorization production và implementation đã vượt qua exit gate";
            MessageBox.Show(
                "Lệnh " + controlId + " " + reason + " nên đang bị khóa an toàn.",
                "Chuẩn hóa",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private void Invalidate(string controlId)
        {
            if (_ribbonUi != null)
            {
                _ribbonUi.InvalidateControl(controlId);
            }
        }

        private void InvalidateAll()
        {
            if (_ribbonUi != null)
            {
                _ribbonUi.Invalidate();
            }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }

    }
}
