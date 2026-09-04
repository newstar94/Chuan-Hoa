using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Deployment.Application;
using ChuanHoa.AddIn.Vsto.Ribbon;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;
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
        private readonly WordDocumentCapabilityProvider _capabilityProvider;
        private readonly WordDocumentReadRuntime _documentReadRuntime;
        private readonly LocalAccessManager _localAccessManager;
        private readonly WordLocalScanRuntime _localScanRuntime;
        private readonly WordLocalCommandRuntime _localCommandRuntime;
        private readonly WordOneClickRuntime _oneClickRuntime;
        private readonly Dictionary<string, Action> _buttonCommands;
        private readonly Dictionary<string, string> _viewProperties;
        private readonly System.Windows.Forms.Timer _accessRefreshTimer;
        private Office.IRibbonUI? _ribbonUi;
        private Word.Document? _lastActivatedDocument;
        private Word.Document? _ribbonCapabilityDocument;
        private WordDocumentCapability? _ribbonCapability;
        private long _ribbonCapabilityTimestamp;
        private int _accessRefreshCompleted;
        private int _documentOperationInProgress;
        private bool _disposed;

        public RibbonRuntime(
            Word.Application application,
            DocumentContextStore contextStore,
            WordMutationRuntime wordMutationRuntime)
        {
            _application = application ?? throw new ArgumentNullException("application");
            _contextStore = contextStore ?? throw new ArgumentNullException("contextStore");
            _wordMutationRuntime = wordMutationRuntime ?? throw new ArgumentNullException("wordMutationRuntime");
            _capabilityProvider = new WordDocumentCapabilityProvider(_application);
            var assemblyVersion = typeof(RibbonRuntime).Assembly.GetName().Version;
            _localAccessManager = new LocalAccessManager(assemblyVersion == null ? "1.0.0.0" : assemblyVersion.ToString());
            _documentReadRuntime = new WordDocumentReadRuntime(_application, _localAccessManager);
            _localScanRuntime = new WordLocalScanRuntime(_application, _localAccessManager);
            _localCommandRuntime = new WordLocalCommandRuntime(_application, _localAccessManager, TryGetCurrentDocument);
            _oneClickRuntime = new WordOneClickRuntime(_application, _localAccessManager);
            _localAccessManager.CacheStateChanged += OnCacheStateChanged;
            _accessRefreshTimer = new System.Windows.Forms.Timer { Interval = 500 };
            _accessRefreshTimer.Tick += OnAccessRefreshTimerTick;
            _accessRefreshTimer.Start();
            _localAccessManager.WarmUp();
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
                { "btnChenQrCode", InsertQrCode },
                { "btnKieuOaUy", () => RunLocalCommand("Kiểu oà, uý", () => _localCommandRuntime.NormalizeTonePlacement(VietnameseTonePlacementStyle.MainVowel)) },
                { "btnKieuOaUy2", () => RunLocalCommand("Kiểu òa, úy", () => _localCommandRuntime.NormalizeTonePlacement(VietnameseTonePlacementStyle.FirstVowel)) },
                { "btnDoiDauThapPhan", () => RunLocalCommand("Dấu phẩy thập phân", _localCommandRuntime.ConvertDecimalSeparators) },
                { "btnTuDienCaNhan", () => _localCommandRuntime.OpenCustomDictionaryDialog() },
                { "btnKiemTraPhienBanMoi", ShowUpdateStatus },
                { "btnGuiPhanHoi", OpenFeedback },
                { "btnGioiThieu", ShowAbout }
            };
            _viewProperties = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "chkRanhGioiVanBan", "ShowTextBoundaries" },
                { "chkDauGoc", "ShowCropMarks" },
                { "chkKyHieuSoanThao", "ShowAll" }
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
            if (string.Equals(controlId, "btnGioiThieu", StringComparison.Ordinal) ||
                string.Equals(controlId, "btnKiemTraPhienBanMoi", StringComparison.Ordinal) ||
                string.Equals(controlId, "btnGuiPhanHoi", StringComparison.Ordinal) ||
                string.Equals(controlId, "mnuThongTinTienIch", StringComparison.Ordinal))
            {
                return true;
            }

            if (Volatile.Read(ref _documentOperationInProgress) != 0)
            {
                return false;
            }

            var capability = EvaluateRibbonCapability();
            var canBeginSavedMutation = capability.CanReadDocument ||
                (capability.HasActiveDocument && capability.HasActiveWindow && !capability.IsSaved);

            if (_viewProperties.ContainsKey(controlId))
            {
                return capability.HasActiveWindow;
            }

            if (string.Equals(controlId, "btnKiemTra", StringComparison.Ordinal) ||
                string.Equals(controlId, "btnKiemTraChinhTa", StringComparison.Ordinal) ||
                string.Equals(controlId, "ddQuyDinh", StringComparison.Ordinal) ||
                string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal))
            {
                if (!capability.CanReadDocument) return false;
                if (string.Equals(controlId, "btnKiemTra", StringComparison.Ordinal))
                    return _localAccessManager.HasCachedFeature(LocalAccessManager.FormatFeature);
                if (string.Equals(controlId, "btnKiemTraChinhTa", StringComparison.Ordinal))
                    return _localAccessManager.HasCachedFeature(LocalAccessManager.SpellingFeature);
                return true;
            }

            if (string.Equals(controlId, "btnAutoFixAll2026", StringComparison.Ordinal) ||
                string.Equals(controlId, "btnSuaTatCaChinhTa", StringComparison.Ordinal) ||
                string.Equals(controlId, "btnSuaLoiDangChon", StringComparison.Ordinal))
            {
                return canBeginSavedMutation && !capability.IsReadOnly && !capability.IsProtected &&
                    !capability.TrackChangesEnabled &&
                    _localAccessManager.HasCachedFeature(LocalAccessManager.AutoFixFeature);
            }

            if (string.Equals(controlId, "mnuBoDau", StringComparison.Ordinal))
            {
                return canBeginSavedMutation && !capability.IsReadOnly && !capability.IsProtected &&
                    !capability.TrackChangesEnabled &&
                    _localAccessManager.HasCachedFeature(LocalAccessManager.DocumentToolsFeature);
            }

            if (_buttonCommands.ContainsKey(controlId))
            {
                return canBeginSavedMutation && !capability.IsReadOnly && !capability.IsProtected &&
                    !capability.TrackChangesEnabled &&
                    _localAccessManager.HasCachedFeature(LocalAccessManager.DocumentToolsFeature);
            }

            return false;
        }

        public bool GetPressed(string controlId)
        {
            ThrowIfDisposed();
            string propertyName;
            if (!_viewProperties.TryGetValue(controlId, out propertyName))
            {
                return false;
            }

            return ReadBooleanViewProperty(propertyName);
        }

        public void SetPressed(string controlId, bool pressed)
        {
            ThrowIfDisposed();
            string propertyName;
            if (!_viewProperties.TryGetValue(controlId, out propertyName))
            {
                FailClosed(controlId);
                return;
            }

            WriteBooleanViewProperty(propertyName, pressed);
            Invalidate(controlId);
        }

        public int GetSelectedItemIndex(string controlId)
        {
            ThrowIfDisposed();
            var context = TryGetActiveContext();
            return context == null ? 0 : context.GetSelection(controlId);
        }

        public void SelectDropDownItem(string controlId, string selectedId, int selectedIndex)
        {
            ThrowIfDisposed();
            var context = TryGetActiveContext();
            if (context == null)
            {
                return;
            }

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
                ? DocumentTypeCodes.Count
                : 0;
        }

        public string GetItemLabel(string controlId, int index)
        {
            ThrowIfDisposed();
            if (!string.Equals(controlId, "ddLoaiVanBan", StringComparison.Ordinal) ||
                index < 0 ||
                index >= DocumentTypeCodes.Count)
            {
                return string.Empty;
            }

            return LocalDocumentTypeCodes.GetDisplayName(DocumentTypeCodes[index]);
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
        }

        public void OnDocumentBeforeClose(Word.Document document)
        {
            if (_disposed || document == null)
            {
                return;
            }

            _contextStore.Remove(document);
            if (IsSameComDocument(_lastActivatedDocument, document))
            {
                _lastActivatedDocument = null;
            }
            if (IsSameComDocument(_ribbonCapabilityDocument, document))
            {
                _ribbonCapabilityDocument = null;
                _ribbonCapability = null;
                _ribbonCapabilityTimestamp = 0;
            }
            InvalidateRibbonCapability();
            InvalidateAll();
        }

        public void OnDocumentBeforeSave(Word.Document document)
        {
            if (_disposed || document == null) return;
            _contextStore.GetOrCreate(document).ClearReadAnalysis();
        }

        public void OnDocumentWindowActivated(Word.Document document)
        {
            if (_disposed || document == null)
            {
                return;
            }
            // WindowActivate fires not only when another document becomes active but
            // also whenever the user Alt-Tabs away from Word and comes back. Never read
            // document content or build a snapshot here: that work runs on Word's UI
            // thread and can freeze large documents without a user command.
            // Word can raise WindowActivate before the user's first window reports
            // visible, so always capture the first document. Once a user document is
            // known, ignore a zero-window activation: that is the hidden recovery clone
            // created by 1-Click, or a transient Modern Comments focus state.
            if (_lastActivatedDocument != null)
            {
                try
                {
                    if (document.Windows.Count == 0) return;
                }
                catch (COMException)
                {
                    return;
                }
            }
            _lastActivatedDocument = document;
            _contextStore.GetOrCreate(document);
            CaptureStableRibbonCapability(document);
            InvalidateAll();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _ribbonUi = null;
            _lastActivatedDocument = null;
            _ribbonCapabilityDocument = null;
            _ribbonCapability = null;
            _localAccessManager.CacheStateChanged -= OnCacheStateChanged;
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

        private void OnAccessRefreshTimerTick(object? sender, EventArgs eventArgs)
        {
            if (_disposed || Interlocked.Exchange(ref _accessRefreshCompleted, 0) == 0) return;
            // The refreshed lease can unlock every DOCUMENT_TOOLS command, not only
            // the scanners and 1-Click. Refresh the complete Ribbon so newly ported
            // style, Unicode, font-size, QR and tone commands do not remain grey.
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
            var fallback = _lastActivatedDocument ?? _ribbonCapabilityDocument;
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

        private WordDocumentCapability EvaluateCurrentDocumentCapability()
        {
            var document = TryGetCurrentDocument();
            return document == null ? _capabilityProvider.Evaluate() : _capabilityProvider.Evaluate(document);
        }

        private WordDocumentCapability EvaluateRibbonCapability()
        {
            var document = TryGetCurrentDocument();
            var now = Stopwatch.GetTimestamp();
            var maximumAge = Stopwatch.Frequency / 2;
            var sameCachedDocument = document != null && _ribbonCapabilityDocument != null &&
                IsSameComDocument(_ribbonCapabilityDocument, document);
            if (_ribbonCapability != null &&
                sameCachedDocument &&
                now - _ribbonCapabilityTimestamp >= 0 &&
                now - _ribbonCapabilityTimestamp <= maximumAge)
            {
                return _ribbonCapability;
            }

            var capability = document == null
                ? _capabilityProvider.Evaluate()
                : _capabilityProvider.Evaluate(document);
            // Office calls getEnabled while focus is moving from a Modern Comment to
            // the Ribbon. If a captured document rejects one transient property read,
            // keep the last capability for that exact COM document. A temporary focus
            // transition must not grey every control or overwrite a known READY state.
            if (sameCachedDocument && _ribbonCapability != null &&
                string.Equals(capability.ReasonCode,
                    WordDocumentCapabilityProvider.TransientStateReasonCode,
                    StringComparison.Ordinal))
            {
                return _ribbonCapability;
            }
            _ribbonCapabilityDocument = document;
            _ribbonCapability = capability;
            _ribbonCapabilityTimestamp = now;
            UpdateCapabilityContext(capability);
            return capability;
        }

        private void InvalidateRibbonCapability()
        {
            // Expire the value, but retain it as a same-document fallback for the
            // brief COM rejection Word produces during Modern Comments focus changes.
            // A different document can never reuse it because evaluation compares COM
            // identity before taking the fallback.
            _ribbonCapabilityTimestamp = 0;
        }

        private void CaptureStableRibbonCapability(Word.Document document)
        {
            // Prime the lightweight capability cache while Word owns normal document
            // focus. A user can open the Chuẩn hóa tab only after selecting a Modern
            // Comment; in that state Word may reject FullName/Path on the very first
            // getEnabled callback. Without this stable baseline the transient result
            // becomes the cache and the scan buttons stay grey after the selected fix.
            // This reads only document metadata and never scans document content.
            var capability = _capabilityProvider.Evaluate(document);
            if (string.Equals(capability.ReasonCode,
                    WordDocumentCapabilityProvider.TransientStateReasonCode,
                    StringComparison.Ordinal))
            {
                InvalidateRibbonCapability();
                return;
            }

            _ribbonCapabilityDocument = document;
            _ribbonCapability = capability;
            _ribbonCapabilityTimestamp = Stopwatch.GetTimestamp();
            UpdateCapabilityContext(capability);
        }

        private void FocusDocumentForCommand(Word.Document document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var annotations = new WordFindingAnnotationAdapter(_application, document);
            // Document.Activate itself is rejected while Modern Comments owns focus.
            // Selecting the comment's document scope is Word's supported focus hand-off
            // and does not read or scan any unrelated document content.
            if (!annotations.TryFocusDocumentSelection())
                throw new InvalidOperationException(
                    "Word đang hoàn tất chuyển tiêu điểm từ khung comment. Hãy thử lại thao tác.");
            _lastActivatedDocument = document;
            System.Windows.Forms.Application.DoEvents();
        }

        private DocumentContext RequireActiveContext()
        {
            var context = TryGetActiveContext();
            if (context == null) throw new InvalidOperationException("Hãy mở một tài liệu Word.");
            return context;
        }

        private void InsertQrCode()
        {
            var content = QrCodeInputDialog.Prompt(null);
            if (content == null) return;
            RunLocalCommand("Chèn mã QR", () => _localCommandRuntime.InsertQrCode(content));
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

        private bool ReadBooleanViewProperty(string propertyName)
        {
            try
            {
                var view = _application.ActiveWindow.View;
                var property = view.GetType().GetProperty(
                    propertyName,
                    BindingFlags.Public | BindingFlags.Instance);
                if (property == null || property.PropertyType != typeof(bool))
                {
                    return false;
                }

                return (bool)property.GetValue(view, null);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private void WriteBooleanViewProperty(string propertyName, bool value)
        {
            var view = _application.ActiveWindow.View;
            var property = view.GetType().GetProperty(
                propertyName,
                BindingFlags.Public | BindingFlags.Instance);
            if (property == null || !property.CanWrite || property.PropertyType != typeof(bool))
            {
                throw new NotSupportedException(
                    "The current Word window does not expose view option " + propertyName + ".");
            }

            property.SetValue(view, value, null);
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
                var activeDocument = TryGetCurrentDocument();
                if (activeDocument == null) throw new InvalidOperationException("Hãy mở một tài liệu Word.");
                FocusDocumentForCommand(activeDocument);
                // Resolve the context again after the confirmation dialog so this
                // command cannot carry state from a previously active document.
                context = _contextStore.GetOrCreate(activeDocument);
                _documentReadRuntime.PrepareForOneClick(context, activeDocument);
                var result = _oneClickRuntime.Execute(context, activeDocument);
                // Rebuild from the mutated document and re-apply only findings that are
                // still present. This guarantees that a comment is never removed merely
                // because 1-Click attempted a fix. It also proves deterministic edits,
                // including the missing word "số", against the actual Word content.
                _documentReadRuntime.Prepare(
                    context,
                    DocumentAnalysisScope.Full,
                    activeDocument,
                    false);
                var remainingFormat = _localScanRuntime.ScanAndAnnotate(context, false, activeDocument);
                var remainingSpelling = _localScanRuntime.ScanAndAnnotate(context, true, activeDocument);
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
            catch (Exception exception)
            {
                MessageBox.Show("Không thể chuẩn hóa toàn bộ.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateRibbonCapability();
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
                FocusDocumentForCommand(document);
                var context = _contextStore.GetOrCreate(document);
                _documentReadRuntime.Prepare(context, DocumentAnalysisScope.Spelling, document, false);
                _localScanRuntime.ScanAndAnnotate(context, true, document);

                var fixedCount = _oneClickRuntime.FixAllSpellingFindings(context, document);

                _documentReadRuntime.Prepare(context, DocumentAnalysisScope.Spelling, document, false);
                var remainingScan = _localScanRuntime.ScanAndAnnotate(context, true, document);

                MessageBox.Show(
                    "Đã sửa nhanh chính tả tại máy.\n\n" +
                    "Số lỗi đã tự động sửa thành công: " + fixedCount + "\n" +
                    "Số lỗi còn lại cần người dùng cân nhắc: " + remainingScan.Findings.Count + "\n\n" +
                    "Bạn có thể hoàn tác toàn bộ bằng Ctrl+Z nếu cần.",
                    "Sửa nhanh chính tả",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Không thể sửa nhanh chính tả.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateRibbonCapability();
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
                DocumentAnalysisScope scope;
                if (string.Equals(selectedLane, "spelling", StringComparison.OrdinalIgnoreCase))
                {
                    scope = DocumentAnalysisScope.Spelling;
                }
                else if (string.Equals(selectedLane, "format", StringComparison.OrdinalIgnoreCase))
                {
                    scope = DocumentAnalysisScope.Format;
                }
                else
                {
                    throw new InvalidOperationException(
                        "Comment đang chọn không thuộc nhóm lỗi Chuẩn hóa hỗ trợ.");
                }
                FocusDocumentForCommand(document);
                var context = _contextStore.GetOrCreate(document);
                var scan = string.Equals(selectedLane, "spelling", StringComparison.OrdinalIgnoreCase)
                    ? context.LastSpellingScan
                    : context.LastFormatScan;
                var hasFinding = context.LastLocalSnapshot != null && scan != null &&
                    scan.Findings.Any(item => string.Equals(item.FindingId, selectedFindingId, StringComparison.Ordinal));
                if (!hasFinding)
                {
                    _documentReadRuntime.Prepare(context, scope, document);
                }
                var result = _oneClickRuntime.ExecuteSelectedFinding(
                    context, selectedLane, selectedFindingId,
                    selectedStory, selectedStart, selectedEnd, document);
                if (!result.Resolved)
                {
                    MessageBox.Show(
                        "Đã áp dụng phương án sửa, nhưng lỗi vẫn còn theo lần kiểm tra lại. " +
                        "Comment được giữ để bạn xem tiếp. Có thể hoàn tác bằng Ctrl+Z.",
                        "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Không thể sửa lỗi đang chọn.\n\n" + exception.Message,
                    "Chuẩn hóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateRibbonCapability();
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
                FocusDocumentForCommand(document);
                var context = _contextStore.GetOrCreate(document);
                _documentReadRuntime.Prepare(
                    context,
                    spelling ? DocumentAnalysisScope.Spelling : DocumentAnalysisScope.Format,
                    document);
                var result = _localScanRuntime.ScanAndAnnotate(context, spelling, document);
                SynchronizeDocumentTypeSelection(context);
                MessageBox.Show(
                    "Đã kiểm tra " + (spelling ? "chính tả" : "thể thức") + " hoàn toàn tại máy.\n" +
                    "Loại văn bản tự nhận diện: " + LocalDocumentTypeCodes.GetDisplayName(context.DocumentTypeCode) + ".\n" +
                    "Phát hiện và đánh dấu: " + result.Findings.Count + " lỗi.\n" +
                    "Gói quy tắc: " + result.RulePackId + ".\n\nNội dung tài liệu không được gửi lên máy chủ.",
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "Không thể chạy kiểm tra. Tài liệu chưa bị thay đổi.\n\n" + exception.Message,
                    "Chuẩn hóa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateRibbonCapability();
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
                Interlocked.Exchange(ref _documentOperationInProgress, 0);
                InvalidateRibbonCapability();
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
                _documentReadRuntime.Prepare(
                    context,
                    DocumentAnalysisScope.SnapshotOnly,
                    document);
                SynchronizeDocumentTypeSelection(context);
                return command(context);
            });
        }

        private void UpdateCapabilityContext(WordDocumentCapability capability)
        {
            var context = TryGetActiveContext();
            if (context == null)
            {
                return;
            }

            context.LastCapabilityReasonCode = capability.ReasonCode;
            context.LastCapabilityReason = capability.Reason;
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
