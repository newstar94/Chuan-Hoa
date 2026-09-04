using System;
using System.Runtime.InteropServices;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public enum DocumentOperationState
    {
        Idle = 0,
        Capturing = 1,
        Scanning = 2,
        Annotating = 3,
        Mutating = 4,
        Cancelling = 5,
        FailedRecoverable = 6
    }

    /// <summary>
    /// Cooperative, command-scoped operation state. All methods that touch Word's
    /// StatusBar must be called on the same STA thread that started the Ribbon
    /// command. Cancellation is polled from the keyboard at bounded batch boundaries;
    /// it never pumps messages and never moves Word COM work to a worker thread.
    /// </summary>
    public sealed class DocumentOperationSession : IDisposable
    {
        private const int VirtualKeyEscape = 0x1B;
        private readonly Word.Application _application;
        private readonly string _title;
        private readonly int _ownerThreadId;
        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private int _state;
        private int _cancellationEnabled;
        private int _cancellationPermanentlyDisabled;
        private int _disposed;

        public DocumentOperationSession(Word.Application application, string title,
            DocumentOperationState initialState)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _title = string.IsNullOrWhiteSpace(title) ? "Chuẩn hóa" : title.Trim();
            _ownerThreadId = Environment.CurrentManagedThreadId;
            _state = (int)initialState;
            var startsMutating = initialState == DocumentOperationState.Mutating;
            _cancellationEnabled = startsMutating ? 0 : 1;
            _cancellationPermanentlyDisabled = startsMutating ? 1 : 0;
            // Do not let a stale key transition from before the command become a
            // cancellation request for this operation.
            GetAsyncKeyState(VirtualKeyEscape);
            UpdateStatus(null, null, null);
        }

        public DocumentOperationState State =>
            (DocumentOperationState)Volatile.Read(ref _state);

        public CancellationToken CancellationToken => _cancellation.Token;

        public void Transition(DocumentOperationState state, string? detail = null,
            bool cancellationEnabled = true)
        {
            EnsureOwnerThread();
            ThrowIfDisposed();
            if (state == DocumentOperationState.Idle)
                throw new ArgumentOutOfRangeException(nameof(state));
            var disablePermanently = state == DocumentOperationState.Mutating || !cancellationEnabled;
            if (disablePermanently && Volatile.Read(ref _cancellationPermanentlyDisabled) == 0)
            {
                // Honor a pending Esc before crossing the first mutation boundary.
                // Once mutation starts, cancellation stays disabled even if the
                // command later returns to capture/scan for post-condition checks.
                Checkpoint();
                Volatile.Write(ref _cancellationPermanentlyDisabled, 1);
            }
            Volatile.Write(ref _state, (int)state);
            Volatile.Write(ref _cancellationEnabled,
                cancellationEnabled && Volatile.Read(ref _cancellationPermanentlyDisabled) == 0 ? 1 : 0);
            Checkpoint();
            UpdateStatus(detail, null, null);
        }

        public void ReportProgress(int completed, int total, string? detail = null)
        {
            EnsureOwnerThread();
            ThrowIfDisposed();
            Checkpoint();
            UpdateStatus(detail, Math.Max(0, completed), Math.Max(0, total));
        }

        public void Checkpoint()
        {
            EnsureOwnerThread();
            ThrowIfDisposed();
            if (Volatile.Read(ref _cancellationEnabled) != 0 &&
                (GetAsyncKeyState(VirtualKeyEscape) & 0x8000) != 0)
            {
                Volatile.Write(ref _state, (int)DocumentOperationState.Cancelling);
                _cancellation.Cancel();
            }
            _cancellation.Token.ThrowIfCancellationRequested();
        }

        public void MarkFailedRecoverable()
        {
            if (Volatile.Read(ref _disposed) != 0) return;
            Volatile.Write(ref _state, (int)DocumentOperationState.FailedRecoverable);
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            EnsureOwnerThread();
            Volatile.Write(ref _state, (int)DocumentOperationState.Idle);
            try { _application.StatusBar = string.Empty; }
            catch (COMException) { }
            _cancellation.Dispose();
        }

        private void UpdateStatus(string? detail, int? completed, int? total)
        {
            var state = State;
            var phase = state == DocumentOperationState.Capturing ? "đang đọc tài liệu" :
                state == DocumentOperationState.Scanning ? "đang kiểm tra quy tắc" :
                state == DocumentOperationState.Annotating ? "đang đánh dấu lỗi" :
                state == DocumentOperationState.Mutating ? "đang chỉnh sửa tài liệu" :
                state == DocumentOperationState.Cancelling ? "đang hủy an toàn" :
                state == DocumentOperationState.FailedRecoverable ? "đã phục hồi sau lỗi" :
                "đang xử lý";
            var completedValue = completed.GetValueOrDefault();
            var totalValue = total.GetValueOrDefault();
            var progress = completed.HasValue && totalValue > 0
                ? " (" + Math.Min(completedValue, totalValue) + "/" + totalValue + ")"
                : string.Empty;
            var suffix = string.IsNullOrWhiteSpace(detail) ? string.Empty : ": " + detail!.Trim();
            var cancellation = Volatile.Read(ref _cancellationEnabled) != 0
                ? " — nhấn Esc để hủy"
                : string.Empty;
            try { _application.StatusBar = _title + " — " + phase + progress + suffix + cancellation; }
            catch (COMException) { }
        }

        private void EnsureOwnerThread()
        {
            if (Environment.CurrentManagedThreadId != _ownerThreadId)
                throw new InvalidOperationException(
                    "Document operation state must remain on the Word STA thread.");
        }

        private void ThrowIfDisposed()
        {
            if (Volatile.Read(ref _disposed) != 0)
                throw new ObjectDisposedException(nameof(DocumentOperationSession));
        }

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int virtualKey);
    }
}
