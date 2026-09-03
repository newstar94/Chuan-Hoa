using System;
using System.Collections.Generic;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class DocumentContext
    {
        private readonly Dictionary<string, int> _dropDownSelections =
            new Dictionary<string, int>(StringComparer.Ordinal);

        public DocumentContext(long documentIdentity)
        {
            DocumentIdentity = documentIdentity;
            CreatedAtUtc = DateTimeOffset.UtcNow;
        }

        public long DocumentIdentity { get; }

        public DateTimeOffset CreatedAtUtc { get; }

        public string RegimeCode { get; set; } = "ND30";

        public string DocumentTypeCode { get; set; } = "UNKNOWN";

        public bool RegimeWasSelectedManually { get; set; }

        public bool DocumentTypeWasSelectedManually { get; set; }

        public WordDocumentSnapshot? LastSnapshot { get; private set; }

        public LocalScanSnapshot? LastLocalSnapshot { get; private set; }

        public LocalScanResult? LastFormatScan { get; private set; }

        public LocalScanResult? LastSpellingScan { get; private set; }

        public DateTimeOffset? LastSnapshotAtUtc { get; private set; }

        public bool SnapshotCapturedFromSavedDocument { get; private set; }

        public string LastCapabilityReasonCode { get; set; } = "ACTIVE_DOCUMENT_REQUIRED";

        public string LastCapabilityReason { get; set; } = "Hãy mở một tài liệu Word.";

        private long _snapshotRevision;

        public long NextSnapshotRevision()
        {
            _snapshotRevision++;
            return _snapshotRevision;
        }

        public void SetAnalysis(
            WordDocumentSnapshot snapshot,
            LocalScanSnapshot localSnapshot,
            LocalScanResult? formatScan,
            LocalScanResult? spellingScan,
            bool capturedFromSavedDocument)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (localSnapshot == null) throw new ArgumentNullException("localSnapshot");
            if (!AnalysisIdentityMatches(snapshot, localSnapshot))
                throw new InvalidOperationException("Dữ liệu đọc tài liệu không đồng nhất.");
            if (formatScan != null && !ScanIdentityMatches(localSnapshot, formatScan, "format"))
                throw new InvalidOperationException("Kết quả kiểm tra thể thức không thuộc dữ liệu tài liệu hiện tại.");
            if (spellingScan != null && !ScanIdentityMatches(localSnapshot, spellingScan, "spelling"))
                throw new InvalidOperationException("Kết quả kiểm tra chính tả không thuộc dữ liệu tài liệu hiện tại.");

            // Commit the complete analysis atomically. Validation above must finish
            // before any property changes so a failed build cannot leave a reusable
            // but only partially populated cache behind.
            LastSnapshot = snapshot;
            LastLocalSnapshot = localSnapshot;
            LastFormatScan = formatScan;
            LastSpellingScan = spellingScan;
            LastSnapshotAtUtc = DateTimeOffset.UtcNow;
            SnapshotCapturedFromSavedDocument = capturedFromSavedDocument;
        }

        public bool CanReuseAnalysis(
            bool documentIsSaved,
            bool requireFormat,
            bool requireSpelling,
            string? formatRulePackId,
            string? spellingRulePackId)
        {
            if (!documentIsSaved || !SnapshotCapturedFromSavedDocument ||
                LastSnapshot == null || LastLocalSnapshot == null ||
                !AnalysisIdentityMatches(LastSnapshot, LastLocalSnapshot))
                return false;
            if (requireFormat && (LastFormatScan == null ||
                !ScanIdentityMatches(LastLocalSnapshot, LastFormatScan, "format") ||
                !string.Equals(LastFormatScan.RulePackId, formatRulePackId, StringComparison.Ordinal)))
                return false;
            if (requireSpelling && (LastSpellingScan == null ||
                !ScanIdentityMatches(LastLocalSnapshot, LastSpellingScan, "spelling") ||
                !string.Equals(LastSpellingScan.RulePackId, spellingRulePackId, StringComparison.Ordinal)))
                return false;
            return true;
        }

        public void ClearReadAnalysis()
        {
            LastSnapshot = null;
            LastLocalSnapshot = null;
            LastFormatScan = null;
            LastSpellingScan = null;
            LastSnapshotAtUtc = null;
            SnapshotCapturedFromSavedDocument = false;
        }

        public void RequireSnapshotAnalysis()
        {
            if (LastSnapshot == null || LastLocalSnapshot == null ||
                !AnalysisIdentityMatches(LastSnapshot, LastLocalSnapshot))
            {
                throw new InvalidOperationException(
                    "Dữ liệu tài liệu chưa được chuẩn bị cho lệnh hiện tại.");
            }
        }

        public void RequireFormatAnalysis()
        {
            RequireSnapshotAnalysis();
            if (LastFormatScan == null ||
                !ScanIdentityMatches(LastLocalSnapshot!, LastFormatScan, "format"))
                throw new InvalidOperationException("Kết quả kiểm tra thể thức chưa được chuẩn bị.");
        }

        public void RequireSpellingAnalysis()
        {
            RequireSnapshotAnalysis();
            if (LastSpellingScan == null ||
                !ScanIdentityMatches(LastLocalSnapshot!, LastSpellingScan, "spelling"))
                throw new InvalidOperationException("Kết quả kiểm tra chính tả chưa được chuẩn bị.");
        }

        public void RequireFullAnalysis()
        {
            RequireFormatAnalysis();
            RequireSpellingAnalysis();
        }

        private static bool AnalysisIdentityMatches(
            WordDocumentSnapshot snapshot,
            LocalScanSnapshot localSnapshot)
        {
            return string.Equals(snapshot.DocumentFingerprint, localSnapshot.DocumentFingerprint,
                       StringComparison.Ordinal) &&
                snapshot.Revision == localSnapshot.Revision;
        }

        private static bool ScanIdentityMatches(
            LocalScanSnapshot snapshot,
            LocalScanResult scan,
            string expectedLane)
        {
            return string.Equals(scan.Lane, expectedLane, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(scan.DocumentFingerprint, snapshot.DocumentFingerprint,
                    StringComparison.Ordinal) &&
                scan.Revision == snapshot.Revision;
        }

        public int GetSelection(string controlId)
        {
            int selectedIndex;
            return _dropDownSelections.TryGetValue(controlId, out selectedIndex)
                ? selectedIndex
                : 0;
        }

        public void SetSelection(string controlId, int selectedIndex)
        {
            if (selectedIndex < 0)
            {
                throw new ArgumentOutOfRangeException("selectedIndex");
            }

            _dropDownSelections[controlId] = selectedIndex;
        }
    }
}
