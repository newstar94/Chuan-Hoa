using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class OneClickResult
    {
        public OneClickResult(string backupPath, int changedParagraphs, int insertedLines,
            int normalizedSections, int normalizedTables, int correctedSpellingItems,
            IReadOnlyList<AnnotationFinding> remainingFindings)
        {
            BackupPath = backupPath;
            ChangedParagraphs = changedParagraphs;
            InsertedLines = insertedLines;
            NormalizedSections = normalizedSections;
            NormalizedTables = normalizedTables;
            CorrectedSpellingItems = correctedSpellingItems;
            RemainingFindingItems = remainingFindings ?? throw new ArgumentNullException(nameof(remainingFindings));
        }

        public string BackupPath { get; }
        public int ChangedParagraphs { get; }
        public int InsertedLines { get; }
        public int NormalizedSections { get; }
        public int NormalizedTables { get; }
        public int CorrectedSpellingItems { get; }
        public int RemainingFindings => RemainingFindingItems.Count;
        public IReadOnlyList<AnnotationFinding> RemainingFindingItems { get; }
    }

    public sealed class SelectedFindingFixResult
    {
        public SelectedFindingFixResult(string backupPath, string lane, string findingId, bool resolved)
        {
            BackupPath = backupPath;
            Lane = lane;
            FindingId = findingId;
            Resolved = resolved;
        }

        public string BackupPath { get; }
        public string Lane { get; }
        public string FindingId { get; }
        public bool Resolved { get; }
    }

    /// <summary>
    /// Applies deterministic layout, typography and safe spelling fixes derived from
    /// the signed local rule pack. It never uploads document content.
    /// </summary>
    public sealed class WordOneClickRuntime
    {
        private const float PointsPerMillimeter = 72.0f / 25.4f;

        // Only header roles read the tier; plain body paragraphs never do.
        private static readonly Nd30HeaderFontSizeTier UnusedHeaderTier =
            new Nd30HeaderFontSizeTier(13d, 14d, 14d);

        // Text edits made during the current 1-Click phase. Word COM runs on one STA
        // thread; the field is set and cleared inside Execute only.
        [ThreadStatic]
        private static MainStoryEditTracker? s_tracker;
        private readonly Word.Application _application;
        private readonly WordDocumentCapabilityProvider _capabilityProvider;
        private readonly LocalAccessManager _accessManager;

        public WordOneClickRuntime(Word.Application application, LocalAccessManager accessManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _capabilityProvider = new WordDocumentCapabilityProvider(application);
            _accessManager = accessManager ?? throw new ArgumentNullException(nameof(accessManager));
        }

        public OneClickResult Execute(DocumentContext context, Word.Document? activeDocument = null,
            DocumentOperationSession? operation = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            var document = activeDocument ?? _application.ActiveDocument;
            // Persistence can raise DocumentBeforeSave and invalidate analysis. It must
            // happen before the invariant check; the production command prepares a
            // fresh full snapshot immediately after this save boundary.
            WordRecoveryCopyManager.EnsurePersistentDocument(_application, document);
            context.RequireFullAnalysis();
            var capability = _capabilityProvider.Evaluate(document);
            if (!capability.CanReadDocument) throw new InvalidOperationException(capability.Reason);
            if (capability.IsReadOnly || capability.IsProtected || capability.TrackChangesEnabled)
                throw new InvalidOperationException("Tài liệu phải cho phép chỉnh sửa, không bảo vệ và tắt Track Changes.");

            var rules = _accessManager.GetRulePack(LocalAccessManager.AutoFixFeature);
            var local = context.LastLocalSnapshot!;
            var formatFindings = context.LastFormatScan!.Findings;
            var spellingFindings = context.LastSpellingScan!.Findings;
            var backup = CreateBackup(document);
            var blocks = context.LastLogicalBlocks;
            var roles = context.LastRolesByParagraphIndex;
            var previousScreenUpdating = _application.ScreenUpdating;
            var previousAlerts = _application.DisplayAlerts;
            var undoStarted = false;
            var changedParagraphs = 0;
            var insertedLines = 0;
            var normalizedSections = 0;
            var normalizedTables = 0;
            var correctedSpellingItems = 0;
            var originalLegacyLineRules = FindLegacyHeaderSeparatorRules(local, roles);

            try
            {
                operation?.Transition(DocumentOperationState.Mutating,
                    "chuẩn hóa toàn bộ", cancellationEnabled: false);
                _application.ScreenUpdating = false;
                _application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                if (WordMajorVersion() >= 15)
                {
                    _application.UndoRecord.StartCustomRecord("Chuẩn hóa: toàn bộ văn bản");
                    undoStarted = true;
                }

                var annotations = new WordFindingAnnotationAdapter(_application, document);
                var removedAnnotations = annotations.ClearLane("format", operation);
                removedAnnotations += annotations.ClearLane("spelling", operation);

                // Line removal does not move text; paragraph splitting resolves its
                // paragraph by index and verifies the text, so both can run on the
                // prepared snapshot. A single fresh snapshot afterwards replaces the
                // up to three re-reads of earlier releases.
                var removedObsoleteLines = RemoveMisclassifiedOwnedLines(document, local, roles);
                var splitHeaders = SplitCombinedNationalHeaders(document, local);
                if (removedAnnotations > 0 || removedObsoleteLines > 0 || splitHeaders > 0)
                {
                    // Removing comment references can move subsequent absolute offsets
                    // and splitting changes paragraph indexes: never reuse the
                    // pre-annotation findings or anchors.
                    new WordDocumentReadRuntime(_application, _accessManager).Prepare(
                        context, DocumentAnalysisScope.Full, document, false, operation);
                    local = context.LastLocalSnapshot!;
                    formatFindings = context.LastFormatScan!.Findings;
                    spellingFindings = context.LastSpellingScan!.Findings;
                    blocks = context.LastLogicalBlocks;
                    roles = context.LastRolesByParagraphIndex;
                }

                // Paragraphs are located from snapshot offsets. Every text change made
                // below is recorded so later operations map their offsets to the
                // current document (no Paragraphs[i] lookups, which are O(n) each).
                s_tracker = new MainStoryEditTracker();
                operation?.ReportProgress(0, 4, "sửa chữ");
                correctedSpellingItems = ApplyDeterministicSpellingFixes(document, local,
                    formatFindings.Concat(spellingFindings).ToArray(), rules, blocks,
                    includeRepeatedSpaces: true);
                changedParagraphs += ApplyDeterministicParagraphFindingFixes(
                    document, local, formatFindings, roles);
                operation?.ReportProgress(1, 4, "định dạng đoạn");
                changedParagraphs += ApplyExplicitFormatFindingFixes(document, local,
                    formatFindings, rules, blocks, roles);
                normalizedSections = NormalizeSections(document, local, rules);
                NormalizeHeaderLayoutTables(document, local, roles, blocks);
                changedParagraphs += ApplyParagraphFormats(document, local, roles, rules, blocks);
                changedParagraphs += ApplyTechnicalDocumentProfile(document, local, blocks, roles);

                operation?.ReportProgress(2, 4, "bảng và số trang");
                EnsurePageNumbers(document, local, rules, blocks);
                normalizedTables = NormalizeTables(document, local, roles);
                WordAppendixPaginationNormalizer.Normalize(document, roles);
                RemoveTrailingBlankParagraphs(document);
                s_tracker = null;
                ReleaseComMemory();

                // Component lines are measured from the laid-out document. A technical
                // document (no administrative block) has no component lines at all, so
                // it skips this extra snapshot entirely.
                operation?.ReportProgress(3, 4, "đường kẻ thành phần");
                if (blocks.Any(block => block.IsAdministrative) ||
                    local.LineShapes.Any(line => LineShapeOwnership.IsOwned(line.Name)))
                {
                    try
                    {
                        document.Repaginate();
                        _application.ScreenRefresh();
                    }
                    catch (COMException) { }
                    new WordDocumentReadRuntime(_application, _accessManager).Prepare(
                        context, DocumentAnalysisScope.Full, document, false, operation);
                    local = context.LastLocalSnapshot!;
                    formatFindings = context.LastFormatScan!.Findings;
                    blocks = context.LastLogicalBlocks;
                    roles = context.LastRolesByParagraphIndex;
                    s_tracker = new MainStoryEditTracker();
                    insertedLines = InsertMissingRequiredLines(document, local, formatFindings,
                        blocks, roles, originalLegacyLineRules);
                    s_tracker = null;
                }

                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                    undoStarted = false;
                }

                context.ClearReadAnalysis();
                var readRuntime = new WordDocumentReadRuntime(_application, _accessManager);
                readRuntime.Prepare(context, DocumentAnalysisScope.Full, document, false, operation);
                var remaining = context.LastFormatScan!.Findings
                    .Concat(context.LastSpellingScan!.Findings).ToArray();
                return new OneClickResult(backup, changedParagraphs, insertedLines,
                    normalizedSections, normalizedTables, correctedSpellingItems, remaining);
            }
            catch (Exception exception)
            {
                if (undoStarted)
                {
                    try { _application.UndoRecord.EndCustomRecord(); }
                    catch (COMException) { }
                    undoStarted = false;
                }
                RestoreWorkingDocumentAfterFailure(document, backup, WordMajorVersion() >= 15);
                context.ClearReadAnalysis();
                throw new InvalidOperationException(exception.Message + "\nBản sao khôi phục: " + backup, exception);
            }
            finally
            {
                s_tracker = null;
                _application.DisplayAlerts = previousAlerts;
                _application.ScreenUpdating = previousScreenUpdating;
                ReleaseComMemory();
            }
        }

        private static void ReleaseComMemory()
        {
            // Release the RCWs of the finished phase so 32-bit Word does not keep
            // tens of thousands of Range/Font/ParagraphFormat proxies alive.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        /// <summary>
        /// Applies role formatting in document order. Consecutive plain body
        /// paragraphs with the same target size are formatted as one Range, so a
        /// long document needs a few hundred COM calls instead of ~15 per paragraph.
        /// </summary>
        private static int ApplyParagraphFormats(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles, LocalRulePack rules,
            IReadOnlyList<LogicalDocumentBlock> blocks)
        {
            var party = IsParty(snapshot);
            var changed = 0;
            var tiers = new Dictionary<int, Nd30HeaderFontSizeTier>();
            var run = new List<LocalParagraphSnapshot>();
            var runSize = 0f;
            foreach (var paragraph in snapshot.Paragraphs
                .Where(p => string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                .OrderBy(p => p.Index))
            {
                string role;
                roles.TryGetValue(paragraph.Index, out role);
                role = role ?? string.Empty;
                float size;
                if (IsPlainBodyParagraph(paragraph, role, snapshot, party, out size))
                {
                    if (run.Count > 0 && (run[run.Count - 1].Index + 1 != paragraph.Index ||
                        Math.Abs(runSize - size) > .01f))
                    {
                        changed += FormatBodyRun(document, snapshot, run, runSize, rules, party);
                        run.Clear();
                    }
                    run.Add(paragraph);
                    runSize = size;
                    continue;
                }
                if (run.Count > 0)
                {
                    changed += FormatBodyRun(document, snapshot, run, runSize, rules, party);
                    run.Clear();
                }
                if (string.IsNullOrWhiteSpace(paragraph.Text)) continue;
                var tier = role.Length == 0
                    ? UnusedHeaderTier
                    : ResolveHeaderTierCached(snapshot, blocks, paragraph.Index, tiers);
                if (ApplyParagraphFormat(document, paragraph, role, snapshot, rules, tier))
                    changed++;
            }
            if (run.Count > 0)
                changed += FormatBodyRun(document, snapshot, run, runSize, rules, party);
            return changed;
        }

        private static bool IsPlainBodyParagraph(LocalParagraphSnapshot paragraph, string role,
            LocalScanSnapshot snapshot, bool party, out float size)
        {
            size = party ? 14f : ValidOr(paragraph.FontSizePoints, 13, 14, 14);
            if (role.Length != 0 || paragraph.IsInTable || string.IsNullOrWhiteSpace(paragraph.Text))
                return false;
            if (IsLegalPointParagraph(paragraph.Text)) return false;
            if (Regex.IsMatch(paragraph.Text.Trim(), @"^Điều\s+\d+\s*[.:]",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200)))
                return false;
            if (snapshot.IntersectsProtectedSpan(paragraph)) return false;
            return IsBodyParagraph(paragraph);
        }

        private static int FormatBodyRun(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyList<LocalParagraphSnapshot> run, float size, LocalRulePack rules, bool party)
        {
            Word.Range? first = null;
            Word.Range? last = null;
            Word.Range? range = null;
            try
            {
                first = ResolveCurrentMainParagraphRange(document, run[0]);
                last = run.Count == 1 ? null : ResolveCurrentMainParagraphRange(document, run[run.Count - 1]);
                range = document.Range(first.Start, last == null ? first.End : last.End);
                ApplyBodyFormat(range, size, rules, party);
                return run.Count;
            }
            catch (Exception exception) when (exception is COMException ||
                exception is InvalidOperationException)
            {
                // Fall back to the per-paragraph path for this run only.
                var changed = 0;
                foreach (var paragraph in run)
                    if (ApplyParagraphFormat(document, paragraph, string.Empty, snapshot, rules,
                            UnusedHeaderTier))
                        changed++;
                return changed;
            }
            finally { Release(range); Release(last); Release(first); }
        }

        private static void ApplyBodyFormat(Word.Range range, float size, LocalRulePack rules, bool party)
        {
            Word.Font? font = null;
            Word.ParagraphFormat? format = null;
            try
            {
                font = range.Font;
                font.Name = rules.BodyFontName;
                font.Color = Word.WdColor.wdColorAutomatic;
                font.Size = size;
                format = range.ParagraphFormat;
                format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                format.RightIndent = 0f;
                format.LeftIndent = (float)(ParagraphIndentPolicy.BodyLeftMillimeters * PointsPerMillimeter);
                format.FirstLineIndent = (float)(ParagraphIndentPolicy.BodyFirstLineMillimeters * PointsPerMillimeter);
                format.SpaceBeforeAuto = 0;
                format.SpaceAfterAuto = 0;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 6f;
                // Widow/orphan control for body text (LATEX-PAGINATION-WIDOW).
                format.WidowControl = -1;
                if (party)
                {
                    format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    format.LineSpacing = 18f;
                }
                else
                {
                    format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    format.LineSpacing = 14.4f;
                }
            }
            finally { Release(format); Release(font); }
        }

        private static readonly Regex TableCaptionPattern = new Regex(
            @"^\s*Bảng\s+(?:số\s+)?(?:\d+(?:\.\d+)*|[IVXLCDM]+)(?:[.\-–—/:]\s*|\s+|$)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200));

        private static readonly Regex FigureCaptionPattern = new Regex(
            @"^\s*(?:Hình|Sơ\s+đồ|Biểu\s+đồ)\s+(?:số\s+)?(?:\d+(?:\.\d+)*|[IVXLCDM]+)(?:[.\-–—/:]\s*|\s+|$)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200));

        /// <summary>
        /// "TechnicalDocument" profile (Đợt 5) for blocks that are not
        /// administrative documents (D4): headings detected with the Vietnamese
        /// hierarchy keep with the next paragraph and their lines together and get
        /// consistent spacing (12/6 pt for level 1, 6/3 pt below); a table caption
        /// keeps with its table and a picture keeps with its caption below.
        /// Text is never changed.
        /// </summary>
        private static int ApplyTechnicalDocumentProfile(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyList<LogicalDocumentBlock> blocks, IReadOnlyDictionary<int, string> roles)
        {
            var technical = blocks.Where(block => !block.IsAdministrative).ToArray();
            if (technical.Length == 0) return 0;
            var main = snapshot.Paragraphs
                .Where(p => string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    technical.Any(block => block.ContainsParagraph(p.Index)))
                .OrderBy(p => p.Index)
                .ToList();
            if (main.Count == 0) return 0;
            var byIndex = main.GroupBy(p => p.Index).ToDictionary(group => group.Key, group => group.First());
            var blockIds = new Dictionary<int, string>();
            foreach (var paragraph in main)
            {
                var owner = technical.First(block => block.ContainsParagraph(paragraph.Index));
                blockIds[paragraph.Index] = "B" + owner.StartParagraphIndex.ToString(CultureInfo.InvariantCulture);
            }

            var changed = 0;
            var headings = new HeadingDetector().Detect(main, new HeadingDetectionContext(roles, blockIds));
            foreach (var heading in headings)
            {
                LocalParagraphSnapshot paragraph;
                if (!byIndex.TryGetValue(heading.ParagraphIndex, out paragraph) ||
                    snapshot.IntersectsProtectedSpan(paragraph)) continue;
                var first = heading.Level <= 1;
                if (SetKeepAndSpacing(document, paragraph, true, true,
                        first ? 12f : 6f, first ? 6f : 3f))
                    changed++;
            }

            var pictureParagraphs = new HashSet<int>(snapshot.GraphicObjects
                .Where(item => item.IsInline && item.AnchorParagraphIndex.HasValue &&
                    string.Equals(item.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                .Select(item => item.AnchorParagraphIndex!.Value));
            for (var position = 0; position < main.Count; position++)
            {
                var paragraph = main[position];
                if (paragraph.IsInTable) continue;
                var text = paragraph.Text ?? string.Empty;
                if (TableCaptionPattern.IsMatch(text))
                {
                    // Caption above its table: never leave it alone at a page end.
                    if (SetKeepAndSpacing(document, paragraph, true, true, null, null)) changed++;
                }
                else if (FigureCaptionPattern.IsMatch(text) && position > 0 &&
                    pictureParagraphs.Contains(main[position - 1].Index) &&
                    main[position - 1].Index + 1 == paragraph.Index)
                {
                    // Picture above its caption: keep the picture with the caption.
                    if (SetKeepAndSpacing(document, main[position - 1], true, false, null, null)) changed++;
                }
            }
            return changed;
        }

        private static bool SetKeepAndSpacing(Word.Document document, LocalParagraphSnapshot paragraph,
            bool keepWithNext, bool keepTogether, float? spaceBefore, float? spaceAfter)
        {
            Word.Range? range = null;
            Word.ParagraphFormat? format = null;
            try
            {
                range = ResolveCurrentMainParagraphRange(document, paragraph);
                format = range.ParagraphFormat;
                if (keepWithNext) format.KeepWithNext = -1;
                if (keepTogether) format.KeepTogether = -1;
                if (spaceBefore.HasValue)
                {
                    format.SpaceBeforeAuto = 0;
                    format.SpaceBefore = spaceBefore.Value;
                }
                if (spaceAfter.HasValue)
                {
                    format.SpaceAfterAuto = 0;
                    format.SpaceAfter = spaceAfter.Value;
                }
                return true;
            }
            catch (Exception exception) when (exception is COMException ||
                exception is InvalidOperationException)
            {
                return false;
            }
            finally { Release(format); Release(range); }
        }

        public SelectedFindingFixResult ExecuteSelectedFinding(
            DocumentContext context,
            Word.Document? activeDocument = null)
        {
            var document = activeDocument ?? _application.ActiveDocument;
            var annotations = new WordFindingAnnotationAdapter(_application, document);
            string selectedLane;
            string selectedFindingId;
            if (!annotations.TryGetSelectedFinding(out selectedLane, out selectedFindingId))
                throw new InvalidOperationException(
                    "Hãy bấm vào phần văn bản đang được comment/tô đỏ bởi Chuẩn hóa rồi bấm lại Sửa lỗi đang chọn.");
            string selectedStory;
            int selectedStart;
            int selectedEnd;
            if (!annotations.TryGetSelectedDocumentRange(out selectedStory, out selectedStart, out selectedEnd))
                throw new InvalidOperationException(
                    "Hãy bấm vào phần văn bản đang được comment/tô đỏ bởi Chuẩn hóa rồi bấm lại Sửa lỗi đang chọn.");
            return ExecuteSelectedFinding(context, selectedLane, selectedFindingId,
                selectedStory, selectedStart, selectedEnd, document);
        }

        public SelectedFindingFixResult ExecuteSelectedFinding(
            DocumentContext context,
            string selectedLane,
            string selectedFindingId,
            string selectedStory,
            int selectedStart,
            int selectedEnd,
            Word.Document? activeDocument = null,
            DocumentOperationSession? operation = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            var document = activeDocument ?? _application.ActiveDocument;
            var capability = _capabilityProvider.Evaluate(document);
            if (!capability.CanReadDocument) throw new InvalidOperationException(capability.Reason);
            if (capability.IsReadOnly || capability.IsProtected || capability.TrackChangesEnabled)
                throw new InvalidOperationException("Tài liệu phải cho phép chỉnh sửa, không bảo vệ và tắt Track Changes.");

            context.RequireSnapshotAnalysis();
            var annotations = new WordFindingAnnotationAdapter(_application, document);
            if (string.IsNullOrWhiteSpace(selectedStory) || selectedStart < 0 || selectedEnd < selectedStart)
                throw new ArgumentException("Vị trí lỗi đang chọn không hợp lệ.", nameof(selectedStory));

            var rules = _accessManager.GetRulePack(LocalAccessManager.AutoFixFeature);
            var local = context.LastLocalSnapshot!;
            LocalScanResult selectedScan;
            if (string.Equals(selectedLane, "spelling", StringComparison.OrdinalIgnoreCase))
            {
                context.RequireSpellingAnalysis();
                selectedScan = context.LastSpellingScan!;
            }
            else if (string.Equals(selectedLane, "format", StringComparison.OrdinalIgnoreCase))
            {
                context.RequireFormatAnalysis();
                selectedScan = context.LastFormatScan!;
            }
            else
            {
                throw new InvalidOperationException("Comment đang chọn không thuộc nhóm lỗi Chuẩn hóa hỗ trợ.");
            }

            var finding = selectedScan.Findings.FirstOrDefault(item =>
                string.Equals(item.FindingId, selectedFindingId, StringComparison.Ordinal));
            if (finding == null)
            {
                finding = selectedScan.Findings
                    .Select(item => new
                {
                    Finding = item,
                    Range = FindingRange(item, local)
                })
                .Where(item => item.Range != null &&
                    string.Equals(item.Range.Item1, selectedStory, StringComparison.Ordinal) &&
                    NumericRangesTouch(item.Range.Item2, item.Range.Item3, selectedStart, selectedEnd))
                .OrderBy(item => item.Range!.Item3 - item.Range.Item2)
                    .Select(item => item.Finding)
                    .FirstOrDefault();
            }
            if (finding == null)
                throw new InvalidOperationException(
                    "Lỗi đang chọn không còn tồn tại trong nội dung hiện tại. Hãy chạy lại chức năng kiểm tra tương ứng.");
            var lane = selectedScan.Lane;
            var findingId = finding.FindingId;

            var previousScreenUpdating = _application.ScreenUpdating;
            var undoStarted = false;
            try
            {
                operation?.Transition(DocumentOperationState.Mutating,
                    "sửa lỗi đang chọn", cancellationEnabled: false);
                _application.ScreenUpdating = false;
                if (WordMajorVersion() >= 15)
                {
                    _application.UndoRecord.StartCustomRecord("Chuẩn hóa: sửa lỗi đang chọn");
                    undoStarted = true;
                }

                var combinedParagraph = local.Paragraphs.FirstOrDefault(p =>
                    p.Index == finding.Anchor.ParagraphIndex &&
                    CombinedNationalHeader.GetBreakOffsets(p.Text).Length > 0);
                var applied = lane == "format" && finding.RuleCode == "ND30-PL1-M2-K1-TN-LINE" &&
                    combinedParagraph != null
                    ? FixCombinedHeaderFinding(document, context, local, combinedParagraph, rules)
                    : string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase)
                    ? ApplyDeterministicSpellingFixes(document, local, new[] { finding }, rules,
                        context.LastLogicalBlocks) > 0
                    : ApplySelectedFormatFix(document, local, finding, rules,
                        context.LastLogicalBlocks, context.LastRolesByParagraphIndex);

                if (!applied && string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase))
                {
                    Word.Range? directRange = null;
                    try
                    {
                        directRange = document.Range(selectedStart, selectedEnd);
                        var directText = directRange.Text ?? string.Empty;
                        var target = ResolveTargetReplacement(finding, rules);
                        var expected = finding.Anchor.ExpectedText ?? string.Empty;
                        if (target != null &&
                            (string.Equals(directText, expected, StringComparison.Ordinal) ||
                             string.Equals(directText.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase) ||
                             (finding.RuleCode == "LOCAL-TYPO-SPACE" && Regex.IsMatch(directText, @"^[ \t]{2,}$")) ||
                             (finding.RuleCode == "LOCAL-TYPO-PUNCT" && string.IsNullOrWhiteSpace(directText))))
                        {
                            directRange.Text = target;
                            applied = true;
                        }
                    }
                    catch (COMException) { }
                    finally { Release(directRange); }
                }

                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                    undoStarted = false;
                }

                if (!applied)
                {
                    // A report-only finding is not an exceptional failure and must not
                    // call Document.Undo: doing so could undo the user's preceding edit.
                    // Keep its annotation and cached finding exactly as they are.
                    return new SelectedFindingFixResult(string.Empty, lane, findingId, false);
                }

                // ApplySelectedFormatFix and the spelling edit path both verify their
                // targeted Word mutation before returning true. Remove only the owned
                // annotation at that coordinate. A selected repair must never perform
                // a hidden whole-document snapshot or scan.
                annotations.ClearOwnedAnnotation(lane, findingId, selectedStory,
                    selectedStart, selectedEnd);
                UpdateContextAfterTargetedFix(context, lane, findingId,
                    preserveSnapshot: !SelectedFixChangesText(lane, finding));
                return new SelectedFindingFixResult(string.Empty, lane, findingId, true);
            }
            catch (Exception exception)
            {
                if (undoStarted)
                {
                    try { _application.UndoRecord.EndCustomRecord(); }
                    catch (COMException) { }
                    undoStarted = false;
                }
                try
                {
                    object count = 1;
                    document.Undo(ref count);
                }
                catch (COMException) { }
                context.ClearReadAnalysis();
                throw new InvalidOperationException(exception.Message, exception);
            }
            finally
            {
                _application.ScreenUpdating = previousScreenUpdating;
            }
        }

        public int FixAllSpellingFindings(
            DocumentContext context,
            Word.Document? activeDocument = null,
            DocumentOperationSession? operation = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            var document = activeDocument ?? _application.ActiveDocument;
            var capability = _capabilityProvider.Evaluate(document);
            if (!capability.CanReadDocument) throw new InvalidOperationException(capability.Reason);
            if (capability.IsReadOnly || capability.IsProtected || capability.TrackChangesEnabled)
                throw new InvalidOperationException("Tài liệu phải cho phép chỉnh sửa, không bảo vệ và tắt Track Changes.");

            context.RequireSnapshotAnalysis();
            context.RequireSpellingAnalysis();

            var local = context.LastLocalSnapshot!;
            var scan = context.LastSpellingScan!;
            var findings = scan.Findings;

            var rules = _accessManager.GetRulePack(LocalAccessManager.AutoFixFeature);
            var annotations = new WordFindingAnnotationAdapter(_application, document);
            var previousScreenUpdating = _application.ScreenUpdating;
            var undoStarted = false;

            try
            {
                operation?.Transition(DocumentOperationState.Mutating,
                    "sửa nhanh chính tả", cancellationEnabled: false);
                _application.ScreenUpdating = false;
                if (WordMajorVersion() >= 15)
                {
                    _application.UndoRecord.StartCustomRecord("Chuẩn hóa: sửa nhanh chính tả");
                    undoStarted = true;
                }

                var fixedCount = 0;
                if (findings.Count > 0)
                {
                    fixedCount += ApplyDeterministicSpellingFixes(document, local, findings, rules,
                        context.LastLogicalBlocks);
                }

                fixedCount += CleanTypographyAndQuotes(document);

                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                    undoStarted = false;
                }

                context.ClearReadAnalysis();
                return fixedCount;
            }
            catch (Exception exception)
            {
                if (undoStarted)
                {
                    try { _application.UndoRecord.EndCustomRecord(); }
                    catch (COMException) { }
                    undoStarted = false;
                }
                try
                {
                    object count = 1;
                    document.Undo(ref count);
                }
                catch (COMException) { }
                context.ClearReadAnalysis();
                throw new InvalidOperationException(exception.Message, exception);
            }
            finally
            {
                _application.ScreenUpdating = previousScreenUpdating;
            }
        }

        private static int CleanTypographyAndQuotes(Word.Document document)
        {
            var modifiedCount = 0;
            foreach (var story in EditableStories(document))
            {
                try
                {
                    foreach (Word.Paragraph paragraph in story.Paragraphs)
                    {
                        Word.Range? range = null;
                        try
                        {
                            range = paragraph.Range.Duplicate;
                            var source = range.Text ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(source)) continue;

                            var cleaned = VietnameseTypographyCleaner.CleanWhitespaceAndPunctuation(source);
                            cleaned = VietnameseTypographyCleaner.NormalizeQuotationMarks(cleaned);
                            if (!string.Equals(source, cleaned, StringComparison.Ordinal))
                            {
                                range.Text = cleaned;
                                modifiedCount++;
                            }
                        }
                        finally
                        {
                            Release(range);
                            Release(paragraph);
                        }
                    }
                }
                catch (COMException) { }
                finally
                {
                    Release(story);
                }
            }
            return modifiedCount;
        }

        private static IEnumerable<Word.Range> EditableStories(Word.Document document)
        {
            var types = new[]
            {
                Word.WdStoryType.wdMainTextStory,
                Word.WdStoryType.wdFootnotesStory,
                Word.WdStoryType.wdEndnotesStory,
                Word.WdStoryType.wdPrimaryHeaderStory,
                Word.WdStoryType.wdPrimaryFooterStory,
                Word.WdStoryType.wdEvenPagesHeaderStory,
                Word.WdStoryType.wdEvenPagesFooterStory,
                Word.WdStoryType.wdFirstPageHeaderStory,
                Word.WdStoryType.wdFirstPageFooterStory,
                Word.WdStoryType.wdTextFrameStory
            };
            foreach (var type in types)
            {
                Word.Range? current = null;
                try { current = document.StoryRanges[type]; }
                catch (COMException) { }
                while (current != null)
                {
                    Word.Range? next = null;
                    try { next = current.NextStoryRange; }
                    catch (COMException) { }
                    yield return current;
                    current = next;
                }
            }
        }

        private static Tuple<string, int, int>? FindingRange(AnnotationFinding finding,
            LocalScanSnapshot snapshot)
        {
            if (!finding.Anchor.ParagraphIndex.HasValue) return null;
            var paragraph = snapshot.Paragraphs.FirstOrDefault(item =>
                item.Index == finding.Anchor.ParagraphIndex.Value);
            if (paragraph == null) return null;
            var start = paragraph.AbsoluteStart + finding.Anchor.StartOffset.GetValueOrDefault();
            var length = finding.Anchor.Length ?? paragraph.Text.Length;
            return Tuple.Create(paragraph.StoryType, start, checked(start + Math.Max(0, length)));
        }

        private static bool NumericRangesTouch(int leftStart, int leftEnd, int rightStart, int rightEnd)
        {
            if (leftStart == leftEnd) return rightStart <= leftStart && leftStart <= rightEnd;
            if (rightStart == rightEnd) return leftStart <= rightStart && rightStart <= leftEnd;
            return leftStart < rightEnd && rightStart < leftEnd;
        }

        private static bool ApplySelectedFormatFix(Word.Document document, LocalScanSnapshot snapshot,
            AnnotationFinding finding, LocalRulePack rules,
            IReadOnlyList<LogicalDocumentBlock> blocks,
            IReadOnlyDictionary<int, string> roles)
        {
            if (IsExplicitReportOnlyFinding(finding)) return false;
            if (!finding.Anchor.ParagraphIndex.HasValue &&
                finding.Anchor.SectionIndex.HasValue)
            {
                var sectionIndex = finding.Anchor.SectionIndex.Value;
                if (sectionIndex < 1 || sectionIndex > document.Sections.Count) return false;
                Word.Section? section = null;
                try
                {
                    section = document.Sections[sectionIndex];
                    if (finding.RuleCode == "ND30-PL1-M1-K1" ||
                        finding.RuleCode == "ND30-PL1-M1-K3")
                    {
                        NormalizeSection(section, IsParty(snapshot), rules);
                        return true;
                    }
                    if (finding.RuleCode == "ND30-PL1-M1-K7")
                    {
                        EnsurePageNumber(document, section, IsParty(snapshot), rules);
                        return ContainsPageField(section);
                    }
                }
                finally { Release(section); }
                return false;
            }
            // This is classified as a format finding, but its safe correction is a
            // deterministic text insertion. Reuse the same guarded edit path as
            // 1-Click so the comment is removed only after "số" was actually added.
            if (finding.Anchor.Kind == AnnotationAnchorKind.TextSpan ||
                IsDeterministicComponentTextFinding(finding.RuleCode))
            {
                var textApplied = ApplyDeterministicSpellingFixes(document, snapshot,
                    new[] { finding }, rules, blocks) > 0;
                if (textApplied) return true;
            }
            if (finding.RuleCode.StartsWith("LATEX-", StringComparison.Ordinal))
            {
                if (!rules.AcademicTypography.IsAutoFixEnabled(finding.RuleCode) ||
                    !finding.Anchor.ParagraphIndex.HasValue) return false;
                var advisoryParagraph = snapshot.Paragraphs.FirstOrDefault(item =>
                    item.Index == finding.Anchor.ParagraphIndex.Value);
                if (advisoryParagraph == null || snapshot.IntersectsProtectedSpan(advisoryParagraph) ||
                    advisoryParagraph.IsInTable ||
                    !string.Equals(advisoryParagraph.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                    return false;
                Word.Range? advisoryRange = null;
                try
                {
                    advisoryRange = ResolveCurrentMainParagraphRange(document, advisoryParagraph);
                    if (string.Equals(finding.RuleCode, AcademicTypographyRuleCodes.PaginationKeep,
                        StringComparison.Ordinal))
                    {
                        advisoryRange.ParagraphFormat.KeepWithNext = -1;
                        return advisoryRange.ParagraphFormat.KeepWithNext != 0;
                    }
                    if (string.Equals(finding.RuleCode, AcademicTypographyRuleCodes.PaginationWidow,
                        StringComparison.Ordinal))
                    {
                        advisoryRange.ParagraphFormat.WidowControl = -1;
                        return advisoryRange.ParagraphFormat.WidowControl != 0;
                    }
                    return false;
                }
                finally { Release(advisoryRange); }
            }
            if (!finding.Anchor.ParagraphIndex.HasValue) return false;
            var paragraph = snapshot.Paragraphs.FirstOrDefault(item =>
                item.Index == finding.Anchor.ParagraphIndex.Value);
            if (paragraph == null || !string.Equals(paragraph.StoryType, "wdMainTextStory",
                    StringComparison.Ordinal))
                return false;

            if (finding.RuleCode.EndsWith("-LINE", StringComparison.Ordinal))
            {
                var ratio = finding.RuleCode == "ND30-PL1-M2-K2-ORG-LINE" ||
                    finding.RuleCode == "ND30-PL1-M2-K5A-SUBJ-LINE" ? .4d : 1d;
                return NormalizeRequiredLine(document, snapshot, paragraph, ratio, finding.RuleCode);
            }

            if (ApplyDeterministicParagraphFindingFixes(document, snapshot,
                    new[] { finding }, roles) > 0)
                return true;

            var headerFontTier = ResolveHeaderTier(snapshot, blocks, paragraph.Index);
            string role;
            roles.TryGetValue(paragraph.Index, out role);
            string explicitRole;
            if (TryResolveExplicitFindingRole(finding.RuleCode, role, out explicitRole))
                role = explicitRole;
            return ApplyParagraphFormat(document, paragraph, role ?? string.Empty, snapshot, rules,
                headerFontTier);
        }

        private static int ApplyExplicitFormatFindingFixes(
            Word.Document document,
            LocalScanSnapshot snapshot,
            IReadOnlyList<AnnotationFinding> findings,
            LocalRulePack rules,
            IReadOnlyList<LogicalDocumentBlock> blocks,
            IReadOnlyDictionary<int, string> roles)
        {
            var paragraphs = snapshot.Paragraphs.ToDictionary(item => item.Index);
            var handled = new HashSet<int>();
            var changed = 0;
            foreach (var finding in findings)
            {
                if (!finding.Anchor.ParagraphIndex.HasValue || IsExplicitReportOnlyFinding(finding))
                    continue;
                var paragraphIndex = finding.Anchor.ParagraphIndex.Value;
                LocalParagraphSnapshot paragraph;
                string detectedRole;
                roles.TryGetValue(paragraphIndex, out detectedRole);
                string role;
                if (!TryResolveExplicitFindingRole(finding.RuleCode, detectedRole, out role) ||
                    !handled.Add(paragraphIndex) ||
                    !paragraphs.TryGetValue(paragraphIndex, out paragraph) ||
                    snapshot.IntersectsProtectedSpan(paragraph))
                    continue;
                var headerFontTier = ResolveHeaderTier(snapshot, blocks, paragraphIndex);
                if (ApplyParagraphFormat(document, paragraph, role, snapshot, rules, headerFontTier))
                    changed++;
            }
            return changed;
        }

        private static bool TryResolveExplicitFindingRole(string ruleCode, string? detectedRole,
            out string role)
        {
            switch (ruleCode)
            {
                case "ND30-PL1-M2-K1-QH": role = "nationalTitle"; return true;
                case "ND30-PL1-M2-K1-TN": role = "nationalMotto"; return true;
                case "ND30-PL1-M2-K2-ORG": role = "organName"; return true;
                case "ND30-PL1-M2-K2-SUP": role = "superiorOrganName"; return true;
                case "ND30-PL1-M2-K4-STYLE": role = "placeAndIssuedDate"; return true;
                case "ND30-PL1-M2-K5A-SUBJ": role = "subject"; return true;
                case "ND30-PL1-M2-K5A-TYPE": role = "typeName"; return true;
                case "ND30-PL1-M2-K5B-STYLE": role = "officialLetterSubject"; return true;
                case "ND30-PL1-M2-K6A-STYLE": role = "legalBasis"; return true;
                case "ND30-PL1-M2-K7D-STYLE": role = "signerAuthority"; return true;
                case "ND30-PL1-M1-K4-FONT":
                case "ND30-PL1-M1-K4-COLOR":
                case "ND30-PL1-M2-K6E-ALIGN":
                case "ND30-PL1-M2-K6E-INDENT":
                case "ND30-PL1-M2-K6E-LINESPACING":
                case "ND30-PL1-M2-K6E-SPACEAFTER":
                    role = string.Empty; return true;
                case "ND30-PL1-M2-K6D-POINT":
                    role = string.Empty; return true;
                case "ND30-PL1-M3-K1B":
                    role = detectedRole == "appendixTitle" ? "appendixTitle" : "appendixLabel";
                    return true;
                case "ND30-PL1-M3-K1C":
                    if (detectedRole == "appendixReference" ||
                        detectedRole == "appendixDigitalSignatureInfo")
                    {
                        role = detectedRole!;
                        return true;
                    }
                    break;
                case "ND30-PL1-MV-CT1":
                    role = detectedRole ?? string.Empty; return true;
                case "ND30-PL1-M2-K9A-LAYOUT":
                case "ND30-PL1-M2-K9B-LABEL":
                case "ND30-PL1-M2-K9B-LIST":
                    if (!string.IsNullOrWhiteSpace(detectedRole))
                    {
                        role = detectedRole!;
                        return true;
                    }
                    break;
            }
            role = string.Empty;
            return false;
        }

        private static bool IsExplicitReportOnlyFinding(AnnotationFinding finding)
        {
            switch (finding.RuleCode)
            {
                case "ND30-PL1-M2-K6B-CITE":
                case "ND30-PL1-M2-K6D-ALPHABET":
                case "ND30-PL1-M2-K6D-TITLE":
                case "ND30-PL1-M3-K1A-REF":
                case "ND30-PL1-M3-K1D":
                    return true;
                case "ND30-PL1-M2-K9A-LAYOUT":
                    return finding.CurrentIssue.IndexOf("tách dòng nhưng không có danh sách",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                case "ND30-PL1-M2-K6D-ARTICLE":
                    return finding.CurrentIssue.IndexOf("số điều không liên tục",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                case "ND30-PL1-M3-K1C":
                    return finding.CurrentIssue.IndexOf("thiếu dòng",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                case "ND30-PL1-M2-K1-QH":
                    // "TÊN CƠ QUAN<TAB>CỘNG HÒA..." on one line is reported only.
                    return finding.CurrentIssue.IndexOf("chung một dòng",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                default:
                    return false;
            }
        }

        private static int ApplyDeterministicParagraphFindingFixes(
            Word.Document document,
            LocalScanSnapshot snapshot,
            IReadOnlyList<AnnotationFinding> findings,
            IReadOnlyDictionary<int, string> roles)
        {
            var paragraphs = snapshot.Paragraphs.ToDictionary(item => item.Index);
            var handled = new HashSet<string>(StringComparer.Ordinal);
            var changed = 0;
            foreach (var finding in findings)
            {
                if (!finding.Anchor.ParagraphIndex.HasValue) continue;
                var paragraphIndex = finding.Anchor.ParagraphIndex.Value;
                var key = finding.RuleCode + "\u001f" + paragraphIndex.ToString(CultureInfo.InvariantCulture);
                if (!handled.Add(key)) continue;
                LocalParagraphSnapshot paragraph;
                if (!paragraphs.TryGetValue(paragraphIndex, out paragraph) ||
                    paragraph.IsInTable || snapshot.IntersectsProtectedSpan(paragraph) ||
                    !string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                    continue;

                Word.Range? range = null;
                try
                {
                    range = ResolveCurrentMainParagraphRange(document, paragraph);
                    switch (finding.RuleCode)
                    {
                        case "ND30-PL1-M2-K1-C":
                            range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                            if (Math.Abs(range.ParagraphFormat.SpaceBefore) > .1f)
                            {
                                range.ParagraphFormat.SpaceBefore = 0f;
                                changed++;
                            }
                            break;
                        case "ND30-PL1-M2-K5B-SPACE":
                            if (TryReplaceParagraphText(range,
                                    NormalizeOfficialLetterSubject(ReadPrintableText(range), IsParty(snapshot))))
                                changed++;
                            break;
                        case "ND30-PL1-M2-K6A-STYLE":
                            if (IsParty(snapshot) && finding.CurrentIssue.IndexOf("thiếu gạch ngang",
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                var legalBasis = ReadPrintableText(range).TrimStart();
                                if (!Regex.IsMatch(legalBasis, @"^[-–—]\s*"))
                                {
                                    if (TryReplaceParagraphText(range, "- " + legalBasis)) changed++;
                                }
                            }
                            break;
                        case "ND30-PL1-M2-K6D-ARTICLE":
                            if (ApplyLegalArticleFormat(range)) changed++;
                            break;
                        case "ND30-PL1-M2-K6D-CLAUSE":
                            if (range.Font.Italic != 0)
                            {
                                range.Font.Italic = 0;
                                changed++;
                            }
                            break;
                        case "ND30-PL1-M2-K9A-COLON":
                            var salutation = ReadPrintableText(range);
                            var normalizedSalutation = Regex.Replace(salutation,
                                @"^(\s*Kính\s+(?:gửi|trình))\s*:?[ \t]*",
                                "$1: ", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                                TimeSpan.FromMilliseconds(200));
                            if (TryReplaceParagraphText(range, normalizedSalutation)) changed++;
                            break;
                        case "ND30-PL1-M2-K9A-INLINE-END":
                            if (TryReplaceParagraphText(range,
                                    EnsureEndingPunctuation(ReadPrintableText(range), "."))) changed++;
                            break;
                        case "ND30-PL1-M2-K9A-PUNCT":
                            var salutationItems = roles.Where(item => item.Value == "recipientSalutationList")
                                .Select(item => item.Key).OrderBy(item => item).ToArray();
                            var salutationEnd = salutationItems.Length > 0 &&
                                salutationItems[salutationItems.Length - 1] == paragraphIndex ? "." : ";";
                            if (TryReplaceParagraphText(range,
                                    NormalizeDashListItem(ReadPrintableText(range), salutationEnd))) changed++;
                            break;
                        case "ND30-PL1-M2-K9B-LABEL":
                            if (finding.CurrentIssue.IndexOf("thiếu dấu hai chấm",
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                var label = ReadPrintableText(range).TrimEnd();
                                if (TryReplaceParagraphText(range, label.TrimEnd(':') + ":")) changed++;
                            }
                            break;
                        case "ND30-PL1-M2-K9B-LIST":
                            var recipientItems = roles.Where(item => item.Value == "recipientList")
                                .Select(item => item.Key).OrderBy(item => item).ToArray();
                            var recipientEnd = recipientItems.Length > 0 &&
                                recipientItems[recipientItems.Length - 1] == paragraphIndex ? "." : ";";
                            if (TryReplaceParagraphText(range,
                                NormalizeDashListItem(ReadPrintableText(range), recipientEnd))) changed++;
                            break;
                        case "ND30-PL1-M2-K9B-LUU":
                            var storageLine = NormalizeStorageLine(ReadPrintableText(range));
                            if (storageLine != null && TryReplaceParagraphText(range, storageLine)) changed++;
                            break;
                        case "ND30-PL1-M3-K1A-NUM":
                            var appendixLabels = roles.Where(item => item.Value == "appendixLabel")
                                .Select(item => item.Key).OrderBy(item => item).ToArray();
                            var appendixOrdinal = Array.IndexOf(appendixLabels, paragraphIndex) + 1;
                            var numberedLabel = NormalizeAppendixLabelNumber(
                                ReadPrintableText(range), appendixOrdinal);
                            if (numberedLabel != null &&
                                TryReplaceParagraphText(range, numberedLabel)) changed++;
                            break;
                        case "ND30-PL1-M3-K1B":
                            if (finding.CurrentIssue.IndexOf("chưa viết hoa",
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                var appendixTitle = ReadPrintableText(range);
                                if (TryReplaceParagraphText(range, appendixTitle.ToUpper(
                                        CultureInfo.GetCultureInfo("vi-VN")))) changed++;
                            }
                            break;
                    }
                }
                catch (COMException)
                {
                }
                finally { Release(range); }
            }
            return changed;
        }

        private static bool ApplyLegalArticleFormat(Word.Range paragraphRange)
        {
            var text = ReadPrintableText(paragraphRange);
            var marker = Regex.Match(text, @"^\s*Điều\s+\d+\s*\.?",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
            if (!marker.Success) return false;

            Word.Range? printable = null;
            Word.Range? markerRange = null;
            Word.Range? contentRange = null;
            try
            {
                paragraphRange.ParagraphFormat.LeftIndent = 0f;
                paragraphRange.ParagraphFormat.FirstLineIndent = 10f * PointsPerMillimeter;

                printable = paragraphRange.Duplicate;
                printable.End = printable.Start + text.Length;
                printable.Font.Italic = 0;
                printable.Font.Bold = 0;

                markerRange = paragraphRange.Duplicate;
                markerRange.End = markerRange.Start + marker.Length;
                markerRange.Font.Bold = -1;

                if (marker.Length < text.Length)
                {
                    contentRange = paragraphRange.Duplicate;
                    contentRange.Start += marker.Length;
                    contentRange.End = paragraphRange.Start + text.Length;
                    contentRange.Font.Bold = 0;
                }

                return markerRange.Font.Bold != 0 &&
                    (contentRange == null || contentRange.Font.Bold == 0);
            }
            finally
            {
                Release(contentRange);
                Release(markerRange);
                Release(printable);
            }
        }

        private static string? NormalizeStorageLine(string value)
        {
            var match = Regex.Match(value ?? string.Empty,
                @"^\s*[-–—]?\s*Lưu\s*:?\s*(?<content>.+?)\s*$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
            if (!match.Success) return null;
            var content = match.Groups["content"].Value.Trim();
            return content.Length == 0 ? null : "- Lưu: " + content;
        }

        private static string? NormalizeAppendixLabelNumber(string value, int ordinal)
        {
            if (ordinal <= 0) return null;
            var match = Regex.Match(value ?? string.Empty,
                @"^\s*Phụ\s+lục(?:\s+[A-Za-zĐđ])?\s*$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
            return match.Success
                ? "Phụ lục " + ordinal.ToString(CultureInfo.InvariantCulture)
                : null;
        }

        private static string NormalizeOfficialLetterSubject(string value, bool party)
        {
            var text = Regex.Replace(value ?? string.Empty, @"[ \t]+", " ").Trim();
            if (party)
            {
                text = Regex.Replace(text, @"^(?:V/v|về\s+việc)\s*", string.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                    TimeSpan.FromMilliseconds(200));
                return "về việc " + text;
            }
            var match = Regex.Match(text, @"^(?:V/v|Về\s+việc)\s*",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
            return match.Success ? "V/v " + text.Substring(match.Length).TrimStart() : "V/v " + text;
        }

        private static string NormalizeDashListItem(string value, string ending)
        {
            var text = Regex.Replace(value ?? string.Empty, @"^\s*[-–—]?\s*", string.Empty,
                RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200)).Trim();
            text = text.TrimEnd('.', ';', ',', ':', ' ');
            return "- " + text + ending;
        }

        private static string EnsureEndingPunctuation(string value, string ending)
        {
            var text = (value ?? string.Empty).TrimEnd();
            return text.TrimEnd('.', ';', ',', ':', '!', '?', ' ') + ending;
        }

        private static string ReadPrintableText(Word.Range paragraphRange)
        {
            var value = paragraphRange.Text ?? string.Empty;
            return value.TrimEnd('\r', '\a');
        }

        private static bool TryReplaceParagraphText(Word.Range paragraphRange, string replacement)
        {
            Word.Range? textRange = null;
            try
            {
                textRange = paragraphRange.Duplicate;
                var existing = ReadPrintableText(textRange);
                if (string.Equals(existing, replacement, StringComparison.Ordinal)) return false;
                var start = textRange.Start;
                textRange.End = textRange.Start + existing.Length;
                textRange.Text = replacement;
                RecordTextEdit(start, existing.Length, replacement.Length);
                paragraphRange.SetRange(start, checked(start + replacement.Length));
                return string.Equals(paragraphRange.Text ?? string.Empty, replacement,
                    StringComparison.Ordinal);
            }
            finally { Release(textRange); }
        }

        private static int NormalizeSections(Word.Document document, LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            var party = IsParty(snapshot);
            var count = 0;
            foreach (Word.Section section in document.Sections)
            {
                try
                {
                    NormalizeSection(section, party, rules);
                    count++;
                }
                finally { Release(section); }
            }
            return count;
        }

        private static void NormalizeSection(Word.Section section, bool party, LocalRulePack rules)
        {
            Word.PageSetup? setup = null;
            try
            {
                setup = section.PageSetup;
                var top = (float)((party ? 20d : rules.TopMinMm) * PointsPerMillimeter);
                var bottom = (float)((party ? 20d : rules.BottomMinMm) * PointsPerMillimeter);
                var left = (float)((party ? 30d : rules.LeftMinMm) * PointsPerMillimeter);
                var right = (float)((party ? 15d : rules.RightMinMm) * PointsPerMillimeter);
                if (setup.Orientation == Word.WdOrientation.wdOrientLandscape)
                {
                    // Never flip the orientation chosen by the author. A landscape page
                    // is bound on the same physical edge as the portrait pages once it
                    // is rotated for reading: the binding margin is at the top.
                    setup.PageWidth = 297f * PointsPerMillimeter;
                    setup.PageHeight = 210f * PointsPerMillimeter;
                    setup.TopMargin = left;
                    setup.BottomMargin = right;
                    setup.LeftMargin = top;
                    setup.RightMargin = bottom;
                    return;
                }
                setup.PaperSize = Word.WdPaperSize.wdPaperA4;
                setup.TopMargin = top;
                setup.BottomMargin = bottom;
                setup.LeftMargin = left;
                setup.RightMargin = right;
            }
            finally { Release(setup); }
        }

        private static bool ApplyParagraphFormat(Word.Document document, LocalParagraphSnapshot paragraph,
            string role, LocalScanSnapshot snapshot, LocalRulePack rules,
            Nd30HeaderFontSizeTier headerFontTier)
        {
            if (string.IsNullOrWhiteSpace(paragraph.Text)) return false;
            Word.Range? range = null;
            try
            {
                range = ResolveCurrentMainParagraphRange(document, paragraph);
                var party = IsParty(snapshot);
                // A point is defined by its structural marker, not by a guessed
                // semantic role. Prioritize this deterministic correction so a stale
                // or overly broad role can never prevent 1-Click/selected-fix from
                // removing bold and italic formatting from the whole point.
                if (!paragraph.IsInTable && IsLegalPointParagraph(paragraph.Text))
                {
                    range.Font.Bold = 0;
                    range.Font.Italic = 0;
                    NormalizeBodySpacing(range, party);
                    return range.Font.Bold == 0 && range.Font.Italic == 0;
                }
                if (string.IsNullOrEmpty(role))
                {
                    if (paragraph.IsInTable) return false;
                    if (Regex.IsMatch(paragraph.Text.Trim(), @"^Điều\s+\d+\s*[.:]",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    {
                        // Preserve mixed bold article labels while normalizing the
                        // paragraph that also contains the operative body text.
                        NormalizeBodySpacing(range, party);
                        return true;
                    }
                    if (!IsBodyParagraph(paragraph)) return false;
                    range.Font.Name = rules.BodyFontName;
                    range.Font.Color = Word.WdColor.wdColorAutomatic;
                    range.Font.Size = party ? 14f : ValidOr(paragraph.FontSizePoints, 13, 14, 14);
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    ApplyBodyParagraphIndent(range);
                    range.ParagraphFormat.SpaceBeforeAuto = 0;
                    range.ParagraphFormat.SpaceAfterAuto = 0;
                    range.ParagraphFormat.SpaceBefore = 0f;
                    range.ParagraphFormat.SpaceAfter = 6f;
                    if (party)
                    {
                        range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                        range.ParagraphFormat.LineSpacing = 18f;
                    }
                    else NormalizeNd30ContentLineSpacing(range);
                    return true;
                }

                var style = StyleFor(role, party, paragraph, headerFontTier);
                if (style == null) return false;
                range.Font.Name = rules.BodyFontName;
                range.Font.Color = Word.WdColor.wdColorAutomatic;
                range.Font.Size = style.Size;
                range.Font.Bold = style.Bold ? -1 : 0;
                range.Font.Italic = style.Italic ? -1 : 0;
                // D6: a single underline of the whole motto is a valid separator
                // chosen by the author; keep it instead of replacing it by a shape.
                if (!(role == "nationalMotto" && paragraph.Underline == 1))
                    range.Font.Underline = style.Underline
                        ? Word.WdUnderline.wdUnderlineSingle
                        : Word.WdUnderline.wdUnderlineNone;
                if (role == "appendixTitle")
                    range.Case = Word.WdCharacterCase.wdUpperCase;
                // Quốc hiệu typed in lower case (accepted only inside the header
                // zone) is capitalized in place; Range.Case keeps runs and anchors.
                // A combined title + motto paragraph is never upper-cased.
                if (role == "nationalTitle" && !party &&
                    ComponentTextNormalizer.IsNationalTitle(paragraph.Text) &&
                    !ComponentTextNormalizer.IsUppercaseNationalTitle(paragraph.Text))
                    range.Case = Word.WdCharacterCase.wdUpperCase;
                range.ParagraphFormat.Alignment = style.Alignment;
                if (!party && role != "nationalTitle" && role != "nationalMotto" &&
                    role != "organName" && role != "superiorOrganName")
                    NormalizeNd30ContentLineSpacing(range);
                if (role == "recipientLabel" || role == "recipientList")
                {
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.FirstLineIndent = 0f;
                    range.ParagraphFormat.RightIndent = 0f;
                }
                if (role == "subject" || role == "subjectContinuation")
                {
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.RightIndent = 0f;
                    range.ParagraphFormat.FirstLineIndent = 0f;
                }
                if (role == "placeAndIssuedDate" && !party)
                {
                    range.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)HeaderLayoutPolicy.DateAlignment(
                        new DocumentRoleDetector().DetectBlocks(snapshot), paragraph.Index);
                    // Center against the full header cell, not an inherited body
                    // indent that visually shifts the date to the right.
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.RightIndent = 0f;
                    range.ParagraphFormat.FirstLineIndent = 0f;
                }
                if (role == "nationalTitle" || role == "nationalMotto" ||
                    (!party && (role == "organName" || role == "superiorOrganName")))
                {
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.RightIndent = 0f;
                    range.ParagraphFormat.FirstLineIndent = 0f;
                    range.ParagraphFormat.SpaceBeforeAuto = 0;
                    range.ParagraphFormat.SpaceAfterAuto = 0;
                    range.ParagraphFormat.SpaceBefore = 0f;
                    range.ParagraphFormat.SpaceAfter = 0f;
                    range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                if (role == "legalBasis")
                {
                    range.ParagraphFormat.SpaceBeforeAuto = 0;
                    range.ParagraphFormat.SpaceAfterAuto = 0;
                    range.ParagraphFormat.SpaceBefore = 0f;
                    range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    if (!party) NormalizeNd30ContentLineSpacing(range);
                    range.ParagraphFormat.LeftIndent = 0f;
                    range.ParagraphFormat.RightIndent = 0f;
                    range.ParagraphFormat.FirstLineIndent = 10f * PointsPerMillimeter;
                    range.ParagraphFormat.SpaceAfter = 6f;
                }
                return true;
            }
            finally { Release(range); }
        }

        private static int RemoveMisclassifiedOwnedLines(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            var removed = 0;
            foreach (var line in snapshot.LineShapes.Where(l => LineShapeOwnership.IsOwned(l.Name)))
            {
                var expectedRole = line.Name.IndexOf("_ORG_", StringComparison.Ordinal) >= 0 ? "organName" :
                    line.Name.IndexOf("_MOTTO_", StringComparison.Ordinal) >= 0 ? "nationalMotto" :
                    line.Name.IndexOf("_SUBJ_", StringComparison.Ordinal) >= 0 ? "subject" :
                    line.Name.IndexOf("_PARTY_", StringComparison.Ordinal) >= 0 ? "partyTitle" : string.Empty;
                if (expectedRole.Length == 0) continue;
                if (roles.Any(r => LineShapeOwnership.IsOwnedForParagraph(line.Name, r.Key) &&
                    (r.Value == expectedRole || (expectedRole == "subject" && r.Value == "subjectContinuation"))))
                    continue;
                // Only discard a derived line whose recorded semantic owner no
                // longer exists. Never remove an unmarked user-created shape.
                DeleteOwnedShapeByName(document, line.Name);
                removed++;
            }
            return removed;
        }

        private bool FixCombinedHeaderFinding(Word.Document document, DocumentContext context,
            LocalScanSnapshot snapshot, LocalParagraphSnapshot paragraph, LocalRulePack rules)
        {
            if (SplitCombinedNationalHeaders(document, snapshot, paragraph.Index) == 0) return false;
            new WordDocumentReadRuntime(_application, _accessManager).Prepare(
                context, DocumentAnalysisScope.Full, document, false);
            var current = context.LastLocalSnapshot!;
            var title = current.Paragraphs.First(p => p.Index == paragraph.Index);
            var motto = current.Paragraphs.FirstOrDefault(p => p.Index == paragraph.Index + 1);
            if (motto == null || !context.LastRolesByParagraphIndex.TryGetValue(motto.Index, out var role) ||
                role != "nationalMotto") return false;
            var tier = ResolveHeaderTier(current, context.LastLogicalBlocks, title.Index);
            ApplyParagraphFormat(document, title, "nationalTitle", current, rules, tier);
            ApplyParagraphFormat(document, motto, "nationalMotto", current, rules, tier);
            return NormalizeRequiredLine(document, current, motto, 1d, "ND30-PL1-M2-K1-TN-LINE");
        }

        private static int SplitCombinedNationalHeaders(Word.Document document, LocalScanSnapshot snapshot,
            int? onlyParagraph = null)
        {
            var count = 0;
            foreach (var paragraph in snapshot.Paragraphs.Where(p =>
                p.StoryType == "wdMainTextStory" && p.HasField != true &&
                p.HasContentControl != true && p.HasHyperlink != true && p.HasMathObject != true &&
                !snapshot.IntersectsProtectedSpan(p) && (!onlyParagraph.HasValue || p.Index == onlyParagraph.Value))
                .OrderByDescending(p => p.AbsoluteStart))
            {
                var offsets = CombinedNationalHeader.GetBreakOffsets(paragraph.Text);
                if (offsets.Length == 0) continue;
                Word.Range? range = null;
                try
                {
                    range = ResolveCurrentMainParagraphRange(document, paragraph);
                    if (!string.Equals((range.Text ?? string.Empty).TrimEnd('\r', '\a'),
                        paragraph.Text, StringComparison.Ordinal)) continue;
                    // Replace break characters only; preserve all runs, bookmarks,
                    // formatting and table-cell terminators.
                    foreach (var offset in offsets.OrderByDescending(value => value))
                    {
                        var start = range.Start;
                        range.SetRange(start + offset, start + offset + 1);
                        range.Text = "\r";
                        range.SetRange(start, start + offset);
                    }
                    count++;
                }
                finally { Release(range); }
            }
            return count;
        }

        private static void ApplyBodyParagraphIndent(Word.Range range)
        {
            Word.ParagraphFormat? format = null;
            try
            {
                format = range.ParagraphFormat;
                format.RightIndent = 0f;
                // A literal dash is text, not a Word list. Use the same geometry
                // as the scanner and preserve existing tab stops without adding any.
                format.LeftIndent = (float)(ParagraphIndentPolicy.BodyLeftMillimeters * PointsPerMillimeter);
                format.FirstLineIndent = (float)(ParagraphIndentPolicy.BodyFirstLineMillimeters * PointsPerMillimeter);
            }
            finally
            {
                Release(format);
            }
        }

        private static ParagraphStyle? StyleFor(string role, bool party, LocalParagraphSnapshot paragraph,
            Nd30HeaderFontSizeTier headerFontTier)
        {
            switch (role)
            {
                case "standaloneTitle": return S(14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "nationalTitle": return party ? null : S((float)headerFontTier.NationalTitle, true, false,
                    Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "nationalMotto": return party ? null : S((float)headerFontTier.NationalMotto, true, false,
                    Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "partyTitle": return party ? S(15, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter) : null;
                case "superiorOrganName": return S(party ? 14 : 13, false, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "organName": return S(party ? 14 : 13, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "codeNumber": return S(party ? 14 : 13, false, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "placeAndIssuedDate": return S(party ? 14f : (float)headerFontTier.PlaceAndIssuedDate,
                    false, true, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "typeName": return S(party ? 16 : 14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "subject":
                case "subjectContinuation": return S(party ? 15 : 14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "officialLetterSubject": return S(party ? 12 : 13, false, party, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "legalBasis": return S(party ? 14 : 14, false, true, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                case "signerAuthority": return S(14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "recipientLabel": return party
                    ? S(14, false, false, Word.WdParagraphAlignment.wdAlignParagraphLeft, true)
                    : S(12, true, true, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                case "recipientList": return S(party ? 12 : 11, false, false, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                case "recipientSalutation":
                case "recipientSalutationInline": return S(party ? 14 : 14, false, party, paragraph.Alignment.HasValue
                    ? (Word.WdParagraphAlignment)paragraph.Alignment.Value : Word.WdParagraphAlignment.wdAlignParagraphLeft);
                case "recipientSalutationList": return S(party ? 14 : 14, false, false, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                case "appendixLabel": return S(14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "appendixTitle": return S(14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "appendixReference": return S(14, false, true, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                case "appendixDigitalSignatureInfo": return S(10, false, false, Word.WdParagraphAlignment.wdAlignParagraphRight);
                case "partChapterHeading":
                case "sectionHeading":
                case "structuralTitle": return S(party ? 15 : 14, true, false, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                default: return null;
            }
        }

        private static int InsertMissingRequiredLines(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyList<AnnotationFinding> findings,
            IReadOnlyList<LogicalDocumentBlock> blocks,
            IReadOnlyDictionary<int, string> roles,
            IReadOnlyDictionary<string, string> originalLegacyLineRules)
        {
            var count = 0;
            var work = findings.Where(f => f.RuleCode.EndsWith("-LINE", StringComparison.Ordinal) &&
                    f.Anchor.ParagraphIndex.HasValue)
                .Select(f => Tuple.Create(f.RuleCode, f.Anchor.ParagraphIndex.GetValueOrDefault()))
                .ToList();

            // Every line owned by Chuẩn hóa must be measured again after the final
            // layout pass. A line can be valid in the input snapshot and become offset
            // when 1-Click changes a header table's indent or column widths. Legacy
            // CHUANHOA_* lines are migrated and current CHUANHOA2_* lines are moved in
            // place. User-created shapes are never added to this maintenance path.
            foreach (var line in snapshot.LineShapes.Where(item =>
                LineShapeOwnership.IsOwned(item.Name)))
            {
                var componentRole = line.Name.IndexOf("_ORG_", StringComparison.Ordinal) >= 0 ? "organName" :
                    line.Name.IndexOf("_SUBJ_", StringComparison.Ordinal) >= 0 ? "subject" :
                    line.Name.IndexOf("_PARTY_", StringComparison.Ordinal) >= 0 ? "partyTitle" :
                    line.Name.IndexOf("_MOTTO_", StringComparison.Ordinal) >= 0 ? "nationalMotto" : string.Empty;
                var ruleCode = componentRole == "organName" ? "ND30-PL1-M2-K2-ORG-LINE" :
                    componentRole == "subject" ? "ND30-PL1-M2-K5A-SUBJ-LINE" :
                    componentRole == "partyTitle" ? "HD05-M1-TITLE-LINE" :
                    componentRole == "nationalMotto" ? "ND30-PL1-M2-K1-TN-LINE" : string.Empty;
                if (ruleCode.Length == 0) continue;
                var paragraph = snapshot.Paragraphs.FirstOrDefault(item =>
                    roles.ContainsKey(item.Index) &&
                    string.Equals(roles[item.Index], componentRole, StringComparison.Ordinal) &&
                    LineShapeOwnership.IsOwnedForParagraph(line.Name, item.Index));
                if (paragraph != null && !work.Any(item => item.Item1 == ruleCode && item.Item2 == paragraph.Index))
                    work.Add(Tuple.Create(ruleCode, paragraph.Index));
            }

            foreach (var block in blocks)
            {
                // D4: a technical document (no national/party title or type name in
                // its header zone) never receives component lines.
                if (!block.IsAdministrative) continue;
                if (IsParty(snapshot))
                {
                    AddRequiredRoleLines(work, snapshot, block, "partyTitle", "HD05-M1-TITLE-LINE", false);
                    continue;
                }
                AddRequiredRoleLines(work, snapshot, block, "nationalMotto", "ND30-PL1-M2-K1-TN-LINE", false);
                AddRequiredRoleLines(work, snapshot, block, "organName", "ND30-PL1-M2-K2-ORG-LINE", false);
                AddRequiredRoleLines(work, snapshot, block, "subject", "ND30-PL1-M2-K5A-SUBJ-LINE", true);
            }
            var targets = new List<Tuple<string, int>>();
            foreach (var item in work)
            {
                if (IsParty(snapshot) != item.Item1.StartsWith("HD05-", StringComparison.Ordinal)) continue;
                var index = item.Item2;
                if (item.Item1 == "ND30-PL1-M2-K5A-SUBJ-LINE")
                {
                    var block = blocks.FirstOrDefault(b => b.ContainsParagraph(index));
                    if (block == null) continue;
                    while (block.ContainsParagraph(index + 1) && block.Roles.TryGetValue(index + 1, out var nextRole) &&
                        nextRole == "subjectContinuation") index++;
                    if (index != item.Item2)
                        foreach (var line in snapshot.LineShapes.Where(l =>
                            LineShapeOwnership.IsOwnedForParagraph(l.Name, item.Item2) &&
                            l.Name.IndexOf("_SUBJ_", StringComparison.Ordinal) >= 0))
                            DeleteOwnedShapeByName(document, line.Name);
                }
                if (!targets.Any(t => t.Item1 == item.Item1 && t.Item2 == index))
                    targets.Add(Tuple.Create(item.Item1, index));
            }
            var paragraphsByIndex = snapshot.Paragraphs
                .Where(p => string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                .GroupBy(p => p.Index)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var item in targets)
            {
                var paragraphIndex = item.Item2;
                LocalParagraphSnapshot paragraph;
                if (!paragraphsByIndex.TryGetValue(paragraphIndex, out paragraph)) continue;
                // D6: an existing solid bottom border (any component) or a single
                // underline of the whole motto already satisfies NĐ30.
                if (paragraph.HasBottomBorder ||
                    (item.Item1 == "ND30-PL1-M2-K1-TN-LINE" && paragraph.Underline == 1))
                    continue;
                var ratio = item.Item1 == "ND30-PL1-M2-K2-ORG-LINE" ||
                    item.Item1 == "ND30-PL1-M2-K5A-SUBJ-LINE" ? .4d : 1d;
                if (NormalizeRequiredLine(document, snapshot, paragraph, ratio, item.Item1,
                        originalLegacyLineRules)) count++;
            }
            return count;
        }

        private static void AddRequiredRoleLines(ICollection<Tuple<string, int>> work,
            LocalScanSnapshot snapshot, LogicalDocumentBlock block, string role,
            string ruleCode, bool useLastSubjectParagraph)
        {
            var matches = snapshot.Paragraphs.Where(paragraph =>
                    string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    block.Roles.TryGetValue(paragraph.Index, out var detectedRole) &&
                    string.Equals(detectedRole, role, StringComparison.Ordinal))
                .OrderBy(paragraph => paragraph.Index)
                .ToArray();
            if (matches.Length == 0) return;
            foreach (var match in matches)
            {
                var paragraph = match;
                if (useLastSubjectParagraph)
                {
                    while (block.Roles.TryGetValue(paragraph.Index + 1, out var nextRole) &&
                        string.Equals(nextRole, "subjectContinuation", StringComparison.Ordinal))
                    {
                        var next = snapshot.Paragraphs.FirstOrDefault(item =>
                            item.Index == paragraph.Index + 1 &&
                            string.Equals(item.StoryType, paragraph.StoryType, StringComparison.Ordinal));
                        if (next == null) break;
                        paragraph = next;
                    }
                }
                if (!work.Any(item => item.Item1 == ruleCode && item.Item2 == paragraph.Index))
                    work.Add(Tuple.Create(ruleCode, paragraph.Index));
            }
        }

        private static bool NormalizeRequiredLine(Word.Document document, LocalScanSnapshot snapshot,
            LocalParagraphSnapshot paragraph, double ratio, string ruleCode) =>
            NormalizeRequiredLine(document, snapshot, paragraph, ratio, ruleCode,
                FindLegacyHeaderSeparatorRules(snapshot,
                    new DocumentRoleDetector().Detect(snapshot)));

        private static bool NormalizeRequiredLine(Word.Document document, LocalScanSnapshot snapshot,
            LocalParagraphSnapshot paragraph, double ratio, string ruleCode,
            IReadOnlyDictionary<string, string> originalLegacyLineRules)
        {
            Word.Range? range = null;
            Word.Range? lineAnchor = null;
            Word.Font? currentFont = null;
            Word.Shape? shape = null;
            try
            {
                RemoveTextualLineSubstitutes(document, snapshot, paragraph);
                range = ResolveCurrentMainParagraphRange(document, paragraph);
                var top = WordTextMeasurement.ReadLastTextLineTop(range)
                    .GetValueOrDefault(paragraph.PageTopPoints.GetValueOrDefault(-1d));
                currentFont = range.Font;
                var fontName = currentFont.Name;
                var fontSize = currentFont.Size > 0 && currentFont.Size < 999999f
                    ? (double?)currentFont.Size
                    : paragraph.FontSizePoints;
                var bold = currentFont.Bold == -1 ? true : currentFont.Bold == 0 ? false : (bool?)null;
                var italic = currentFont.Italic == -1 ? true : currentFont.Italic == 0 ? false : (bool?)null;
                var textWidth = WordTextMeasurement.MeasureParagraphWidth(range, fontName, fontSize, bold, italic);
                var center = WordTextMeasurement.ReadHorizontalCenter(range);
                if (top < 0 || textWidth <= 0 || !center.HasValue) return false;
                var width = (float)Math.Max(28d, textWidth * ratio);
                var beginX = (float)(center.Value - width / 2d);
                var y = (float)(top + fontSize.GetValueOrDefault(13d) * 1.25d + 2d);
                // The author's underline or paragraph border is never removed here
                // (D6 accepts both); paragraphs carrying one are not sent to this
                // method by InsertMissingRequiredLines.
                var inTable = range.get_Information(Word.WdInformation.wdWithInTable);
                if (ruleCode == "ND30-PL1-M2-K5A-SUBJ-LINE")
                {
                    var detected = new DocumentRoleDetector().Detect(snapshot);
                    var firstIndex = paragraph.Index;
                    while (firstIndex > 1 && detected.TryGetValue(firstIndex, out var currentRole) &&
                        currentRole == "subjectContinuation") firstIndex--;
                    var firstParagraph = snapshot.Paragraphs.FirstOrDefault(p => p.Index == firstIndex);
                    if (firstParagraph != null && firstIndex < paragraph.Index)
                    {
                        Word.Range? firstRange = null;
                        try
                        {
                            firstRange = ResolveCurrentMainParagraphRange(document, firstParagraph);
                            var firstTop = WordTextMeasurement.ReadFirstTextLineTop(firstRange);
                            if (firstTop.HasValue)
                                foreach (var line in snapshot.LineShapes.Where(l => l.ShapeType == 9 &&
                                    l.AnchorStoryType == paragraph.StoryType && l.AnchorPageNumber == paragraph.PageNumber &&
                                    l.PageTopPoints.HasValue && l.PageLeftPoints.HasValue &&
                                    l.PageTopPoints.Value > firstTop.Value && l.PageTopPoints.Value < top &&
                                    Math.Abs(l.HeightPoints) <= 3 &&
                                    Math.Abs(l.PageLeftPoints.Value + l.WidthPoints / 2d - center.Value) <= 18d &&
                                    l.WidthPoints >= 10 && l.WidthPoints <= textWidth))
                                    DeleteVerifiedLine(document, line.Name);
                        }
                        finally { Release(firstRange); }
                    }
                }
                // Reuse the original separator when a previous release added a
                // second owned line beside it. Match page and tight rendered
                // geometry, not a nearby paragraph anchor alone.
                var page = WordTextMeasurement.SafeInformation(range,
                    Word.WdInformation.wdActiveEndPageNumber).GetValueOrDefault();
                var ownerBlock = new DocumentRoleDetector().DetectBlocks(snapshot)
                    .FirstOrDefault(b => b.ContainsParagraph(paragraph.Index));
                var availableWidth = WordTextMeasurement.ReadAvailableWidth(range)
                    .GetValueOrDefault(textWidth);
                var originals = snapshot.LineShapes.Where(l => !LineShapeOwnership.IsOwned(l.Name) &&
                    l.ShapeType == 9 && l.AnchorStoryType == paragraph.StoryType &&
                    l.AnchorSectionIndex == paragraph.SectionIndex && ownerBlock != null &&
                    l.AnchorParagraphIndex.HasValue && (ownerBlock.ContainsParagraph(l.AnchorParagraphIndex.Value) ||
                        IsAdjacentBlankHeaderAnchor(snapshot, ownerBlock, l)) &&
                    page > 0 && l.AnchorPageNumber == page && l.PageLeftPoints.HasValue && l.PageTopPoints.HasValue &&
                    (Math.Abs(l.PageTopPoints.Value - y) <= 18d &&
                     Math.Abs(l.PageLeftPoints.Value + l.WidthPoints / 2d - center.Value) <= 18d ||
                    IsLegacyHeaderSeparator(snapshot, paragraph, l, ruleCode, top, center.Value, textWidth,
                         availableWidth)) &&
                    Math.Abs(l.HeightPoints) <= 3d &&
                    l.WidthPoints >= 10d && l.WidthPoints <= availableWidth).ToArray();
                var preparedLegacyNames = originalLegacyLineRules.Where(item =>
                        string.Equals(item.Value, ruleCode, StringComparison.Ordinal))
                    .Select(item => item.Key).ToArray();
                var currentNames = new HashSet<string>(
                    snapshot.LineShapes.Select(item => item.Name), StringComparer.Ordinal);
                var preparedOriginals = preparedLegacyNames.Where(currentNames.Contains).ToArray();
                if (originals.Length >= 1)
                {
                    RemoveObsoleteComponentLines(document, snapshot, paragraph, ruleCode);
                    foreach (var duplicate in originals)
                        DeleteVerifiedLine(document, duplicate.Name);
                }
                foreach (var duplicateName in preparedOriginals)
                    DeleteVerifiedLine(document, duplicateName);
                // Page-relative coordinates do not relocate an existing shape's
                // anchor. Recreate only generated separators with a verified local
                // anchor so a stale name can never drag a later header onto a cover.
                RemoveObsoleteComponentLines(document, snapshot, paragraph, ruleCode);
                lineAnchor = inTable ? ResolveNonTableLineAnchor(document, range, snapshot, paragraph) : range.Duplicate;
                object anchor = lineAnchor;
                shape = document.Shapes.AddLine(beginX, y, beginX + width, y, ref anchor);
                shape.Name = OwnedLineName(ruleCode, paragraph.Index);
                shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                // AddLine initially interprets coordinates relative to its anchor.
                // Re-apply page coordinates after switching the relative-position
                // modes so table-cell headings do not move the line into a later row.
                shape.Left = beginX;
                shape.Top = y;
                NormalizeLineStyle(shape);
                AnchorLineToParagraph(shape, lineAnchor, y);
                return true;
            }
            finally { Release(shape); Release(currentFont); Release(lineAnchor); Release(range); }
        }

        /// <summary>
        /// D5: keep the horizontal page position (the caption does not move
        /// sideways when content changes) but make the vertical position relative
        /// to the anchor paragraph and lock the anchor, so the rule follows its
        /// component when paragraphs are added above it or the file is reopened on
        /// another printer. Word either keeps the visual position when the
        /// relative-position mode changes or keeps the old number; both cases are
        /// handled by comparing with the expected paragraph-relative offset.
        /// </summary>
        private static void AnchorLineToParagraph(Word.Shape shape, Word.Range lineAnchor, float pageTop)
        {
            try
            {
                var anchorTop = WordTextMeasurement.ReadFirstTextLineTop(lineAnchor);
                if (!anchorTop.HasValue)
                {
                    shape.LockAnchor = 0;
                    return;
                }
                var expected = (float)(pageTop - anchorTop.Value);
                shape.RelativeVerticalPosition =
                    Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
                if (Math.Abs(shape.Top - expected) > 6f) shape.Top = expected;
                shape.LockAnchor = -1;
            }
            catch (COMException)
            {
                // Fall back to the previous, page-relative placement.
                try
                {
                    shape.RelativeVerticalPosition =
                        Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    shape.Top = pageTop;
                    shape.LockAnchor = 0;
                }
                catch (COMException) { }
            }
        }

        private static void RemoveTextualLineSubstitutes(Word.Document document,
            LocalScanSnapshot snapshot, LocalParagraphSnapshot paragraph)
        {
            Word.Range? componentRange = null;
            try
            {
                componentRange = ResolveCurrentMainParagraphRange(document, paragraph);
                var current = componentRange.Text ?? string.Empty;
                var cleaned = Regex.Replace(current,
                    @"[\v\n][ \t]*[-_.\u2010\u2011\u2012\u2013\u2014\u2015](?:[ \t]*[-_.\u2010\u2011\u2012\u2013\u2014\u2015]){2,}[ \t]*(?=[\r\a]*$)",
                    string.Empty, RegexOptions.CultureInvariant);
                if (!string.Equals(current, cleaned, StringComparison.Ordinal))
                {
                    var componentStart = componentRange.Start;
                    componentRange.Text = cleaned;
                    RecordTextEdit(componentStart, current.Length, cleaned.Length);
                }
            }
            finally { Release(componentRange); }

            foreach (var candidate in snapshot.Paragraphs.Where(item =>
                item.Index > paragraph.Index && item.Index <= paragraph.Index + 2 &&
                string.Equals(item.StoryType, paragraph.StoryType, StringComparison.Ordinal) &&
                item.SectionIndex == paragraph.SectionIndex &&
                SameTextContainer(item, paragraph) &&
                IsTextualLineSeparator(item.Text)))
            {
                Word.Range? candidateRange = null;
                try
                {
                    candidateRange = ResolveCurrentMainParagraphRange(document, candidate);
                    var current = candidateRange.Text ?? string.Empty;
                    var contentLength = current.TrimEnd('\r', '\a').Length;
                    if (contentLength <= 0) continue;
                    candidateRange.End = candidateRange.Start + contentLength;
                    var separatorStart = candidateRange.Start;
                    candidateRange.Text = string.Empty;
                    RecordTextEdit(separatorStart, contentLength, 0);
                }
                finally { Release(candidateRange); }
            }
        }

        private static bool SameTextContainer(LocalParagraphSnapshot left,
            LocalParagraphSnapshot right)
        {
            if (!left.IsInTable && !right.IsInTable) return true;
            return left.IsInTable && right.IsInTable &&
                left.TableIndex == right.TableIndex &&
                left.RowIndex == right.RowIndex &&
                left.CellIndex == right.CellIndex;
        }

        private static bool IsTextualLineSeparator(string text)
        {
            var value = (text ?? string.Empty).Trim(' ', '\t', '\r', '\n', '\v', '\a');
            if (value.Length < 3 || value.Length > 120) return false;
            var marks = 0;
            foreach (var character in value)
            {
                if (character == ' ' || character == '\t') continue;
                if (character == '-' || character == '_' || character == '.' ||
                    character == '\u2010' || character == '\u2011' || character == '\u2012' ||
                    character == '\u2013' || character == '\u2014' || character == '\u2015')
                {
                    marks++;
                    continue;
                }
                return false;
            }
            return marks >= 3;
        }

        private static bool IsLegacyHeaderSeparator(LocalScanSnapshot snapshot,
            LocalParagraphSnapshot paragraph, LocalLineShapeSnapshot line, string ruleCode,
            double textTop, double center, double textWidth, double containerWidth)
        {
            if (!paragraph.IsInTable || !paragraph.TableIndex.HasValue ||
                (ruleCode != "ND30-PL1-M2-K1-TN-LINE" && ruleCode != "ND30-PL1-M2-K2-ORG-LINE")) return false;
            var anchor = snapshot.Paragraphs.FirstOrDefault(p => p.Index == line.AnchorParagraphIndex &&
                p.StoryType == paragraph.StoryType);
            if (anchor == null) return false;
            var sameColumn = anchor.TableIndex == paragraph.TableIndex && anchor.CellIndex == paragraph.CellIndex &&
                anchor.RowIndex.HasValue && paragraph.RowIndex.HasValue &&
                anchor.RowIndex.Value >= paragraph.RowIndex.Value && anchor.RowIndex.Value <= paragraph.RowIndex.Value + 1;
            // Legacy Word files can serialize an organ separator physically in
            // the left header cell while moving its anchor to the national-header
            // cell on the right. Accept only that immediate sibling on the same
            // row; the geometry checks below still require the line to be fully
            // contained inside the organ column.
            var siblingNationalHeaderAnchor = ruleCode == "ND30-PL1-M2-K2-ORG-LINE" &&
                anchor.TableIndex == paragraph.TableIndex && anchor.RowIndex == paragraph.RowIndex &&
                anchor.CellIndex.HasValue && paragraph.CellIndex.HasValue &&
                anchor.CellIndex.Value == paragraph.CellIndex.Value + 1;
            // Word can move a legacy floating-line anchor onto the blank paragraph
            // immediately preceding the header table. Require both this structural
            // adjacency and containment inside the owner's measured column.
            var adjacentBlank = !anchor.IsInTable && string.IsNullOrWhiteSpace(anchor.Text) &&
                anchor.Index < paragraph.Index && paragraph.Index - anchor.Index <= 2 && paragraph.RowIndex == 1;
            if (!sameColumn && !siblingNationalHeaderAnchor && !adjacentBlank) return false;
            var lineCenter = line.PageLeftPoints.GetValueOrDefault(double.MinValue) + line.WidthPoints / 2d;
            if (siblingNationalHeaderAnchor)
            {
                if (!paragraph.PageLeftPoints.HasValue || !paragraph.PageTopPoints.HasValue ||
                    !paragraph.TextWidthPoints.HasValue || paragraph.TextWidthPoints.Value <= 0) return false;
                var originalTextWidth = paragraph.TextWidthPoints.Value;
                var originalCenter = paragraph.PageLeftPoints.Value + originalTextWidth / 2d;
                var originalDistance = line.PageTopPoints.GetValueOrDefault(double.MinValue) -
                    paragraph.PageTopPoints.Value;
                return originalDistance >= 0 && originalDistance <= 60 &&
                    Math.Abs(lineCenter - originalCenter) + line.WidthPoints / 2d <= originalTextWidth / 2d + 1d &&
                    line.WidthPoints >= originalTextWidth * .25 && line.WidthPoints <= originalTextWidth * 1.2 &&
                    line.BeginArrowheadStyle == 1 && line.EndArrowheadStyle == 1;
            }
            var distance = line.PageTopPoints.GetValueOrDefault(double.MinValue) - textTop;
            return distance >= 0 && distance <= 60 &&
                Math.Abs(lineCenter - center) + line.WidthPoints / 2d <= containerWidth / 2d + 1d &&
                line.WidthPoints >= textWidth * .25 && line.WidthPoints <= textWidth * 1.2 &&
                line.BeginArrowheadStyle == 1 && line.EndArrowheadStyle == 1;
        }

        private static bool IsLegacyHeaderSeparatorFromPreparedLayout(
            IReadOnlyDictionary<int, LocalParagraphSnapshot> preparedParagraphs,
            LocalParagraphSnapshot paragraph, LocalLineShapeSnapshot line, string ruleCode)
        {
            if (!paragraph.IsInTable || !paragraph.TableIndex.HasValue ||
                (ruleCode != "ND30-PL1-M2-K1-TN-LINE" &&
                 ruleCode != "ND30-PL1-M2-K2-ORG-LINE")) return false;
            if (!line.AnchorParagraphIndex.HasValue ||
                !preparedParagraphs.TryGetValue(line.AnchorParagraphIndex.Value, out var anchor) ||
                !string.Equals(anchor.StoryType, paragraph.StoryType, StringComparison.Ordinal)) return false;
            var sameColumn = anchor.TableIndex == paragraph.TableIndex &&
                anchor.CellIndex == paragraph.CellIndex && anchor.RowIndex.HasValue &&
                paragraph.RowIndex.HasValue && anchor.RowIndex.Value >= paragraph.RowIndex.Value &&
                anchor.RowIndex.Value <= paragraph.RowIndex.Value + 1;
            var siblingNationalHeaderAnchor = ruleCode == "ND30-PL1-M2-K2-ORG-LINE" &&
                anchor.TableIndex == paragraph.TableIndex && anchor.RowIndex == paragraph.RowIndex &&
                anchor.CellIndex.HasValue && paragraph.CellIndex.HasValue &&
                anchor.CellIndex.Value == paragraph.CellIndex.Value + 1;
            if (!sameColumn && !siblingNationalHeaderAnchor) return false;
            if (!line.PageLeftPoints.HasValue || !line.PageTopPoints.HasValue ||
                !paragraph.PageLeftPoints.HasValue || !paragraph.PageTopPoints.HasValue ||
                !paragraph.TextWidthPoints.HasValue || paragraph.TextWidthPoints.Value <= 0 ||
                line.AnchorSectionIndex != paragraph.SectionIndex ||
                line.AnchorPageNumber != paragraph.PageNumber || line.BeginArrowheadStyle != 1 ||
                line.EndArrowheadStyle != 1) return false;
            var textWidth = paragraph.TextWidthPoints.Value;
            var center = paragraph.PageLeftPoints.Value + textWidth / 2d;
            var lineCenter = line.PageLeftPoints.Value + line.WidthPoints / 2d;
            var distance = line.PageTopPoints.Value - paragraph.PageTopPoints.Value;
            return distance >= 0d && distance <= 60d &&
                Math.Abs(lineCenter - center) + line.WidthPoints / 2d <= textWidth / 2d + 1d &&
                line.WidthPoints >= textWidth * .25d && line.WidthPoints <= textWidth * 1.2d;
        }

        private static IReadOnlyDictionary<string, string> FindLegacyHeaderSeparatorRules(
            LocalScanSnapshot snapshot, IReadOnlyDictionary<int, string> roles)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            var paragraphs = snapshot.Paragraphs.ToDictionary(item => item.Index);
            foreach (var paragraph in snapshot.Paragraphs)
            {
                if (!roles.TryGetValue(paragraph.Index, out var role)) continue;
                var ruleCode = string.Equals(role, "organName", StringComparison.Ordinal)
                    ? "ND30-PL1-M2-K2-ORG-LINE"
                    : string.Equals(role, "nationalMotto", StringComparison.Ordinal)
                        ? "ND30-PL1-M2-K1-TN-LINE" : string.Empty;
                if (ruleCode.Length == 0) continue;
                foreach (var line in snapshot.LineShapes.Where(item =>
                    !LineShapeOwnership.IsOwned(item.Name) && item.ShapeType == 9 &&
                    item.PageLeftPoints.HasValue && item.PageTopPoints.HasValue &&
                    Math.Abs(item.HeightPoints) <= 3d && item.WidthPoints >= 10d &&
                    IsLegacyHeaderSeparatorFromPreparedLayout(paragraphs, paragraph,
                        item, ruleCode)))
                {
                    if (!result.ContainsKey(line.Name)) result.Add(line.Name, ruleCode);
                }
            }
            return result;
        }

        private static bool IsAdjacentBlankHeaderAnchor(LocalScanSnapshot snapshot,
            LogicalDocumentBlock block, LocalLineShapeSnapshot line)
        {
            var anchor = snapshot.Paragraphs.FirstOrDefault(p => p.Index == line.AnchorParagraphIndex &&
                p.StoryType == line.AnchorStoryType);
            return anchor != null && !anchor.IsInTable && string.IsNullOrWhiteSpace(anchor.Text) &&
                anchor.Index < block.StartParagraphIndex && block.StartParagraphIndex - anchor.Index <= 2;
        }

        private static bool TryNormalizeExistingComponentLine(Word.Document document,
            LocalScanSnapshot snapshot, LocalParagraphSnapshot paragraph, string ruleCode,
            float targetLeft, float targetTop, float targetWidth)
        {
            var matches = snapshot.LineShapes
                .Where(line => IsAssociatedLine(line, paragraph))
                .OrderBy(line => line.PageTopPoints.HasValue && paragraph.PageTopPoints.HasValue
                    ? Math.Abs(line.PageTopPoints.Value - paragraph.PageTopPoints.Value)
                    : Math.Abs(line.AnchorAbsoluteStart - paragraph.AbsoluteStart))
                .ToArray();
            if (matches.Length == 0) return false;

            var ownedName = OwnedLineName(ruleCode, paragraph.Index);
            var selected = matches.FirstOrDefault(line =>
                string.Equals(line.Name, ownedName, StringComparison.Ordinal));
            if (selected != null)
            {
                return NormalizeOwnedComponentLine(document, selected.Name, targetLeft, targetTop,
                    targetWidth, ownedName, deleteDuplicateCurrentOwned: true);
            }

            // Legacy names are ours and may be replaced. Current CHUANHOA2_* names
            // are never classified as legacy. User-created Shape objects are neither
            // selected, mutated, renamed nor deleted by this path.
            var legacy = matches.FirstOrDefault(line =>
                LineShapeOwnership.IsLegacyOwned(line.Name) &&
                LineShapeOwnership.IsOwnedForParagraph(line.Name, paragraph.Index));
            if (legacy != null)
            {
                DeleteOwnedShapeByName(document, legacy.Name);
            }
            return false;
        }

        private static bool NormalizeOwnedComponentLine(Word.Document document, string selectedName,
            float targetLeft, float targetTop, float targetWidth, string ownedName,
            bool deleteDuplicateCurrentOwned)
        {
            var normalized = false;
            for (var index = document.Shapes.Count; index >= 1; index--)
            {
                Word.Shape? candidate = null;
                try
                {
                    candidate = document.Shapes[index];
                    if (!WordLineShapeKind.IsLineLike(candidate)) continue;
                    var name = candidate.Name ?? string.Empty;
                    if (normalized && deleteDuplicateCurrentOwned && string.Equals(name, ownedName, StringComparison.Ordinal))
                    {
                        candidate.Delete();
                        continue;
                    }
                    if (!string.Equals(name, selectedName, StringComparison.Ordinal))
                    {
                        if (deleteDuplicateCurrentOwned &&
                            string.Equals(name, ownedName, StringComparison.Ordinal) && normalized)
                            candidate.Delete();
                        continue;
                    }

                    NormalizeLineStyle(candidate);
                    candidate.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    candidate.RelativeVerticalPosition =
                        Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    candidate.Left = targetLeft;
                    candidate.Top = targetTop;
                    candidate.Width = targetWidth;
                    candidate.Height = 0f;
                    candidate.Name = ownedName;
                    candidate.LockAnchor = 0;
                    normalized = true;
                }
                catch (COMException)
                {
                }
                finally { Release(candidate); }
            }
            return normalized;
        }

        private static void DeleteOwnedShapeByName(Word.Document document, string name)
        {
            if (!LineShapeOwnership.IsOwned(name)) return;
            DeleteVerifiedLine(document, name);
        }

        private static void DeleteVerifiedLine(Word.Document document, string name)
        {
            for (var index = document.Shapes.Count; index >= 1; index--)
            {
                Word.Shape? candidate = null;
                try
                {
                    candidate = document.Shapes[index];
                    if (string.Equals(candidate.Name, name, StringComparison.Ordinal) &&
                        WordLineShapeKind.IsLineLike(candidate)) candidate.Delete();
                }
                catch (COMException) { }
                finally { Release(candidate); }
            }
        }

        private static Word.Range ResolveNonTableLineAnchor(Word.Document document, Word.Range source,
            LocalScanSnapshot snapshot, LocalParagraphSnapshot owner)
        {
            var sourcePage = WordTextMeasurement.SafeInformation(source,
                Word.WdInformation.wdActiveEndPageNumber).GetValueOrDefault();
            if (sourcePage <= 0) return source.Duplicate;
            var block = new DocumentRoleDetector().DetectBlocks(snapshot).FirstOrDefault(b => b.ContainsParagraph(owner.Index));
            if (block == null) return source.Duplicate;
            Word.Paragraphs? paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                var count = paragraphs.Count;
                for (var index = block.StartParagraphIndex; index <= Math.Min(count, block.EndParagraphIndex); index++)
                {
                    Word.Paragraph? paragraph = null;
                    Word.Range? candidate = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        candidate = paragraph.Range.Duplicate;
                        if (candidate.get_Information(Word.WdInformation.wdWithInTable)) continue;
                        var page = WordTextMeasurement.SafeInformation(candidate,
                            Word.WdInformation.wdActiveEndPageNumber).GetValueOrDefault();
                        var candidateSnapshot = snapshot.Paragraphs.FirstOrDefault(p => p.Index == index && p.StoryType == owner.StoryType);
                        if (candidateSnapshot == null || !ComponentLineAnchorPolicy.CanAnchor((int)sourcePage, (int)page,
                            owner.SectionIndex, candidateSnapshot.SectionIndex, index,
                            block.StartParagraphIndex, block.EndParagraphIndex)) continue;
                        return candidate.Duplicate;
                    }
                    catch (COMException)
                    {
                    }
                    finally { Release(candidate); Release(paragraph); }
                }
            }
            finally { Release(paragraphs); }
            return source.Duplicate;
        }

        private static void NormalizeLineStyle(Word.Shape shape)
        {
            // When the anchor lives in a layout table, Word otherwise adds the
            // cell's origin to page-relative Left/Top and visibly displaces the line.
            shape.LayoutInCell = 0;
            shape.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            // Legacy shapes can retain a negative wrap distance. OOXML serializes that
            // value as 4294967295 EMU and Word/PDF may stop painting the otherwise valid
            // line. Reset every distance whenever a component line is normalized.
            shape.WrapFormat.DistanceTop = 0f;
            shape.WrapFormat.DistanceBottom = 0f;
            shape.WrapFormat.DistanceLeft = 0f;
            shape.WrapFormat.DistanceRight = 0f;
            shape.Line.Visible = MsoTriState.msoTrue;
            shape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
            shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
            shape.Line.Weight = .75f;
            shape.Line.ForeColor.RGB = 0;
            shape.ZOrder(MsoZOrderCmd.msoBringToFront);
        }

        private static void RemoveObsoleteComponentLines(Word.Document document, LocalScanSnapshot snapshot,
            LocalParagraphSnapshot paragraph, string ruleCode)
        {
            var ownedName = OwnedLineName(ruleCode, paragraph.Index);
            for (var index = document.Shapes.Count; index >= 1; index--)
            {
                Word.Shape? candidate = null;
                try
                {
                    candidate = document.Shapes[index];
                    if (!WordLineShapeKind.IsLineLike(candidate)) continue;
                    var name = candidate.Name ?? string.Empty;
                    if (string.Equals(name, ownedName, StringComparison.Ordinal) ||
                        LineShapeOwnership.IsOwnedForParagraph(name, paragraph.Index))
                        candidate.Delete();
                }
                catch (COMException)
                {
                }
                finally { Release(candidate); }
            }
        }

        private static bool IsAssociatedLine(LocalLineShapeSnapshot line, LocalParagraphSnapshot paragraph)
        {
            if (line.ShapeType != 9 ||
                !string.Equals(line.AnchorStoryType, paragraph.StoryType, StringComparison.Ordinal) ||
                line.AnchorSectionIndex != paragraph.SectionIndex)
                return false;
            if (paragraph.PageNumber > 0 && line.AnchorPageNumber > 0 &&
                line.AnchorPageNumber != paragraph.PageNumber)
                return false;
            if (line.AnchorParagraphIndex.HasValue && line.AnchorParagraphIndex.Value == paragraph.Index)
                return true;
            if (line.PageTopPoints.HasValue && paragraph.PageTopPoints.HasValue &&
                line.PageLeftPoints.HasValue && paragraph.PageLeftPoints.HasValue &&
                paragraph.TextWidthPoints.HasValue && paragraph.TextWidthPoints.Value > 0d)
            {
                var verticalDistance = line.PageTopPoints.Value + Math.Abs(line.HeightPoints) / 2d -
                    paragraph.PageTopPoints.Value;
                if (verticalDistance < -6d || verticalDistance > 110d) return false;
                var lineCenter = line.PageLeftPoints.Value + Math.Abs(line.WidthPoints) / 2d;
                var textCenter = paragraph.PageLeftPoints.Value + paragraph.TextWidthPoints.Value / 2d;
                return Math.Abs(lineCenter - textCenter) <=
                    Math.Max(24d, paragraph.TextWidthPoints.Value * .35d);
            }
            return line.AnchorParagraphIndex.HasValue &&
                line.AnchorParagraphIndex.Value >= paragraph.Index - 1 &&
                line.AnchorParagraphIndex.Value <= paragraph.Index + 2;
        }

        private static void EnsurePageNumbers(Word.Document document, LocalScanSnapshot snapshot,
            LocalRulePack rules, IReadOnlyList<LogicalDocumentBlock> blocks)
        {
            var party = IsParty(snapshot);
            var administrativeFirstSections = AdministrativeFirstSections(snapshot, blocks);
            var technicalOnly = administrativeFirstSections.Count == 0;
            Word.Sections? sections = null;
            try
            {
                sections = document.Sections;
                var count = sections.Count;
                for (var index = 1; index <= count; index++)
                {
                    Word.Section? section = null;
                    try
                    {
                        section = sections[index];
                        EnsurePageNumber(document, section, party, rules,
                            administrativeFirstSections.Contains(index), technicalOnly);
                    }
                    finally { Release(section); }
                }
            }
            finally { Release(sections); }
        }

        /// <summary>Sections that open an administrative document (NĐ30 hides page 1's number there).</summary>
        private static HashSet<int> AdministrativeFirstSections(LocalScanSnapshot snapshot,
            IReadOnlyList<LogicalDocumentBlock> blocks)
        {
            var result = new HashSet<int>();
            var bySection = snapshot.Paragraphs
                .Where(p => string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                .GroupBy(p => p.Index)
                .ToDictionary(group => group.Key, group => group.First().SectionIndex);
            foreach (var block in blocks)
            {
                int sectionIndex;
                if (block.IsAdministrative && bySection.TryGetValue(block.StartParagraphIndex, out sectionIndex))
                    result.Add(sectionIndex);
            }
            return result;
        }

        private static void EnsurePageNumber(Word.Document document, Word.Section section,
            bool party, LocalRulePack rules) =>
            EnsurePageNumber(document, section, party, rules, section.Index == 1, false);

        private static void EnsurePageNumber(Word.Document document, Word.Section section,
            bool party, LocalRulePack rules, bool opensAdministrativeDocument, bool technicalOnly)
        {
            Word.HeaderFooter? header = null;
            Word.Range? range = null;
            Word.Range? headerRange = null;
            Word.Field? field = null;
            try
            {
                // Only the first page of an administrative document hides its number.
                // Later sections (for example a landscape progress chart) keep their
                // header links and their page numbers.
                if (opensAdministrativeDocument)
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                    RemoveFirstPageNumber(section);
                }
                // A technical document keeps page numbers the author put in the footer.
                if (!party && !technicalOnly) RemoveFooterPageNumbers(section);
                if (ContainsPageField(section)) return;
                header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (!header.Exists) header.Exists = true;
                headerRange = header.Range.Duplicate;
                var existingText = (headerRange.Text ?? string.Empty).Trim('\r', '\a', ' ', '\t');
                range = header.Range.Duplicate;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (existingText.Length > 0)
                {
                    // Keep the author's header paragraph (company/package name) and its
                    // alignment: the page number goes into its own new paragraph.
                    range.InsertParagraphBefore();
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                field = document.Fields.Add(range, Word.WdFieldType.wdFieldPage,
                    PreserveFormatting: true);
                field.Result.Font.Name = rules.BodyFontName;
                field.Result.Font.Size = 14f;
                field.Result.Font.Bold = 0;
                field.Result.Font.Italic = 0;
            }
            finally { Release(field); Release(headerRange); Release(range); Release(header); }
        }

        private static int NormalizeTables(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            // Tables holding a recognized component are layout tables, not data.
            var roleTables = new HashSet<int>(snapshot.Paragraphs
                .Where(paragraph => paragraph.TableIndex.HasValue && roles.ContainsKey(paragraph.Index) &&
                    string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                .Select(paragraph => paragraph.TableIndex!.Value));
            var count = 0;
            Word.Tables? tables = null;
            try
            {
                tables = document.Tables;
                var tableCount = tables.Count;
                for (var index = 1; index <= tableCount; index++)
                {
                    if (roleTables.Contains(index)) continue;
                    Word.Table? table = null;
                    Word.Row? first = null;
                    try
                    {
                        table = tables[index];
                        if (!IsLikelyDataTable(table)) continue;
                        // Repeat and keep together only the header row. Forbidding
                        // every row to break across pages left large blank areas
                        // below long quality-control rows in technical documents.
                        first = table.Rows[1];
                        first.HeadingFormat = -1;
                        first.AllowBreakAcrossPages = 0;
                        count++;
                    }
                    catch (COMException) { }
                    finally { Release(first); Release(table); }
                }
            }
            finally { Release(tables); }
            return count;
        }

        private static void NormalizeHeaderLayoutTables(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles, IReadOnlyList<LogicalDocumentBlock> blocks)
        {
            // A header layout table is identified only by detected roles, never by
            // the words it contains: a data table listing "Quốc hội nước Cộng hòa xã
            // hội chủ nghĩa Việt Nam..." must keep its borders and column widths.
            var headerTableIndexes = new HashSet<int>(snapshot.Paragraphs
                .Where(paragraph => paragraph.IsInTable && paragraph.TableIndex.HasValue &&
                    string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    roles.TryGetValue(paragraph.Index, out var role) &&
                    (role == "nationalTitle" || role == "partyTitle" || role == "organName") &&
                    blocks.Any(block => block.IsInHeaderZone(paragraph.Index)))
                .Select(paragraph => paragraph.TableIndex!.Value));
            if (headerTableIndexes.Count == 0) return;

            var normalizedAny = false;
            Word.Tables? tables = null;
            try
            {
                tables = document.Tables;
                var count = tables.Count;
                foreach (var tableIndex in headerTableIndexes.OrderBy(value => value))
                {
                    if (tableIndex < 1 || tableIndex > count) continue;
                    Word.Table? table = null;
                    try
                    {
                        table = tables[tableIndex];
                        if (IsHeaderLayoutTable(table, snapshot, tableIndex))
                        {
                            NormalizeHeaderTable(table);
                            normalizedAny = true;
                        }
                    }
                    catch (COMException) { }
                    finally { Release(table); }
                }
            }
            finally { Release(tables); }

            if (normalizedAny)
            {
                try { document.Repaginate(); } catch (COMException) { }
            }
        }

        private static bool IsHeaderLayoutTable(Word.Table table, LocalScanSnapshot snapshot,
            int tableIndex)
        {
            Word.Rows? rows = null;
            Word.Columns? columns = null;
            try
            {
                // Bounded geometry: an NĐ30 header layout table is at most 4 rows by
                // 3 columns. Larger tables are data tables even if they mention the
                // national title.
                rows = table.Rows;
                columns = table.Columns;
                if (rows.Count > 4 || columns.Count > 3) return false;
                // The snapshot must agree that this table holds only a handful of
                // short header paragraphs.
                var paragraphs = snapshot.Paragraphs.Where(paragraph => paragraph.IsInTable &&
                    paragraph.TableIndex == tableIndex &&
                    string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    !string.IsNullOrWhiteSpace(paragraph.Text)).ToArray();
                return paragraphs.Length > 0 && paragraphs.Length <= 16;
            }
            catch (COMException)
            {
                return false;
            }
            finally { Release(columns); Release(rows); }
        }

        private static void NormalizeHeaderTable(Word.Table table)
        {
            Word.Range? tableRange = null;
            Word.Sections? sections = null;
            Word.Section? section = null;
            try
            {
                tableRange = table.Range;
                sections = tableRange.Sections;
                section = sections.Count > 0 ? sections[1] : null;
                var pageWidth = section != null ? section.PageSetup.PageWidth : 595.3f;
                var leftMargin = section != null ? section.PageSetup.LeftMargin : 85.05f;
                var rightMargin = section != null ? section.PageSetup.RightMargin : 42.5f;

                // Allow table to overflow into the left margin so that left and right margins
                // of the header table are balanced and aesthetically pleasing ("đều với 2 lề").
                var targetLeftEdge = Math.Min(leftMargin, Math.Max(36f, rightMargin));
                var leftIndent = (float)Math.Round(targetLeftEdge - leftMargin, 1);
                var totalTableWidth = (float)Math.Max(200f, pageWidth - targetLeftEdge - rightMargin);

                table.Rows.LeftIndent = leftIndent;
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                try { table.Borders.Enable = 0; } catch (COMException) { }
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = totalTableWidth;

                // Ensure Column 2 is wide enough so "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
                // (which measures ~285pt at 13pt bold) never wraps into 2 lines.
                var col2Width = Math.Max(310f, (float)Math.Round(totalTableWidth * 0.60f, 1));
                var col1Width = (float)Math.Round(totalTableWidth - col2Width, 1);
                if (col1Width < 160f && totalTableWidth > 320f)
                {
                    col1Width = 160f;
                    col2Width = totalTableWidth - 160f;
                }

                for (var r = 1; r <= table.Rows.Count; r++)
                {
                    Word.Row? row = null;
                    try
                    {
                        row = table.Rows[r];
                        row.LeftIndent = leftIndent;
                        row.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                        if (row.Cells.Count == 2)
                        {
                            try
                            {
                                row.Cells[1].LeftPadding = 2f;
                                row.Cells[1].RightPadding = 2f;
                                row.Cells[1].Width = col1Width;
                            }
                            catch (COMException) { }

                            try
                            {
                                row.Cells[2].LeftPadding = 2f;
                                row.Cells[2].RightPadding = 2f;
                                row.Cells[2].Width = col2Width;
                            }
                            catch (COMException) { }
                        }
                    }
                    catch (COMException) { }
                    finally { Release(row); }
                }

                try
                {
                    if (table.Uniform && table.Columns.Count == 2)
                    {
                        table.Columns[1].SetWidth(col1Width, Word.WdRulerStyle.wdAdjustNone);
                        table.Columns[2].SetWidth(col2Width, Word.WdRulerStyle.wdAdjustNone);
                    }
                }
                catch (COMException) { }
            }
            catch (COMException) { }
            finally
            {
                Release(section);
                Release(sections);
                Release(tableRange);
            }
        }

        private static int ApplyDeterministicSpellingFixes(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyList<AnnotationFinding> findings, LocalRulePack rules,
            IReadOnlyList<LogicalDocumentBlock> blocks, bool includeRepeatedSpaces = false)
        {
            var paragraphs = snapshot.Paragraphs.ToDictionary(item => item.Index);
            var edits = new List<SpellingEdit>();
            var lexicon = new VietnameseLexiconSpellChecker(rules.Lexicon);
            var componentFindingCodes = new HashSet<string>(StringComparer.Ordinal)
            {
                "ND30-PL1-M2-K1-TN-SEP",
                "ND30-PL1-M2-K3-PREFIX",
                "ND30-PL1-M2-K3-SEP",
                "ND30-PL1-M2-K3-ABBR",
                "ND30-PL1-M2-K3-SPACE",
                "ND30-PL1-M2-K3-CASE",
                "ND30-PL1-M2-K3-PAD"
            };

            foreach (var group in findings.Where(item =>
                    item.Anchor.ParagraphIndex.HasValue &&
                    string.Equals(item.Anchor.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    componentFindingCodes.Contains(item.RuleCode))
                .GroupBy(item => item.Anchor.ParagraphIndex.GetValueOrDefault()))
            {
                LocalParagraphSnapshot paragraph;
                if (!paragraphs.TryGetValue(group.Key, out paragraph)) continue;
                var printable = (paragraph.Text ?? string.Empty).TrimEnd('\r', '\a');
                if (group.Any(item => item.RuleCode == "ND30-PL1-M2-K1-TN-SEP"))
                {
                    // Never overwrite the paragraph with the canonical motto. Only a
                    // paragraph that is exactly the motto is edited, and only the
                    // differing separators/letters are replaced, so character
                    // formatting and Shape anchors inside the paragraph survive.
                    var mottoEdits = NationalMottoEditPlanner.Plan(printable);
                    if (mottoEdits == null) continue;
                    foreach (var mottoEdit in mottoEdits)
                        edits.Add(new SpellingEdit(
                            checked(paragraph.AbsoluteStart + mottoEdit.Offset), mottoEdit.Length,
                            mottoEdit.Replacement, mottoEdit.ExpectedText, 100));
                    continue;
                }

                var abbreviation = ResolveExpectedTypeAbbreviation(snapshot,
                    paragraph.Index, rules, blocks);
                var replacement = LocalAdministrativeTextNormalizer.NormalizeCodeNumber(
                    printable, IsParty(snapshot), abbreviation);
                // Guard every whole-paragraph replacement of a fixed component: a
                // rewrite that changes the length a lot or keeps little of the
                // original text is a misclassification, not a format correction.
                if (!string.Equals(printable, replacement, StringComparison.Ordinal) &&
                    ComponentTextNormalizer.IsSafeWholeParagraphReplacement(printable, replacement))
                    edits.Add(new SpellingEdit(paragraph.AbsoluteStart, printable.Length,
                        replacement, printable, 100));
            }

            foreach (var finding in findings)
            {
                if (componentFindingCodes.Contains(finding.RuleCode)) continue;
                if (finding.Anchor.Kind != AnnotationAnchorKind.TextSpan ||
                    !finding.Anchor.ParagraphIndex.HasValue || !finding.Anchor.StartOffset.HasValue ||
                    !string.Equals(finding.Anchor.StoryType, "wdMainTextStory", StringComparison.Ordinal))
                    continue;
                LocalParagraphSnapshot paragraph;
                if (!paragraphs.TryGetValue(finding.Anchor.ParagraphIndex.Value, out paragraph)) continue;
                var edit = CreateSpellingEdit(finding, paragraph, rules, lexicon);
                if (edit != null) edits.Add(edit);
            }

            if (includeRepeatedSpaces)
                AddRepeatedSpaceEdits(snapshot, edits);

            var applied = new List<SpellingEdit>();
            var count = 0;
            int? contentEnd = null;
            foreach (var edit in edits.OrderByDescending(item => item.Start).ThenByDescending(item => item.Priority))
            {
                // Edits are applied from the end of the document; a sorted list of
                // already-applied starts makes the overlap test logarithmic-ish in
                // practice (only the closest applied edits can intersect).
                if (IntersectsApplied(applied, edit)) continue;
                Word.Range? range = null;
                try
                {
                    var tracker = s_tracker;
                    var start = tracker == null ? edit.Start : tracker.Map(edit.Start);
                    if (!string.IsNullOrEmpty(edit.ExpectedText))
                    {
                        var validationEnd = checked(start + edit.ExpectedText.Length);
                        if (!contentEnd.HasValue)
                        {
                            Word.Range? content = null;
                            try { content = document.Content; contentEnd = content.End; }
                            finally { Release(content); }
                        }
                        if (validationEnd > contentEnd.Value) continue;
                        Word.Range? validation = null;
                        try
                        {
                            validation = document.Range(start, validationEnd);
                            if (!string.Equals(validation.Text ?? string.Empty, edit.ExpectedText,
                                    StringComparison.Ordinal))
                                continue;
                        }
                        finally { Release(validation); }
                    }
                    range = document.Range(start, checked(start + edit.Length));
                    if (range.Editors.Count > 0 || range.Fields.Count > 0 ||
                        range.ContentControls.Count > 0 || range.InlineShapes.Count > 0)
                        continue;
                    if (edit.Length == 0) range.InsertBefore(edit.Replacement);
                    else range.Text = edit.Replacement;
                    RecordTextEdit(start, edit.Length, edit.Replacement.Length);
                    contentEnd = contentEnd.HasValue
                        ? contentEnd.Value + edit.Replacement.Length - edit.Length
                        : (int?)null;
                    applied.Add(edit);
                    count++;
                }
                finally { Release(range); }
            }
            return count;
        }

        private static bool IntersectsApplied(List<SpellingEdit> applied, SpellingEdit edit)
        {
            // Applied edits are appended in descending start order; only the most
            // recent ones can reach back to this edit's span.
            for (var index = applied.Count - 1; index >= 0; index--)
            {
                var item = applied[index];
                if (Intersects(edit, item)) return true;
                if (item.Start > edit.Start + Math.Max(edit.Length, 1) + 4096) break;
            }
            return false;
        }

        /// <summary>
        /// Replaces runs of two or more spaces by one space inside ordinary
        /// main-story paragraphs. Table cells (often aligned with spaces),
        /// protected spans, fields, hyperlinks and content controls are skipped.
        /// </summary>
        private static void AddRepeatedSpaceEdits(LocalScanSnapshot snapshot, ICollection<SpellingEdit> edits)
        {
            foreach (var paragraph in snapshot.Paragraphs)
            {
                if (!string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) ||
                    paragraph.IsInTable || paragraph.HasField == true || paragraph.HasHyperlink == true ||
                    paragraph.HasContentControl == true || paragraph.HasMathObject == true ||
                    paragraph.Text.IndexOf("  ", StringComparison.Ordinal) < 0 ||
                    snapshot.IntersectsProtectedSpan(paragraph))
                    continue;
                foreach (Match match in Regex.Matches(paragraph.Text, @" {2,}",
                    RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200)))
                {
                    edits.Add(new SpellingEdit(checked(paragraph.AbsoluteStart + match.Index),
                        match.Length, " ", match.Value, 5));
                }
            }
        }

        private static string ResolveExpectedTypeAbbreviation(LocalScanSnapshot snapshot,
            int paragraphIndex, LocalRulePack rules,
            IReadOnlyList<LogicalDocumentBlock> blocks)
        {
            var block = blocks.FirstOrDefault(item => item.ContainsParagraph(paragraphIndex));
            if (block == null) return string.Empty;
            var typeParagraph = snapshot.Paragraphs.FirstOrDefault(item =>
                block.Roles.TryGetValue(item.Index, out var role) && role == "typeName");
            if (typeParagraph == null) return string.Empty;
            var typeName = Regex.Replace(typeParagraph.Text ?? string.Empty, @"\s+", " ")
                .Trim().Trim('.', ':');
            var entry = rules.DocumentTypeAbbreviations.FirstOrDefault(item =>
                string.Equals(item.TypeName, typeName, StringComparison.OrdinalIgnoreCase));
            return entry == null ? string.Empty : entry.Abbreviation;
        }

        private static bool IsDeterministicComponentTextFinding(string ruleCode)
        {
            switch (ruleCode)
            {
                case "ND30-PL1-M2-K1-TN-SEP":
                case "ND30-PL1-M2-K3-PREFIX":
                case "ND30-PL1-M2-K3-SEP":
                case "ND30-PL1-M2-K3-ABBR":
                case "ND30-PL1-M2-K3-SPACE":
                case "ND30-PL1-M2-K3-CASE":
                case "ND30-PL1-M2-K3-PAD":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Locates a snapshot paragraph in the current document from its absolute
        /// offsets (mapped through the text edits made since the snapshot) and
        /// verifies both paragraph boundaries. Only when that check fails does it
        /// fall back to Paragraphs[index], which is O(paragraphs) per call.
        /// </summary>
        private static Word.Range ResolveCurrentMainParagraphRange(
            Word.Document document,
            LocalParagraphSnapshot paragraph)
        {
            if (!string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) ||
                paragraph.Index <= 0)
                throw new InvalidOperationException(
                    "Vị trí đoạn văn đã thay đổi. Hãy bấm lại chức năng để phân tích trạng thái mới.");

            var tracker = s_tracker;
            var start = tracker == null ? paragraph.AbsoluteStart : tracker.Map(paragraph.AbsoluteStart);
            var end = tracker == null ? paragraph.AbsoluteEnd : tracker.Map(paragraph.AbsoluteEnd);
            Word.Range? probe = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? containing = null;
            Word.Range? candidate = null;
            try
            {
                probe = document.Range(start, start);
                paragraphs = probe.Paragraphs;
                containing = paragraphs.First;
                candidate = containing.Range.Duplicate;
                if (candidate.Start == start && candidate.End == end)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(candidate);
                Release(containing);
                Release(paragraphs);
                Release(probe);
            }
            return ResolveMainParagraphRangeByIndex(document, paragraph);
        }

        private static Word.Range ResolveMainParagraphRangeByIndex(
            Word.Document document,
            LocalParagraphSnapshot paragraph)
        {
            if (paragraph.Index > document.Paragraphs.Count)
                throw new InvalidOperationException(
                    "Vị trí đoạn văn đã thay đổi. Hãy bấm lại chức năng để phân tích trạng thái mới.");
            Word.Paragraph? current = null;
            try
            {
                current = document.Paragraphs[paragraph.Index];
                return current.Range.Duplicate;
            }
            finally { Release(current); }
        }

        /// <summary>
        /// Records every text change (in current document coordinates, in the
        /// order applied) and maps an offset of the snapshot to the current document.
        /// </summary>
        private sealed class MainStoryEditTracker
        {
            private readonly List<TrackedEdit> _edits = new List<TrackedEdit>();

            public void Record(int currentStart, int oldLength, int newLength)
            {
                if (oldLength == newLength) return;
                _edits.Add(new TrackedEdit(currentStart, oldLength, newLength - oldLength));
            }

            public int Map(int original)
            {
                var position = original;
                foreach (var edit in _edits)
                {
                    // A paragraph boundary at the start of an insertion stays put; a
                    // boundary at or after the end of a replaced span moves.
                    var shifts = edit.OldLength == 0
                        ? position > edit.Start
                        : position >= edit.Start + edit.OldLength;
                    if (shifts) position += edit.Delta;
                }
                return position;
            }

            private sealed class TrackedEdit
            {
                public TrackedEdit(int start, int oldLength, int delta)
                {
                    Start = start;
                    OldLength = oldLength;
                    Delta = delta;
                }

                public int Start { get; }
                public int OldLength { get; }
                public int Delta { get; }
            }
        }

        private static void RecordTextEdit(int currentStart, int oldLength, int newLength)
        {
            s_tracker?.Record(currentStart, oldLength, newLength);
        }

        private static Nd30HeaderFontSizeTier ResolveHeaderTierCached(LocalScanSnapshot snapshot,
            IReadOnlyList<LogicalDocumentBlock> blocks, int paragraphIndex,
            IDictionary<int, Nd30HeaderFontSizeTier> cache)
        {
            var blockStart = -1;
            foreach (var block in blocks)
                if (block.ContainsParagraph(paragraphIndex))
                {
                    blockStart = block.StartParagraphIndex;
                    break;
                }
            Nd30HeaderFontSizeTier tier;
            if (cache.TryGetValue(blockStart, out tier)) return tier;
            tier = ResolveHeaderTier(snapshot, blocks, paragraphIndex);
            cache[blockStart] = tier;
            return tier;
        }

        private static SpellingEdit? CreateSpellingEdit(AnnotationFinding finding,
            LocalParagraphSnapshot paragraph, LocalRulePack rules,
            VietnameseLexiconSpellChecker lexicon)
        {
            var expectedText = finding.Anchor.ExpectedText ?? string.Empty;
            var start = checked(paragraph.AbsoluteStart + finding.Anchor.StartOffset.GetValueOrDefault());
            var length = finding.Anchor.Length.GetValueOrDefault();
            switch (finding.RuleCode)
            {
                case "FORMAT-CODE-NUMBER-PAD":
                case "ND30-PL1-M2-K3-PAD":
                case "FORMAT-PLACE-DATE-PAD":
                case "ND30-PL1-M2-K4-PAD":
                    return expectedText.Length == 1 && char.IsDigit(expectedText[0])
                        ? new SpellingEdit(start, length, "0" + expectedText, expectedText, 35)
                        : null;
                case "FORMAT-CODE-NOTATION-UPPERCASE":
                case "SPELLING-ABBREVIATION-UPPERCASE":
                case "ND30-PL1-M2-K3-CASE":
                case "ND30-PL1-M2-K7B-AUTH":
                case "ND30-PL2-M3-K1B":
                    return new SpellingEdit(start, length,
                        expectedText.ToUpper(CultureInfo.GetCultureInfo("vi-VN")), expectedText, 35);
                case "FORMAT-PLACE-DATE-COMMA":
                case "ND30-PL1-M2-K4-COMMA":
                    return new SpellingEdit(start, length, expectedText + ",", expectedText, 35);
                case "ND30-PL1-M2-K4-CASE":
                case "ND30-PL2-M2-K1":
                    return new SpellingEdit(start, length, ToVietnameseTitleCase(expectedText), expectedText, 30);
                case "SPELLING-SENTENCE-CAPITALIZATION":
                case "ND30-PL2-M1":
                    return expectedText.Length == 0 ? null : new SpellingEdit(start, length,
                        UppercaseFirst(expectedText), expectedText, 30);
                case "ND30-PL2-M3-K1A":
                case "ND30-PL2-M3-K1C":
                case "ND30-PL2-M3-K1D":
                case "ND30-PL2-M3-K1E":
                case "ND30-PL2-M4-K1A":
                case "ND30-PL2-M4-K1B":
                case "ND30-PL2-M5-K5":
                case "ND30-PL2-M5-K8A":
                    var configuredCapitalization = ResolveTargetReplacement(finding, rules);
                    return string.IsNullOrWhiteSpace(configuredCapitalization) ? null :
                        new SpellingEdit(start, length, configuredCapitalization!, expectedText, 30);
                case "ND30-PL2-M5-K7":
                    var structuralReplacement = finding.Expected.IndexOf("Viết thường", StringComparison.OrdinalIgnoreCase) >= 0
                        ? LowercaseFirst(expectedText)
                        : UppercaseFirst(expectedText);
                    return new SpellingEdit(start, length, structuralReplacement, expectedText, 30);
                case "ND30-PL1-M2-K6A-PUNCT":
                    var requiredPunctuation = finding.Expected.IndexOf("chấm phẩy", StringComparison.OrdinalIgnoreCase) >= 0
                        ? ";"
                        : ".";
                    var replacesPunctuation = expectedText.Length == 1 &&
                        ".;,:".IndexOf(expectedText[0]) >= 0;
                    return replacesPunctuation
                        ? new SpellingEdit(start, length, requiredPunctuation, expectedText, 15)
                        : new SpellingEdit(start + length, 0, requiredPunctuation, string.Empty, 15);
                case "LOCAL-TYPO-PUNCT":
                    return new SpellingEdit(start, length, string.Empty,
                        expectedText, 50);
                case "LOCAL-TYPO-SPACE":
                    return new SpellingEdit(start, length, " ",
                        expectedText, 40);
                case "LOCAL-TYPO-HIDDEN":
                    return new SpellingEdit(start, length, string.Empty,
                        expectedText, 40);
                case "LOCAL-TYPO-DICT":
                    var correction = rules.Corrections.FirstOrDefault(item =>
                        string.Equals(item.Wrong, expectedText, StringComparison.OrdinalIgnoreCase));
                    if (correction != null)
                    {
                        return new SpellingEdit(start,
                            length, ApplyCase(expectedText, correction.Replacement),
                            expectedText, 30);
                    }
                    if (ChuanHoa.Client.Core.Lexicon.VietnameseConfusionSets.AdministrativeConfusionPairs.TryGetValue(expectedText, out var adminRep))
                    {
                        return new SpellingEdit(start,
                            length, ApplyCase(expectedText, adminRep),
                            expectedText, 30);
                    }
                    return null;
                case "LOCAL-TYPO-LEXICON":
                    var suggestion = lexicon.FindDeterministicCorrection(expectedText);
                    return suggestion == null ? null : new SpellingEdit(start,
                        length, suggestion, expectedText, 25);
                case "LOCAL-TYPO-TELEX":
                    var telexRule = rules.TelexRules?.FirstOrDefault(item =>
                        Regex.IsMatch(expectedText, item.Pattern, RegexOptions.IgnoreCase));
                    if (telexRule != null)
                    {
                        var rep = Regex.Replace(expectedText, telexRule.Pattern, telexRule.Replacement, RegexOptions.IgnoreCase);
                        return new SpellingEdit(start, length, ApplyCase(expectedText, rep), expectedText, 25);
                    }
                    break;
                case "ND30-PL1-M2-K6B-DATE":
                    return new SpellingEdit(start, 0, "ngày ", expectedText, 10);
                case "ND30-PL1-M2-K6B-SO":
                    var citedType = Regex.Match(expectedText,
                        @"^\s*(?:nghị\s+quyết|nghị\s+định|quyết\s+định|chỉ\s+thị|thông\s+tư)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                        TimeSpan.FromMilliseconds(200));
                    if (!citedType.Success) return null;
                    var insertionOffset = citedType.Index + citedType.Length;
                    return new SpellingEdit(start + insertionOffset, 0, " số",
                        expectedText.Substring(insertionOffset), 15);
                case "ND30-PL1-M2-K6E-DOTSLASH":
                    return new SpellingEdit(start, length,
                        NormalizeFinalPunctuation(expectedText), expectedText, 20);
            }

            var target = ResolveTargetReplacement(finding, rules);
            if (!string.IsNullOrEmpty(target) && !string.Equals(target, expectedText, StringComparison.Ordinal))
            {
                return new SpellingEdit(start, length, target!, expectedText, 30);
            }

            return null;
        }

        private static string? ResolveTargetReplacement(AnnotationFinding finding, LocalRulePack rules)
        {
            var expectedText = finding.Anchor.ExpectedText ?? string.Empty;
            switch (finding.RuleCode)
            {
                case "ND30-PL1-M2-K6A-PUNCT":
                    return finding.Expected.IndexOf("chấm phẩy", StringComparison.OrdinalIgnoreCase) >= 0
                        ? ";"
                        : ".";
                case "LOCAL-TYPO-SPACE":
                    return " ";
                case "LOCAL-TYPO-PUNCT":
                case "LOCAL-TYPO-HIDDEN":
                    return string.Empty;
                case "ND30-PL1-M2-K6B-DATE":
                    return "ngày " + expectedText;
                case "ND30-PL2-M5-K7":
                    return expectedText.ToLower(CultureInfo.GetCultureInfo("vi-VN"));
            }

            if (string.Equals(finding.RuleCode, "LOCAL-TYPO-TELEX", StringComparison.OrdinalIgnoreCase) && rules.TelexRules != null)
            {
                var telexRule = rules.TelexRules.FirstOrDefault(item =>
                    Regex.IsMatch(expectedText, item.Pattern, RegexOptions.IgnoreCase));
                if (telexRule != null)
                {
                    var rep = Regex.Replace(expectedText, telexRule.Pattern, telexRule.Replacement, RegexOptions.IgnoreCase);
                    return ApplyCase(expectedText, rep);
                }
            }

            if (string.Equals(finding.RuleCode, "LOCAL-TYPO-LEXICON", StringComparison.OrdinalIgnoreCase) && rules.Lexicon != null)
            {
                var lexicon = new VietnameseLexiconSpellChecker(rules.Lexicon);
                var suggestion = lexicon.FindDeterministicCorrection(expectedText);
                if (suggestion != null) return suggestion;
            }

            if (rules.Capitalizations != null && rules.Capitalizations.Count > 0)
            {
                var cap = rules.Capitalizations.FirstOrDefault(item =>
                    string.Equals(item.Expected, expectedText, StringComparison.OrdinalIgnoreCase));
                if (cap != null && !string.Equals(cap.Expected, expectedText, StringComparison.Ordinal))
                    return cap.Expected;
            }

            if (!string.IsNullOrWhiteSpace(finding.Expected))
            {
                var quoteMatch = Regex.Match(finding.Expected,
                    @"(?:Viết|Sửa thành|thành|Dùng|Nên dùng)\s+[“""]([^”""]+)[”""]",
                    RegexOptions.IgnoreCase);
                if (quoteMatch.Success)
                {
                    var extracted = quoteMatch.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(extracted) && !string.Equals(extracted, expectedText, StringComparison.Ordinal))
                        return extracted;
                }
            }

            if (rules.Corrections != null)
            {
                var corr = rules.Corrections.FirstOrDefault(item =>
                    string.Equals(item.Wrong, expectedText, StringComparison.OrdinalIgnoreCase));
                if (corr != null) return ApplyCase(expectedText, corr.Replacement);
            }

            if (ChuanHoa.Client.Core.Lexicon.VietnameseConfusionSets.AdministrativeConfusionPairs.TryGetValue(expectedText, out var adminRep2))
            {
                return ApplyCase(expectedText, adminRep2);
            }

            return null;
        }

        private static bool SelectedFixChangesText(string lane, AnnotationFinding finding)
        {
            if (string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase)) return true;
            var ruleCode = finding.RuleCode;
            if (string.Equals(ruleCode, "ND30-PL1-M2-K6A-STYLE", StringComparison.Ordinal))
                return finding.CurrentIssue.IndexOf("thiếu gạch ngang",
                    StringComparison.OrdinalIgnoreCase) >= 0;
            if (string.Equals(ruleCode, "ND30-PL1-M2-K9B-LABEL", StringComparison.Ordinal))
                return finding.CurrentIssue.IndexOf("thiếu dấu hai chấm",
                    StringComparison.OrdinalIgnoreCase) >= 0;
            if (string.Equals(ruleCode, "ND30-PL1-M2-K9B-LIST", StringComparison.Ordinal))
                return finding.CurrentIssue.IndexOf("sai gạch đầu dòng",
                    StringComparison.OrdinalIgnoreCase) >= 0;
            if (string.Equals(ruleCode, "ND30-PL1-M3-K1B", StringComparison.Ordinal))
                return finding.CurrentIssue.IndexOf("chưa viết hoa",
                    StringComparison.OrdinalIgnoreCase) >= 0;
            return string.Equals(ruleCode, "ND30-PL1-M2-K6A-PUNCT", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K6B-SO", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K4-COMMA", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K4-CASE", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K4-PAD", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K5B-SPACE", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K6E-DOTSLASH", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K7B-AUTH", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K9A-COLON", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K9A-INLINE-END", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K9A-PUNCT", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M2-K9B-LUU", StringComparison.Ordinal) ||
                string.Equals(ruleCode, "ND30-PL1-M3-K1A-NUM", StringComparison.Ordinal) ||
                IsDeterministicComponentTextFinding(ruleCode);
        }

        private static void UpdateContextAfterTargetedFix(DocumentContext context, string lane,
            string findingId, bool preserveSnapshot)
        {
            if (!preserveSnapshot)
            {
                // Text replacement can shift every downstream absolute coordinate.
                // Discard that positional snapshot instead of silently reusing it.
                context.ClearReadAnalysis();
                return;
            }

            try
            {
                if (context.LastSnapshot == null || context.LastLocalSnapshot == null)
                {
                    context.ClearReadAnalysis();
                    return;
                }

                var format = context.LastFormatScan;
                var spelling = context.LastSpellingScan;
                if (string.Equals(lane, "format", StringComparison.OrdinalIgnoreCase) && format != null)
                    format = CopyWithoutFinding(format, findingId);
                else if (string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase) && spelling != null)
                    spelling = CopyWithoutFinding(spelling, findingId);
                else
                {
                    context.ClearReadAnalysis();
                    return;
                }

                context.SetAnalysis(context.LastSnapshot, context.LastLocalSnapshot,
                    format, spelling, false);
            }
            catch
            {
                context.ClearReadAnalysis();
            }
        }

        private static LocalScanResult CopyWithoutFinding(LocalScanResult source, string findingId)
        {
            return new LocalScanResult(
                source.ScanId,
                source.Lane,
                source.RulePackId,
                source.DocumentFingerprint,
                source.Revision,
                source.Findings.Where(item =>
                    !string.Equals(item.FindingId, findingId, StringComparison.Ordinal)).ToArray(),
                source.RulePackVersion,
                source.DetectorPolicyVersion,
                source.NotEvaluatedRuleCodes,
                source.AcademicTypographyEnabled,
                source.AcademicHeadingCount,
                source.HeadingLevel1Count,
                source.HeadingLevel2Count,
                source.HeadingLevel3Count);
        }

        private static string ApplyCase(string source, string target)
        {
            var vietnamese = CultureInfo.GetCultureInfo("vi-VN");
            if (source.All(character => !char.IsLetter(character) || char.IsUpper(character)))
                return target.ToUpper(vietnamese);
            if (source.Length > 0 && char.IsUpper(source[0]))
                return char.ToUpper(target[0], vietnamese) + target.Substring(1);
            return target;
        }

        private static string UppercaseFirst(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            var vietnamese = CultureInfo.GetCultureInfo("vi-VN");
            return char.ToUpper(value[0], vietnamese) + value.Substring(1);
        }

        private static string LowercaseFirst(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            var vietnamese = CultureInfo.GetCultureInfo("vi-VN");
            return char.ToLower(value[0], vietnamese) + value.Substring(1);
        }

        private static string ToVietnameseTitleCase(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return value;
            return CultureInfo.GetCultureInfo("vi-VN").TextInfo.ToTitleCase(
                value.ToLower(CultureInfo.GetCultureInfo("vi-VN")));
        }

        private static string NormalizeFinalPunctuation(string value)
        {
            if (string.IsNullOrEmpty(value)) return ".";
            var trimmed = value.TrimEnd();
            if (trimmed.EndsWith("./.", StringComparison.Ordinal) ||
                trimmed.EndsWith(". / .", StringComparison.Ordinal))
            {
                var marker = trimmed.EndsWith("./.", StringComparison.Ordinal) ? "./." : ". / .";
                return trimmed.Substring(0, trimmed.Length - marker.Length) + ".";
            }
            if (";,:!?/".IndexOf(trimmed[trimmed.Length - 1]) >= 0)
                return trimmed.Substring(0, trimmed.Length - 1) + ".";
            return trimmed.EndsWith(".", StringComparison.Ordinal) ? trimmed : trimmed + ".";
        }

        private static bool Intersects(SpellingEdit left, SpellingEdit right)
        {
            if (left.Length == 0 || right.Length == 0) return left.Start == right.Start;
            return left.Start < right.Start + right.Length && right.Start < left.Start + left.Length;
        }

        private static void RemoveTrailingBlankParagraphs(Word.Document document)
        {
            WordTrailingBlankPageCleaner.Remove(document);
        }

        private static void NormalizeBodySpacing(Word.Range range, bool party)
        {
            range.ParagraphFormat.SpaceBeforeAuto = 0;
            range.ParagraphFormat.SpaceAfterAuto = 0;
            range.ParagraphFormat.SpaceBefore = 0f;
            range.ParagraphFormat.SpaceAfter = 6f;
            if (party)
            {
                range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                range.ParagraphFormat.LineSpacing = 18f;
            }
            else NormalizeNd30ContentLineSpacing(range);
        }

        private static void NormalizeNd30ContentLineSpacing(Word.Range range)
        {
            // Normalization preference, not a stricter legal validation bound.
            // Word represents Multiple 1.2 as 12 * 1.2 points, not font size * 1.2.
            range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
            range.ParagraphFormat.LineSpacing = 14.4f;
        }

        private static bool IsBodyParagraph(LocalParagraphSnapshot paragraph)
        {
            var text = paragraph.Text.Trim();
            // Length and list-marker heuristics used to skip legitimate short
            // body paragraphs ("Điều 1.", "1. ...", "a) ..."). Structural
            // roles are resolved before this fallback; every remaining
            // non-empty main-text paragraph must receive the canonical body
            // spacing instead of inheriting arbitrary direct/style formatting.
            if (text.Length == 0 || text.Length > 5000) return false;
            if (text.Where(char.IsLetter).Any() && text.Where(char.IsLetter).All(char.IsUpper)) return false;
            return true;
        }

        private static bool IsLegalPointParagraph(string text)
        {
            return Regex.IsMatch(text ?? string.Empty, @"^\s*[a-zđ]\)\s+\p{L}",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
        }

        private static string OwnedLineName(string ruleCode, int paragraphIndex)
        {
            // The binary .doc format rejects long Shape.Name values. Keep the
            // ownership marker short and ASCII for Word 2010 compatibility.
            return OwnedLinePrefix(ruleCode) + "P" + paragraphIndex.ToString(CultureInfo.InvariantCulture);
        }

        private static bool IsOwnedLineForParagraph(string name, int paragraphIndex)
        {
            return LineShapeOwnership.IsOwnedForParagraph(name, paragraphIndex);
        }

        private static string OwnedLinePrefix(string ruleCode)
        {
            var component = ruleCode.IndexOf("ORG-LINE", StringComparison.Ordinal) >= 0 ? "ORG" :
                ruleCode.IndexOf("SUBJ-LINE", StringComparison.Ordinal) >= 0 ? "SUBJ" :
                ruleCode.IndexOf("TITLE-LINE", StringComparison.Ordinal) >= 0 ? "PARTY" : "MOTTO";
            return "CHUANHOA2_" + component + "_";
        }

        private static float ValidOr(double? value, double minimum, double maximum, double fallback) =>
            (float)(value.HasValue && value.Value >= minimum && value.Value <= maximum ? value.Value : fallback);

        private static bool IsParty(LocalScanSnapshot snapshot) =>
            string.Equals(snapshot.RegimeCode, "PARTY_HD05", StringComparison.OrdinalIgnoreCase);

        private static ParagraphStyle S(double size, bool bold, bool italic,
            Word.WdParagraphAlignment alignment, bool underline = false) =>
            new ParagraphStyle((float)size, bold, italic, underline, alignment);

        private static string CreateBackup(Word.Document document)
        {
            return WordRecoveryCopyManager.Create(document.Application, document, "one-click");
        }

        private int WordMajorVersion()
        {
            var value = _application.Version ?? string.Empty;
            var dot = value.IndexOf('.');
            int major;
            return int.TryParse(dot < 0 ? value : value.Substring(0, dot), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out major) ? major : 0;
        }

        private static bool ContainsPageField(Word.Section section)
        {
            Word.HeadersFooters? headers = null;
            Word.HeadersFooters? footers = null;
            try
            {
                headers = section.Headers;
                if (ContainsPageField(headers)) return true;
                footers = section.Footers;
                return ContainsPageField(footers);
            }
            finally { Release(footers); Release(headers); }
        }

        private static bool ContainsPageField(Word.HeadersFooters collection)
        {
            foreach (Word.WdHeaderFooterIndex index in new[]
            {
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage
            })
            {
                Word.HeaderFooter? item = null;
                Word.Range? range = null;
                Word.Fields? fields = null;
                try
                {
                    item = collection[index];
                    if (!item.Exists) continue;
                    range = item.Range;
                    fields = range.Fields;
                    for (var fieldIndex = 1; fieldIndex <= fields.Count; fieldIndex++)
                    {
                        Word.Field? field = null;
                        try
                        {
                            field = fields[fieldIndex];
                            if (field.Type == Word.WdFieldType.wdFieldPage) return true;
                        }
                        finally { Release(field); }
                    }
                }
                catch (COMException) { }
                finally { Release(fields); Release(range); Release(item); }
            }
            return false;
        }

        private static void RemoveFirstPageNumber(Word.Section section)
        {
            Word.HeadersFooters? headers = null;
            Word.HeadersFooters? footers = null;
            try
            {
                headers = section.Headers;
                footers = section.Footers;
                RemovePageFields(headers, Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage);
                RemovePageFields(footers, Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage);
            }
            finally { Release(footers); Release(headers); }
        }

        private static void RemoveFooterPageNumbers(Word.Section section)
        {
            Word.HeadersFooters? footers = null;
            try
            {
                footers = section.Footers;
                foreach (Word.WdHeaderFooterIndex index in new[]
                {
                    Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                    Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                    Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage
                }) RemovePageFields(footers, index);
            }
            finally { Release(footers); }
        }

        private static void RemovePageFields(Word.HeadersFooters collection, Word.WdHeaderFooterIndex index)
        {
            Word.HeaderFooter? item = null;
            Word.Range? range = null;
            Word.Fields? fields = null;
            try
            {
                item = collection[index];
                if (!item.Exists) return;
                range = item.Range;
                fields = range.Fields;
                for (var fieldIndex = fields.Count; fieldIndex >= 1; fieldIndex--)
                {
                    Word.Field? field = null;
                    try
                    {
                        field = fields[fieldIndex];
                        if (field.Type == Word.WdFieldType.wdFieldPage) field.Delete();
                    }
                    finally { Release(field); }
                }
            }
            catch (COMException) { }
            finally { Release(fields); Release(range); Release(item); }
        }

        private static bool IsLikelyDataTable(Word.Table table)
        {
            if (table.Rows.Count < 2 || table.Columns.Count < 2) return false;
            Word.Row? first = null;
            try
            {
                first = table.Rows[1];
                var nonEmpty = 0;
                var boldCells = 0;
                for (var index = 1; index <= first.Cells.Count; index++)
                {
                    Word.Cell? cell = null;
                    Word.Range? range = null;
                    try
                    {
                        cell = first.Cells[index];
                        range = cell.Range.Duplicate;
                        var text = (range.Text ?? string.Empty).Trim('\r', '\a', ' ', '\t');
                        if (text.Length == 0) continue;
                        nonEmpty++;
                        if (range.Font.Bold != 0) boldCells++;
                    }
                    finally { Release(range); Release(cell); }
                }
                return nonEmpty >= 2 && boldCells == nonEmpty;
            }
            catch (COMException)
            {
                return false;
            }
            finally { Release(first); }
        }

        private void RestoreWorkingDocumentAfterFailure(Word.Document document, string backupPath, bool customUndoAvailable)
        {
            if (customUndoAvailable)
            {
                try
                {
                    object count = 1;
                    document.Undo(ref count);
                    return;
                }
                catch (COMException) { }
            }

            // Word 2010 has no custom undo record. The source was required to be saved
            // before execution, so closing without saving and reopening is the only
            // deterministic way to avoid leaving a partially modified document open.
            try
            {
                var sourcePath = document.FullName;
                object doNotSave = Word.WdSaveOptions.wdDoNotSaveChanges;
                document.Close(ref doNotSave);
                object fileName = sourcePath;
                object readOnly = false;
                object addToRecent = false;
                _application.Documents.Open(ref fileName, ReadOnly: ref readOnly, AddToRecentFiles: ref addToRecent);
            }
            catch (COMException)
            {
                // The independently copied backup remains available even if Word cannot reopen the source.
                if (!File.Exists(backupPath)) throw;
            }
        }

        private static Nd30HeaderFontSizeTier ResolveHeaderTier(LocalScanSnapshot snapshot,
            IReadOnlyList<LogicalDocumentBlock> blocks, int paragraphIndex)
        {
            var block = blocks.FirstOrDefault(item => item.ContainsParagraph(paragraphIndex));
            if (block != null)
                return Nd30HeaderFontSizeTierResolver.Resolve(snapshot, block.Roles,
                    block.StartParagraphIndex, block.EndParagraphIndex);
            return Nd30HeaderFontSizeTierResolver.Resolve(snapshot,
                new Dictionary<int, string>());
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        private sealed class ParagraphStyle
        {
            public ParagraphStyle(float size, bool bold, bool italic, bool underline,
                Word.WdParagraphAlignment alignment)
            {
                Size = size; Bold = bold; Italic = italic; Underline = underline; Alignment = alignment;
            }
            public float Size { get; }
            public bool Bold { get; }
            public bool Italic { get; }
            public bool Underline { get; }
            public Word.WdParagraphAlignment Alignment { get; }
        }

        private sealed class SpellingEdit
        {
            public SpellingEdit(int start, int length, string replacement, string expectedText, int priority)
            {
                Start = start;
                Length = length;
                Replacement = replacement;
                ExpectedText = expectedText;
                Priority = priority;
            }
            public int Start { get; }
            public int Length { get; }
            public string Replacement { get; }
            public string ExpectedText { get; }
            public int Priority { get; }
        }
    }
}
