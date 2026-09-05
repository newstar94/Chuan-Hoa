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
                annotations.ClearLane("format");
                annotations.ClearLane("spelling");

                // Apply text edits first while every explicit-read offset is still exact.
                // Later layout operations resolve main-story paragraphs by their stable
                // paragraph index, so replacements that change character counts cannot
                // shift formatting onto another paragraph.
                correctedSpellingItems = ApplyDeterministicSpellingFixes(document, local,
                    formatFindings.Concat(spellingFindings).ToArray(), rules, blocks);
                correctedSpellingItems += CollapseMultipleSpaces(document);
                changedParagraphs += ApplyDeterministicParagraphFindingFixes(
                    document, local, formatFindings, roles);
                changedParagraphs += ApplyExplicitFormatFindingFixes(document, local,
                    formatFindings, rules, blocks, roles);
                normalizedSections = NormalizeSections(document, local, rules);
                NormalizeHeaderLayoutTables(document, local, roles);
                foreach (var paragraph in local.Paragraphs.Where(p =>
                    string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal)))
                {
                    string role;
                    roles.TryGetValue(paragraph.Index, out role);
                    var headerFontTier = ResolveHeaderTier(local, blocks, paragraph.Index);
                    if (ApplyParagraphFormat(document, paragraph, role ?? string.Empty, local, rules,
                            headerFontTier))
                        changedParagraphs++;
                }

                insertedLines = InsertMissingRequiredLines(document, local, formatFindings,
                    blocks, roles);
                EnsurePageNumbers(document, local, rules);
                normalizedTables = NormalizeTables(document, local, roles);
                WordAppendixPaginationNormalizer.Normalize(document, roles);
                RemoveTrailingBlankParagraphs(document);

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
                _application.DisplayAlerts = previousAlerts;
                _application.ScreenUpdating = previousScreenUpdating;
            }
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

                var applied = string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase)
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
            section.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            section.PageSetup.TopMargin = (float)((party ? 20d : rules.TopMinMm) * PointsPerMillimeter);
            section.PageSetup.BottomMargin = (float)((party ? 20d : rules.BottomMinMm) * PointsPerMillimeter);
            section.PageSetup.LeftMargin = (float)((party ? 30d : rules.LeftMinMm) * PointsPerMillimeter);
            section.PageSetup.RightMargin = (float)((party ? 15d : rules.RightMinMm) * PointsPerMillimeter);
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
                    return range.Font.Bold == 0 && range.Font.Italic == 0;
                }
                if (string.IsNullOrEmpty(role))
                {
                    if (paragraph.IsInTable || !IsBodyParagraph(paragraph)) return false;
                    range.Font.Name = rules.BodyFontName;
                    range.Font.Color = Word.WdColor.wdColorAutomatic;
                    range.Font.Size = party ? 14f : ValidOr(paragraph.FontSizePoints, 13, 14, 14);
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    ApplyBodyParagraphIndent(range, paragraph.Text);
                    range.ParagraphFormat.SpaceAfter = 6f;
                    if (party)
                    {
                        range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                        range.ParagraphFormat.LineSpacing = 18f;
                    }
                    else range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    return true;
                }

                var style = StyleFor(role, party, paragraph, headerFontTier);
                if (style == null) return false;
                range.Font.Name = rules.BodyFontName;
                range.Font.Color = Word.WdColor.wdColorAutomatic;
                range.Font.Size = style.Size;
                range.Font.Bold = style.Bold ? -1 : 0;
                range.Font.Italic = style.Italic ? -1 : 0;
                range.Font.Underline = style.Underline
                    ? Word.WdUnderline.wdUnderlineSingle
                    : Word.WdUnderline.wdUnderlineNone;
                if (role == "appendixTitle")
                    range.Case = Word.WdCharacterCase.wdUpperCase;
                range.ParagraphFormat.Alignment = style.Alignment;
                if (role == "legalBasis")
                {
                    range.ParagraphFormat.FirstLineIndent = 10f * PointsPerMillimeter;
                    range.ParagraphFormat.SpaceAfter = 6f;
                }
                return true;
            }
            finally { Release(range); }
        }

        private static void ApplyBodyParagraphIndent(Word.Range range, string text)
        {
            Word.ParagraphFormat? format = null;
            Word.TabStops? tabStops = null;
            Word.TabStop? tabStop = null;
            try
            {
                format = range.ParagraphFormat;
                format.RightIndent = 0f;
                if (ParagraphIndentPolicy.IsDashListParagraph(text))
                {
                    tabStops = format.TabStops;
                    var textPosition = (float)(ParagraphIndentPolicy.ListTextMillimeters * PointsPerMillimeter);
                    var markerPosition = (float)(ParagraphIndentPolicy.ListMarkerMillimeters * PointsPerMillimeter);
                    format.LeftIndent = textPosition;
                    format.FirstLineIndent = markerPosition - textPosition;
                    if (!HasEquivalentTabStop(tabStops, textPosition))
                        tabStop = tabStops.Add(textPosition, Word.WdTabAlignment.wdAlignTabLeft,
                            Word.WdTabLeader.wdTabLeaderSpaces);
                    return;
                }

                format.LeftIndent = 0f;
                format.FirstLineIndent = (float)(ParagraphIndentPolicy.BodyFirstLineMillimeters * PointsPerMillimeter);
            }
            finally
            {
                Release(tabStop);
                Release(tabStops);
                Release(format);
            }
        }

        private static bool HasEquivalentTabStop(Word.TabStops tabStops, float position)
        {
            var count = tabStops.Count;
            for (var index = 1; index <= count; index++)
            {
                Word.TabStop? existing = null;
                try
                {
                    existing = tabStops[index];
                    if (Math.Abs(existing.Position - position) <= .5f &&
                        existing.Alignment == Word.WdTabAlignment.wdAlignTabLeft) return true;
                }
                catch (COMException) { }
                finally { Release(existing); }
            }
            return false;
        }

        private static ParagraphStyle? StyleFor(string role, bool party, LocalParagraphSnapshot paragraph,
            Nd30HeaderFontSizeTier headerFontTier)
        {
            switch (role)
            {
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
            IReadOnlyDictionary<int, string> roles)
        {
            var count = 0;
            var work = findings.Where(f => f.RuleCode.EndsWith("-LINE", StringComparison.Ordinal) &&
                    f.Anchor.ParagraphIndex.HasValue)
                .Select(f => Tuple.Create(f.RuleCode, f.Anchor.ParagraphIndex.GetValueOrDefault()))
                .ToList();

            // 1.0.0.43 could leave an OOXML line object that Word exposed through COM
            // but did not paint. Those owned shapes have the legacy CHUANHOA_* marker,
            // so migrate them once even when the geometry-only scanner calls them valid.
            // A valid user-created line is left untouched. Missing/invalid lines are
            // already represented by findings. Only lines explicitly owned by an older
            // Chuẩn hóa build are added here for a one-time migration.
            foreach (var line in snapshot.LineShapes.Where(item =>
                LineShapeOwnership.IsLegacyOwned(item.Name)))
            {
                var componentRole = line.Name.StartsWith("CHUANHOA_ORG_", StringComparison.Ordinal) ? "organName" :
                    line.Name.StartsWith("CHUANHOA_SUBJ_", StringComparison.Ordinal) ? "subject" :
                    line.Name.StartsWith("CHUANHOA_PARTY_", StringComparison.Ordinal) ? "partyTitle" :
                    line.Name.StartsWith("CHUANHOA_MOTTO_", StringComparison.Ordinal) ? "nationalMotto" : string.Empty;
                var ruleCode = componentRole == "organName" ? "ND30-PL1-M2-K2-ORG-LINE" :
                    componentRole == "subject" ? "ND30-PL1-M2-K5A-SUBJ-LINE" :
                    componentRole == "partyTitle" ? "HD05-M1-TITLE-LINE" :
                    componentRole == "nationalMotto" ? "ND30-PL1-M2-K1-TN-LINE" : string.Empty;
                if (ruleCode.Length == 0) continue;
                var paragraph = snapshot.Paragraphs.FirstOrDefault(item =>
                    roles.ContainsKey(item.Index) &&
                    string.Equals(roles[item.Index], componentRole, StringComparison.Ordinal) &&
                    IsAssociatedLine(line, item));
                if (paragraph != null && !work.Any(item => item.Item1 == ruleCode && item.Item2 == paragraph.Index))
                    work.Add(Tuple.Create(ruleCode, paragraph.Index));
            }

            foreach (var item in work)
            {
                var paragraphIndex = item.Item2;
                var paragraph = snapshot.Paragraphs.FirstOrDefault(p => p.Index == paragraphIndex);
                if (paragraph == null || !string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal)) continue;
                var ratio = item.Item1 == "ND30-PL1-M2-K2-ORG-LINE" ||
                    item.Item1 == "ND30-PL1-M2-K5A-SUBJ-LINE" ? .4d : 1d;
                if (NormalizeRequiredLine(document, snapshot, paragraph, ratio, item.Item1)) count++;
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
            LocalParagraphSnapshot paragraph, double ratio, string ruleCode)
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
                var inTable = range.get_Information(Word.WdInformation.wdWithInTable);
                if (TryNormalizeExistingComponentLine(document, snapshot, paragraph, ruleCode,
                        beginX, y, width))
                    return true;
                RemoveObsoleteComponentLines(document, snapshot, paragraph, ruleCode);
                lineAnchor = inTable ? ResolveNonTableLineAnchor(document, range) : range.Duplicate;
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
                shape.LockAnchor = 0;
                return true;
            }
            finally { Release(shape); Release(currentFont); Release(lineAnchor); Release(range); }
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
                    componentRange.Text = cleaned;
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
                    candidateRange.Text = string.Empty;
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
                    if ((int)candidate.Type != 9) continue;
                    var name = candidate.Name ?? string.Empty;
                    if (!string.Equals(name, selectedName, StringComparison.Ordinal))
                    {
                        if (deleteDuplicateCurrentOwned &&
                            string.Equals(name, ownedName, StringComparison.Ordinal) && normalized)
                            candidate.Delete();
                        continue;
                    }

                    candidate.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    candidate.RelativeVerticalPosition =
                        Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    candidate.Left = targetLeft;
                    candidate.Top = targetTop;
                    candidate.Width = targetWidth;
                    candidate.Height = 0f;
                    candidate.Name = ownedName;
                    NormalizeLineStyle(candidate);
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
            for (var index = document.Shapes.Count; index >= 1; index--)
            {
                Word.Shape? candidate = null;
                try
                {
                    candidate = document.Shapes[index];
                    if (string.Equals(candidate.Name, name, StringComparison.Ordinal)) candidate.Delete();
                }
                catch (COMException) { }
                finally { Release(candidate); }
            }
        }

        private static Word.Range ResolveNonTableLineAnchor(Word.Document document, Word.Range source)
        {
            var sourcePage = WordTextMeasurement.SafeInformation(source,
                Word.WdInformation.wdActiveEndAdjustedPageNumber).GetValueOrDefault();
            Word.Paragraphs? paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                var count = paragraphs.Count;
                for (var index = 1; index <= count; index++)
                {
                    Word.Paragraph? paragraph = null;
                    Word.Range? candidate = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        candidate = paragraph.Range.Duplicate;
                        if (candidate.get_Information(Word.WdInformation.wdWithInTable)) continue;
                        var page = WordTextMeasurement.SafeInformation(candidate,
                            Word.WdInformation.wdActiveEndAdjustedPageNumber).GetValueOrDefault();
                        if (sourcePage > 0d && page > 0d && Math.Abs(page - sourcePage) > .1d) continue;
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
                    if ((int)candidate.Type != 9) continue;
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

        private static void EnsurePageNumbers(Word.Document document, LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            var party = IsParty(snapshot);
            foreach (Word.Section section in document.Sections)
            {
                try
                {
                    EnsurePageNumber(document, section, party, rules);
                }
                finally { Release(section); }
            }
        }

        private static void EnsurePageNumber(Word.Document document, Word.Section section,
            bool party, LocalRulePack rules)
        {
            Word.HeaderFooter? header = null;
            Word.Range? range = null;
            Word.Field? field = null;
            try
            {
                section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                RemoveFirstPageNumber(section);
                if (!party) RemoveFooterPageNumbers(section);
                if (ContainsPageField(section)) return;
                header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (!header.Exists) header.Exists = true;
                range = header.Range.Duplicate;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                field = document.Fields.Add(range, Word.WdFieldType.wdFieldPage,
                    PreserveFormatting: true);
                field.Result.Font.Name = rules.BodyFontName;
                field.Result.Font.Size = 14f;
                field.Result.Font.Bold = 0;
                field.Result.Font.Italic = 0;
            }
            finally { Release(field); Release(range); Release(header); }
        }

        private static int NormalizeTables(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            var count = 0;
            foreach (Word.Table table in document.Tables)
            {
                Word.Row? first = null;
                try
                {
                    if (ContainsRecognizedRole(table, snapshot, roles)) continue;
                    if (!IsLikelyDataTable(table)) continue;
                    table.Rows.AllowBreakAcrossPages = 0;
                    first = table.Rows[1];
                    first.HeadingFormat = -1;
                    count++;
                }
                catch (COMException) { }
                finally { Release(first); Release(table); }
            }
            return count;
        }

        private static void NormalizeHeaderLayoutTables(Word.Document document, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            var headerRoles = new HashSet<string>(StringComparer.Ordinal)
            {
                "nationalTitle",
                "nationalMotto",
                "organName",
                "superiorOrganName",
                "partyTitle",
                "codeNumber",
                "placeAndIssuedDate"
            };

            var normalizedAny = false;
            foreach (Word.Table table in document.Tables)
            {
                try
                {
                    if (IsHeaderLayoutTable(table, snapshot, roles, headerRoles))
                    {
                        NormalizeHeaderTable(table);
                        normalizedAny = true;
                    }
                }
                catch (COMException) { }
                finally { Release(table); }
            }

            if (normalizedAny)
            {
                try { document.Repaginate(); } catch (COMException) { }
            }
        }

        private static bool IsHeaderLayoutTable(Word.Table table, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles, ISet<string> headerRoles)
        {
            Word.Range? range = null;
            try
            {
                range = table.Range;
                if (snapshot.Paragraphs.Any(paragraph =>
                    roles.ContainsKey(paragraph.Index) &&
                    headerRoles.Contains(roles[paragraph.Index]) &&
                    paragraph.AbsoluteStart >= range.Start && paragraph.AbsoluteStart < range.End))
                {
                    return true;
                }

                var text = range.Text ?? string.Empty;
                return text.IndexOf("CỘNG HÒA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("CỘNG HOÀ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("ĐẢNG CỘNG SẢN", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (text.IndexOf("Độc lập", StringComparison.OrdinalIgnoreCase) >= 0 &&
                     text.IndexOf("Hạnh phúc", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch (COMException)
            {
                return false;
            }
            finally { Release(range); }
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
            IReadOnlyList<LogicalDocumentBlock> blocks)
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
                string? replacement = null;
                if (group.Any(item => item.RuleCode == "ND30-PL1-M2-K1-TN-SEP"))
                {
                    replacement = LocalAdministrativeTextNormalizer.NationalMotto;
                }
                else
                {
                    var abbreviation = ResolveExpectedTypeAbbreviation(snapshot,
                        paragraph.Index, rules, blocks);
                    replacement = LocalAdministrativeTextNormalizer.NormalizeCodeNumber(
                        printable, IsParty(snapshot), abbreviation);
                }
                if (!string.Equals(printable, replacement, StringComparison.Ordinal))
                    edits.Add(new SpellingEdit(paragraph.AbsoluteStart, printable.Length,
                        replacement ?? printable, printable, 100));
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

            var applied = new List<SpellingEdit>();
            var count = 0;
            foreach (var edit in edits.OrderByDescending(item => item.Start).ThenByDescending(item => item.Priority))
            {
                if (applied.Any(item => Intersects(edit, item))) continue;
                Word.Range? range = null;
                try
                {
                    if (!string.IsNullOrEmpty(edit.ExpectedText))
                    {
                        var validationEnd = checked(edit.Start + edit.ExpectedText.Length);
                        if (validationEnd > document.Content.End) continue;
                        Word.Range? validation = null;
                        try
                        {
                            validation = document.Range(edit.Start, validationEnd);
                            if (!string.Equals(validation.Text ?? string.Empty, edit.ExpectedText,
                                    StringComparison.Ordinal))
                                continue;
                        }
                        finally { Release(validation); }
                    }
                    range = document.Range(edit.Start, checked(edit.Start + edit.Length));
                    if (range.Editors.Count > 0 || range.Fields.Count > 0 ||
                        range.ContentControls.Count > 0 || range.InlineShapes.Count > 0)
                        continue;
                    if (edit.Length == 0) range.InsertBefore(edit.Replacement);
                    else range.Text = edit.Replacement;
                    applied.Add(edit);
                    count++;
                }
                finally { Release(range); }
            }
            return count;
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

        private static Word.Range ResolveCurrentMainParagraphRange(
            Word.Document document,
            LocalParagraphSnapshot paragraph)
        {
            if (!string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) ||
                paragraph.Index <= 0 || paragraph.Index > document.Paragraphs.Count)
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

        private static int CollapseMultipleSpaces(Word.Document document)
        {
            var collapsed = 0;
            try
            {
                for (var iteration = 0; iteration < 3; iteration++)
                {
                    Word.Range? content = null;
                    Word.Find? find = null;
                    Word.Replacement? replacement = null;
                    try
                    {
                        content = document.Content;
                        find = content.Find;
                        find.ClearFormatting();
                        replacement = find.Replacement;
                        replacement.ClearFormatting();

                        find.Text = "  ";
                        replacement.Text = " ";
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.Format = false;
                        find.MatchCase = false;
                        find.MatchWholeWord = false;
                        find.MatchWildcards = false;

                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        object missing = Type.Missing;
                        var found = find.Execute(ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                        if (found)
                        {
                            collapsed++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    catch (COMException)
                    {
                        break;
                    }
                    finally
                    {
                        Release(replacement);
                        Release(find);
                        Release(content);
                    }
                }
            }
            catch (COMException) { }
            return collapsed;
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

        private static bool ContainsRecognizedRole(Word.Table table, LocalScanSnapshot snapshot,
            IReadOnlyDictionary<int, string> roles)
        {
            Word.Range? range = null;
            try
            {
                range = table.Range;
                return snapshot.Paragraphs.Any(paragraph => roles.ContainsKey(paragraph.Index) &&
                    paragraph.AbsoluteStart >= range.Start && paragraph.AbsoluteStart < range.End);
            }
            finally { Release(range); }
        }

        private static void RemoveTrailingBlankParagraphs(Word.Document document)
        {
            WordTrailingBlankPageCleaner.Remove(document);
        }

        private static bool IsBodyParagraph(LocalParagraphSnapshot paragraph)
        {
            var text = paragraph.Text.Trim();
            if (text.Length < 20 || text.Length > 5000) return false;
            if (text.Where(char.IsLetter).Any() && text.Where(char.IsLetter).All(char.IsUpper)) return false;
            if (Regex.IsMatch(text, @"^(Điều\s+\d+|\d+\.\s|[a-zđ]\)\s)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) return false;
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
