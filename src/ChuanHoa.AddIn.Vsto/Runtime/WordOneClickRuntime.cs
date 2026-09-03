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
        private readonly DocumentRoleDetector _roleDetector = new DocumentRoleDetector();

        public WordOneClickRuntime(Word.Application application, LocalAccessManager accessManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _capabilityProvider = new WordDocumentCapabilityProvider(application);
            _accessManager = accessManager ?? throw new ArgumentNullException(nameof(accessManager));
        }

        public OneClickResult Execute(DocumentContext context, Word.Document? activeDocument = null)
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
            var roles = _roleDetector.Detect(local);
            var headerFontTier = Nd30HeaderFontSizeTierResolver.Resolve(local, roles);
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
                    formatFindings.Concat(spellingFindings).ToArray(), rules);
                normalizedSections = NormalizeSections(document, local, rules);
                NormalizeHeaderLayoutTables(document, local, roles);
                foreach (var paragraph in local.Paragraphs.Where(p =>
                    string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal)))
                {
                    string role;
                    roles.TryGetValue(paragraph.Index, out role);
                    if (ApplyParagraphFormat(document, paragraph, role ?? string.Empty, local, rules,
                            headerFontTier))
                        changedParagraphs++;
                }

                insertedLines = InsertMissingRequiredLines(document, local, formatFindings);
                EnsurePageNumbers(document, local, rules);
                normalizedTables = NormalizeTables(document, local, roles);
                WordAppendixPaginationNormalizer.Normalize(document, roles);
                RemoveTrailingBlankParagraphs(document);

                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                    undoStarted = false;
                }

                // Never rebuild or rescan after the mutation. The next explicit command
                // prepares a fresh, command-scoped analysis when it needs one.
                context.ClearReadAnalysis();
                var remaining = formatFindings.Concat(spellingFindings).ToArray();
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
            Word.Document? activeDocument = null)
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
                _application.ScreenUpdating = false;
                if (WordMajorVersion() >= 15)
                {
                    _application.UndoRecord.StartCustomRecord("Chuẩn hóa: sửa lỗi đang chọn");
                    undoStarted = true;
                }

                var applied = string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase)
                    ? ApplyDeterministicSpellingFixes(document, local, new[] { finding }, rules) > 0
                    : ApplySelectedFormatFix(document, local, finding, rules);

                if (!applied && string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase))
                {
                    Word.Range? directRange = null;
                    try
                    {
                        directRange = document.Range(selectedStart, selectedEnd);
                        var directText = directRange.Text ?? string.Empty;
                        var target = ResolveTargetReplacement(finding, rules);
                        var expected = finding.Anchor.ExpectedText ?? string.Empty;
                        if (!string.IsNullOrEmpty(target) &&
                            (string.Equals(directText, expected, StringComparison.Ordinal) ||
                             string.Equals(directText.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase)))
                        {
                            directRange.Text = target;
                            applied = true;
                        }
                    }
                    catch (COMException) { }
                    finally { Release(directRange); }
                }

                if (!applied)
                    throw new InvalidOperationException(
                        "Lỗi đang chọn cần người dùng quyết định hoặc chưa có phương án tự sửa an toàn. " +
                        "Comment được giữ nguyên.");

                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                    undoStarted = false;
                }

                annotations.ClearOwnedAnnotationsAt(lane, selectedStory, selectedStart, selectedEnd);
                UpdateContextAfterSingleFix(context, lane, findingId);
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
            AnnotationFinding finding, LocalRulePack rules)
        {
            // This is classified as a format finding, but its safe correction is a
            // deterministic text insertion. Reuse the same guarded edit path as
            // 1-Click so the comment is removed only after "số" was actually added.
            if (string.Equals(finding.RuleCode, "ND30-PL1-M2-K6B-SO", StringComparison.Ordinal))
                return ApplyDeterministicSpellingFixes(document, snapshot,
                    new[] { finding }, rules) > 0;
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

            var roles = new DocumentRoleDetector().Detect(snapshot);
            var headerFontTier = Nd30HeaderFontSizeTierResolver.Resolve(snapshot, roles);
            string role;
            roles.TryGetValue(paragraph.Index, out role);
            return ApplyParagraphFormat(document, paragraph, role ?? string.Empty, snapshot, rules,
                headerFontTier);
        }

        private static int NormalizeSections(Word.Document document, LocalScanSnapshot snapshot, LocalRulePack rules)
        {
            var party = IsParty(snapshot);
            var count = 0;
            foreach (Word.Section section in document.Sections)
            {
                try
                {
                    section.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                    section.PageSetup.TopMargin = (float)((party ? 20d : rules.TopMinMm) * PointsPerMillimeter);
                    section.PageSetup.BottomMargin = (float)((party ? 20d : rules.BottomMinMm) * PointsPerMillimeter);
                    section.PageSetup.LeftMargin = (float)((party ? 30d : rules.LeftMinMm) * PointsPerMillimeter);
                    section.PageSetup.RightMargin = (float)((party ? 15d : rules.RightMinMm) * PointsPerMillimeter);
                    count++;
                }
                finally { Release(section); }
            }
            return count;
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
                if (string.IsNullOrEmpty(role))
                {
                    if (paragraph.IsInTable || !IsBodyParagraph(paragraph)) return false;
                    range.Font.Name = rules.BodyFontName;
                    range.Font.Color = Word.WdColor.wdColorAutomatic;
                    range.Font.Size = party ? 14f : ValidOr(paragraph.FontSizePoints, 13, 14, 14);
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    range.ParagraphFormat.FirstLineIndent = 10f * PointsPerMillimeter;
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
            IReadOnlyList<AnnotationFinding> findings)
        {
            var count = 0;
            var work = findings.Where(f => f.RuleCode.EndsWith("-LINE", StringComparison.Ordinal) &&
                    f.Anchor.ParagraphIndex.HasValue)
                .Select(f => Tuple.Create(f.RuleCode, f.Anchor.ParagraphIndex.GetValueOrDefault()))
                .ToList();

            // 1.0.0.43 could leave an OOXML line object that Word exposed through COM
            // but did not paint. Those owned shapes have the legacy CHUANHOA_* marker,
            // so migrate them once even when the geometry-only scanner calls them valid.
            var roles = new DocumentRoleDetector().Detect(snapshot);
            // 1-Click guarantees every required component line, even when an earlier
            // scan considered an existing Shape valid. Word can keep a line object in
            // OOXML/COM while failing to paint it after save or Modern Comments layout
            // reconciliation. Re-normalizing the canonical components on every explicit
            // 1-Click makes the operation self-healing without scanning at activation.
            if (IsParty(snapshot))
            {
                AddRequiredRoleLine(work, snapshot, roles, "partyTitle", "HD05-M1-TITLE-LINE", false);
            }
            else
            {
                AddRequiredRoleLine(work, snapshot, roles, "organName", "ND30-PL1-M2-K2-ORG-LINE", false);
                AddRequiredRoleLine(work, snapshot, roles, "nationalMotto", "ND30-PL1-M2-K1-TN-LINE", false);
                AddRequiredRoleLine(work, snapshot, roles, "subject", "ND30-PL1-M2-K5A-SUBJ-LINE", true);
                AddRequiredRoleLine(work, snapshot, roles, "subjectContinuation", "ND30-PL1-M2-K5A-SUBJ-LINE", true);
            }
            foreach (var line in snapshot.LineShapes.Where(item =>
                item.Name.StartsWith("CHUANHOA_", StringComparison.Ordinal)))
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

        private static void AddRequiredRoleLine(ICollection<Tuple<string, int>> work,
            LocalScanSnapshot snapshot, IDictionary<int, string> roles, string role,
            string ruleCode, bool useLast)
        {
            var matches = snapshot.Paragraphs.Where(paragraph =>
                    string.Equals(paragraph.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    roles.ContainsKey(paragraph.Index) &&
                    string.Equals(roles[paragraph.Index], role, StringComparison.Ordinal))
                .OrderBy(paragraph => paragraph.Index)
                .ToArray();
            if (matches.Length == 0) return;
            var paragraph = useLast ? matches[matches.Length - 1] : matches[0];
            // Subject and subjectContinuation share one required line. Prefer the last
            // detected subject row and replace an earlier candidate for the same rule.
            if (useLast)
            {
                var earlier = work.Where(item => item.Item1 == ruleCode).ToArray();
                foreach (var item in earlier) work.Remove(item);
            }
            if (!work.Any(item => item.Item1 == ruleCode && item.Item2 == paragraph.Index))
                work.Add(Tuple.Create(ruleCode, paragraph.Index));
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
                if (!inTable && TryNormalizeExistingComponentLine(document, snapshot, paragraph, ruleCode,
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

            var selected = matches[0];
            if (selected.Name.StartsWith("CHUANHOA", StringComparison.Ordinal))
            {
                for (var index = document.Shapes.Count; index >= 1; index--)
                {
                    Word.Shape? legacy = null;
                    try
                    {
                        legacy = document.Shapes[index];
                        if (string.Equals(legacy.Name, selected.Name, StringComparison.Ordinal)) legacy.Delete();
                    }
                    catch (COMException) { }
                    finally { Release(legacy); }
                }
                return false;
            }
            var ownedPrefix = OwnedLinePrefix(ruleCode);
            var obsoleteNames = new HashSet<string>(matches.Skip(1).Select(line => line.Name),
                StringComparer.Ordinal);
            for (var index = document.Shapes.Count; index >= 1; index--)
            {
                Word.Shape? candidate = null;
                try
                {
                    candidate = document.Shapes[index];
                    if ((int)candidate.Type != 9) continue;
                    var name = candidate.Name ?? string.Empty;
                    try
                    {
                        if (candidate.Anchor.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            candidate.Delete();
                            continue;
                        }
                    }
                    catch (COMException) { }
                    if (obsoleteNames.Contains(name) ||
                        (name.StartsWith(ownedPrefix, StringComparison.Ordinal) &&
                         !string.Equals(name, selected.Name, StringComparison.Ordinal)))
                    {
                        candidate.Delete();
                        continue;
                    }
                    if (!string.Equals(name, selected.Name, StringComparison.Ordinal)) continue;

                    candidate.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    candidate.RelativeVerticalPosition =
                        Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    candidate.Left = targetLeft;
                    candidate.Top = targetTop;
                    candidate.Width = targetWidth;
                    candidate.Height = 0f;
                    candidate.Name = OwnedLineName(ruleCode, paragraph.Index);
                    NormalizeLineStyle(candidate);
                    candidate.LockAnchor = 0;
                    return true;
                }
                catch (COMException)
                {
                }
                finally { Release(candidate); }
            }
            return false;
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
            var ownedPrefix = OwnedLinePrefix(ruleCode);
            var associatedNames = new HashSet<string>(snapshot.LineShapes
                .Where(line => IsAssociatedLine(line, paragraph))
                .Select(line => line.Name), StringComparer.Ordinal);
            for (var index = document.Shapes.Count; index >= 1; index--)
            {
                Word.Shape? candidate = null;
                try
                {
                    candidate = document.Shapes[index];
                    if ((int)candidate.Type != 9) continue;
                    var name = candidate.Name ?? string.Empty;
                    if (name.StartsWith(ownedPrefix, StringComparison.Ordinal) ||
                        (name.StartsWith("CHUANHOA", StringComparison.Ordinal) && name.Contains("P" + paragraph.Index)) ||
                        associatedNames.Contains(name))
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
            foreach (Word.Section section in document.Sections)
            {
                Word.HeaderFooter? header = null;
                Word.Range? range = null;
                Word.Field? field = null;
                try
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                    RemoveFirstPageNumber(section);
                    if (!IsParty(snapshot)) RemoveFooterPageNumbers(section);
                    if (ContainsPageField(section)) continue;
                    header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    if (!header.Exists) header.Exists = true;
                    range = header.Range.Duplicate;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    field = document.Fields.Add(range, Word.WdFieldType.wdFieldPage, PreserveFormatting: true);
                    field.Result.Font.Name = rules.BodyFontName;
                    field.Result.Font.Size = 14f;
                    field.Result.Font.Bold = 0;
                    field.Result.Font.Italic = 0;
                }
                finally { Release(field); Release(range); Release(header); Release(section); }
            }
        }

        private static int NormalizeTables(Word.Document document, LocalScanSnapshot snapshot,
            IDictionary<int, string> roles)
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
            IDictionary<int, string> roles)
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

            foreach (Word.Table table in document.Tables)
            {
                try
                {
                    if (IsHeaderLayoutTable(table, snapshot, roles, headerRoles))
                    {
                        NormalizeHeaderTable(table);
                    }
                }
                catch (COMException) { }
                finally { Release(table); }
            }
        }

        private static bool IsHeaderLayoutTable(Word.Table table, LocalScanSnapshot snapshot,
            IDictionary<int, string> roles, ISet<string> headerRoles)
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
                var availableWidth = Math.Max(200f, pageWidth - leftMargin - rightMargin);

                table.Rows.LeftIndent = 0f;
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                try { table.Borders.Enable = 0; } catch (COMException) { }
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = availableWidth;

                var col1Width = (float)Math.Round(availableWidth * 0.38d, 1);
                var col2Width = (float)Math.Round(availableWidth - col1Width, 1);

                for (var r = 1; r <= table.Rows.Count; r++)
                {
                    Word.Row? row = null;
                    try
                    {
                        row = table.Rows[r];
                        row.LeftIndent = 0f;
                        row.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                        if (row.Cells.Count == 2)
                        {
                            row.Cells[1].Width = col1Width;
                            row.Cells[2].Width = col2Width;
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
            IReadOnlyList<AnnotationFinding> findings, LocalRulePack rules)
        {
            var paragraphs = snapshot.Paragraphs.ToDictionary(item => item.Index);
            var edits = new List<SpellingEdit>();
            var lexicon = new VietnameseLexiconSpellChecker(rules.Lexicon);
            foreach (var finding in findings)
            {
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
                    return correction == null ? null : new SpellingEdit(start,
                        length, ApplyCase(expectedText, correction.Replacement),
                        expectedText, 30);
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
                case "ND30-PL2-M1":
                    if (expectedText.Length != 1 || !char.IsLower(expectedText[0])) return null;
                    return new SpellingEdit(start, 1,
                        char.ToUpper(expectedText[0], CultureInfo.GetCultureInfo("vi-VN")).ToString(),
                        expectedText, 20);
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
                case "ND30-PL2-M5-K7":
                    return new SpellingEdit(start, length,
                        expectedText.ToLower(CultureInfo.GetCultureInfo("vi-VN")), expectedText, 20);
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
                    @"(?:Viết|Sửa thành|thành)\s+[“""]([^”""]+)[”""]",
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

            return null;
        }

        private static void UpdateContextAfterSingleFix(DocumentContext context, string lane, string findingId)
        {
            try
            {
                if (context.LastSnapshot != null && context.LastLocalSnapshot != null)
                {
                    if (string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase) && context.LastSpellingScan != null)
                    {
                        var scan = context.LastSpellingScan;
                        var remaining = scan.Findings
                            .Where(f => !string.Equals(f.FindingId, findingId, StringComparison.Ordinal)).ToArray();
                        context.SetAnalysis(context.LastSnapshot, context.LastLocalSnapshot, context.LastFormatScan,
                            new LocalScanResult(scan.ScanId, scan.Lane, scan.RulePackId, scan.DocumentFingerprint, scan.Revision, remaining),
                            false);
                        return;
                    }
                    else if (string.Equals(lane, "format", StringComparison.OrdinalIgnoreCase) && context.LastFormatScan != null)
                    {
                        var scan = context.LastFormatScan;
                        var remaining = scan.Findings
                            .Where(f => !string.Equals(f.FindingId, findingId, StringComparison.Ordinal)).ToArray();
                        context.SetAnalysis(context.LastSnapshot, context.LastLocalSnapshot,
                            new LocalScanResult(scan.ScanId, scan.Lane, scan.RulePackId, scan.DocumentFingerprint, scan.Revision, remaining),
                            context.LastSpellingScan, false);
                        return;
                    }
                }

                context.ClearReadAnalysis();
            }
            catch
            {
                context.ClearReadAnalysis();
            }
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

        private static bool Intersects(SpellingEdit left, SpellingEdit right)
        {
            if (left.Length == 0 || right.Length == 0) return left.Start == right.Start;
            return left.Start < right.Start + right.Length && right.Start < left.Start + left.Length;
        }

        private static bool ContainsRecognizedRole(Word.Table table, LocalScanSnapshot snapshot,
            IDictionary<int, string> roles)
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
            if (Regex.IsMatch(text, @"^(Điều\s+\d+|\d+\.\s|[a-zđ]\)\s|[-–—]?\s*(Căn cứ|Xét|Theo đề nghị))\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) return false;
            return true;
        }

        private static string OwnedLineName(string ruleCode, int paragraphIndex)
        {
            // The binary .doc format rejects long Shape.Name values. Keep the
            // ownership marker short and ASCII for Word 2010 compatibility.
            return OwnedLinePrefix(ruleCode) + "P" + paragraphIndex.ToString(CultureInfo.InvariantCulture);
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
