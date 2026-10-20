using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Scanning;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>What the last annotation pass did (D2 comment budget).</summary>
    public sealed class LocalAnnotationSummary
    {
        public LocalAnnotationSummary(string lane, int totalFindings, int commentCount,
            int annotatedFindings, int omittedFindings, string reportPath)
        {
            Lane = lane;
            TotalFindings = totalFindings;
            CommentCount = commentCount;
            AnnotatedFindings = annotatedFindings;
            OmittedFindings = omittedFindings;
            ReportPath = reportPath ?? string.Empty;
        }

        public string Lane { get; }
        public int TotalFindings { get; }
        public int CommentCount { get; }
        public int AnnotatedFindings { get; }
        public int OmittedFindings { get; }
        public string ReportPath { get; }

        public string Describe()
        {
            if (TotalFindings == 0) return "Không phát hiện lỗi.";
            var text = "Đã đánh dấu " + AnnotatedFindings.ToString(CultureInfo.CurrentCulture) + "/" +
                TotalFindings.ToString(CultureInfo.CurrentCulture) + " điểm bằng " +
                CommentCount.ToString(CultureInfo.CurrentCulture) + " comment.";
            if (OmittedFindings > 0)
                text += "\n" + OmittedFindings.ToString(CultureInfo.CurrentCulture) +
                    " điểm còn lại chỉ có trong báo cáo đầy đủ" +
                    (ReportPath.Length > 0 ? ":\n" + ReportPath : ".");
            return text;
        }
    }

    public sealed class WordLocalScanRuntime
    {
        private readonly Word.Application _application;
        private readonly WordDocumentCapabilityProvider _capabilityProvider;
        private readonly LocalAccessManager _accessManager;
        private readonly AnnotationPlanner _planner = new AnnotationPlanner();

        /// <summary>Summary of the most recent <see cref="ScanAndAnnotate"/> call.</summary>
        public LocalAnnotationSummary? LastAnnotationSummary { get; private set; }

        public WordLocalScanRuntime(Word.Application application, LocalAccessManager accessManager)
        {
            _application = application;
            _capabilityProvider = new WordDocumentCapabilityProvider(application);
            _accessManager = accessManager;
        }

        public LocalScanResult ScanAndAnnotate(
            DocumentContext context,
            bool spelling,
            Word.Document? activeDocument = null,
            DocumentOperationSession? operation = null)
        {
            var document = activeDocument ?? _application.ActiveDocument;
            var capability = _capabilityProvider.Evaluate(document);
            if (!capability.CanReadDocument) throw new InvalidOperationException(capability.Reason);
            if (capability.IsReadOnly || capability.IsProtected || capability.TrackChangesEnabled)
                throw new InvalidOperationException("Để comment và tô đỏ, tài liệu phải cho phép chỉnh sửa, không bảo vệ và tắt Track Changes.");
            if (spelling) context.RequireSpellingAnalysis();
            else context.RequireFormatAnalysis();
            var wordSnapshot = context.LastSnapshot!;
            var result = spelling ? context.LastSpellingScan! : context.LastFormatScan!;
            var rules = _accessManager.GetRulePack(spelling ? LocalAccessManager.SpellingFeature : LocalAccessManager.FormatFeature);
            if (!string.Equals(result.RulePackId, rules.PackId, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Gói quy tắc đã thay đổi. Hãy bấm lại chức năng kiểm tra.");
            // D2: at most 200 comments per lane, one per paragraph, and one summary
            // comment for a rule repeated in more than 20 paragraphs. Everything
            // else stays in the scan result and in the local HTML report.
            var options = AnnotationPlanOptions.Default;
            var plan = _planner.CreatePlan(result.Lane, result.ScanId, result.DocumentFingerprint, result.Revision,
                ToAnnotationSnapshot(wordSnapshot), result.Findings, options);
            if (plan.Unresolved.Count > 0)
            {
                var unresolvedIds = new HashSet<string>(plan.Unresolved.Select(u => u.FindingId), StringComparer.Ordinal);
                var fallbackFindings = result.Findings.Select(f =>
                {
                    if (unresolvedIds.Contains(f.FindingId) &&
                        f.Anchor.Kind == AnnotationAnchorKind.TextSpan &&
                        f.Anchor.ParagraphIndex.HasValue)
                    {
                        return new AnnotationFinding(f.FindingId, f.RuleCode, f.Severity, f.CurrentIssue, f.Expected, f.Citation,
                            new AnnotationAnchor(AnnotationAnchorKind.Paragraph, f.Anchor.StoryType, f.Anchor.ParagraphIndex,
                                null, null, string.Empty, f.Anchor.SectionIndex, f.Anchor.TableIndex, f.Anchor.RowIndex, f.Anchor.CellIndex));
                    }
                    return f;
                }).ToArray();
                plan = _planner.CreatePlan(result.Lane, result.ScanId, result.DocumentFingerprint, result.Revision,
                    ToAnnotationSnapshot(wordSnapshot), fallbackFindings, options);
            }

            var reportPath = plan.Omitted.Count > 0 || plan.Unresolved.Count > 0
                ? TryWriteReport(document, result, plan)
                : string.Empty;
            LastAnnotationSummary = new LocalAnnotationSummary(result.Lane, result.Findings.Count,
                plan.Comments.Count, plan.AnnotatedFindingCount, plan.Omitted.Count, reportPath);

            if (plan.Comments.Count > 0 || plan.VisualRanges.Count > 0)
            {
                operation?.Transition(DocumentOperationState.Annotating,
                    "comment và tô đỏ " + result.Findings.Count + " lỗi");
                new WordFindingAnnotationAdapter(_application, document).Apply(plan, operation);
            }
            else if (plan.Unresolved.Count > 0 && result.Findings.Count > 0)
            {
                throw new InvalidOperationException("Không thể neo chính xác " + plan.Unresolved.Count + " lỗi; tài liệu không bị đánh dấu.");
            }

            return result;
        }

        private static string TryWriteReport(Word.Document document, LocalScanResult result,
            AnnotationPlan plan)
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ChuanHoa", "Reports");
                Directory.CreateDirectory(directory);
                string name;
                try { name = document.Name ?? string.Empty; }
                catch (System.Runtime.InteropServices.COMException) { name = string.Empty; }
                var safeName = new string((Path.GetFileNameWithoutExtension(name) ?? "TaiLieu")
                    .Select(character => Path.GetInvalidFileNameChars().Contains(character) ? '_' : character)
                    .ToArray());
                if (safeName.Length == 0) safeName = "TaiLieu";
                if (safeName.Length > 60) safeName = safeName.Substring(0, 60);
                var now = DateTime.Now;
                var path = Path.Combine(directory, safeName + "_" + result.Lane + "_" +
                    now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".html");
                File.WriteAllText(path, AnnotationReportBuilder.BuildHtml(name, result.Lane, now,
                    result.Findings, plan), new UTF8Encoding(true));
                return path;
            }
            catch (IOException) { return string.Empty; }
            catch (UnauthorizedAccessException) { return string.Empty; }
        }

        internal static LocalScanSnapshot ToLocalSnapshot(WordDocumentSnapshot source, DocumentContext? context = null) =>
            new LocalScanSnapshot(source.DocumentFingerprint, source.Revision,
                source.Sections.Select(item => new LocalSectionSnapshot(item.Index, item.PageWidthPoints,
                    item.PageHeightPoints, item.TopMarginPoints, item.BottomMarginPoints, item.LeftMarginPoints,
                    item.RightMarginPoints, item.IsLandscape, item.HasPageNumbers, item.RestartPageNumbering,
                    item.StartingPageNumber, item.PageNumberAlignment)).ToArray(),
                source.Paragraphs.Select(item => new LocalParagraphSnapshot(item.Index, item.Text, item.StoryType,
                    item.SectionIndex, item.AbsoluteStart, item.FontName ?? string.Empty, item.TableIndex, item.RowIndex, item.CellIndex,
                    item.FontSizePoints, item.Bold, item.Italic, item.Alignment, item.FirstLineIndentPoints,
                    item.SpaceBeforePoints, item.SpaceAfterPoints, item.IsInTable, item.Role, item.FontColor,
                    item.Underline, item.HasBottomBorder, item.LineSpacingPoints, item.LineSpacingRule,
                    item.OutlineLevel, item.PageNumber, item.PageLeftPoints, item.PageTopPoints,
                    item.TextWidthPoints, item.KeepWithNext, item.WidowControl, item.StyleName,
                    item.AbsoluteEnd, item.BuiltInStyleId, item.HasField, item.HasMathObject,
                    item.HasHyperlink, item.HasContentControl, item.CaptionKind,
                    item.TableNestingDepth, item.LeftIndentPoints, item.ListMarker)).ToArray(),
                source.ProtectedSpans.Select(item => new AnnotationProtectedSpan(item.StoryType, item.AbsoluteStart, item.Length)).ToArray(),
                source.LineShapes.Select(item => new LocalLineShapeSnapshot(item.Index, item.Name, item.ShapeType,
                    item.AnchorStoryType, item.AnchorSectionIndex, item.AnchorAbsoluteStart,
                    item.AnchorParagraphIndex, item.AnchorPageNumber, item.LeftPoints, item.TopPoints,
                    item.WidthPoints, item.HeightPoints, item.PageLeftPoints, item.PageTopPoints,
                    item.RelativeHorizontalPosition, item.RelativeVerticalPosition, item.LineVisible,
                    item.DashStyle, item.WeightPoints, item.Color, item.BeginArrowheadStyle,
                    item.EndArrowheadStyle)).ToArray(), context == null ? source.RegimeCode : context.RegimeCode,
                context == null ? source.DocumentTypeCode : context.DocumentTypeCode,
                context == null ? source.RegimeWasSelectedManually : context.RegimeWasSelectedManually,
                context == null ? source.DocumentTypeWasSelectedManually : context.DocumentTypeWasSelectedManually,
                context == null ? source.DocumentFingerprint : context.DictionaryScopeId,
                source.Tables.Select(t => new LocalTableSnapshot(t.Index, t.RowCount, t.ColumnCount, t.HasMergedCells,
                    t.IsNested, t.HeaderRowIndexes, t.HasVerticalBorders,
                    t.HeaderSeparatorBorder.State == LocalSnapshotValueState.Present,
                    storyType: t.StoryType, sectionIndex: t.SectionIndex,
                    absoluteStart: t.AbsoluteStart, absoluteEnd: t.AbsoluteEnd,
                    nestingDepth: t.NestingDepth, topBorder: t.TopBorder,
                    bottomBorder: t.BottomBorder, leftBorder: t.LeftBorder,
                    rightBorder: t.RightBorder, insideHorizontalBorder: t.InsideHorizontalBorder,
                    insideVerticalBorder: t.InsideVerticalBorder,
                    headerSeparatorBorder: t.HeaderSeparatorBorder,
                    hasMergedCellsState: t.HasMergedCellsState)).ToArray(),
                source.GraphicObjects.Select(g => new LocalGraphicObjectSnapshot(g.Index, g.ObjectKind,
                    g.IsInline, g.StoryType, g.SectionIndex, g.AbsoluteStart, g.AbsoluteEnd,
                    g.AnchorParagraphIndex, g.IsProtected)).ToArray(), source.SchemaVersion);

        internal static AnnotationDocumentSnapshot ToAnnotationSnapshot(WordDocumentSnapshot source) =>
            new AnnotationDocumentSnapshot(source.DocumentFingerprint, source.Revision,
                source.Paragraphs.Select(item => new AnnotationParagraphSnapshot(item.StoryType, item.Index,
                    item.SectionIndex, item.AbsoluteStart, item.Text, item.TableIndex, item.RowIndex, item.CellIndex)).ToArray(),
                source.ProtectedSpans.Select(item => new AnnotationProtectedSpan(item.StoryType, item.AbsoluteStart, item.Length)).ToArray());
    }

    internal static class WordDocumentTypeClassifier
    {
        internal static string DetectAndApply(DocumentContext context, LocalScanSnapshot snapshot,
            bool requireKnownType)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

            var detector = new DocumentRoleDetector();
            var detectedFromContent = detector.ResolveDocumentTypeFromContent(snapshot);
            var resolved = string.Equals(detectedFromContent, LocalDocumentTypeCodes.Unknown,
                    StringComparison.OrdinalIgnoreCase)
                ? detector.ResolveDocumentType(snapshot)
                : detectedFromContent;

            if (string.Equals(resolved, LocalDocumentTypeCodes.Unknown, StringComparison.OrdinalIgnoreCase))
            {
                context.DocumentTypeCode = LocalDocumentTypeCodes.Unknown;
                context.DocumentTypeWasSelectedManually = false;
                if (requireKnownType)
                    throw new InvalidOperationException(
                        "Ứng dụng chưa tự xác định được loại văn bản. Hãy bảo đảm tài liệu có dòng tên loại như QUYẾT ĐỊNH, THÔNG BÁO, BÁO CÁO… hoặc trích yếu Công văn bắt đầu bằng V/v hoặc Về việc.");
                return LocalDocumentTypeCodes.Unknown;
            }

            context.DocumentTypeCode = resolved;
            if (!string.Equals(detectedFromContent, LocalDocumentTypeCodes.Unknown,
                    StringComparison.OrdinalIgnoreCase))
                context.DocumentTypeWasSelectedManually = false;
            return resolved;
        }
    }
}
