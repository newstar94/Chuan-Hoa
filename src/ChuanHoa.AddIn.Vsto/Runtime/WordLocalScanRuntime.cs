using System;
using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Scanning;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class WordLocalScanRuntime
    {
        private readonly Word.Application _application;
        private readonly WordDocumentCapabilityProvider _capabilityProvider;
        private readonly LocalAccessManager _accessManager;
        private readonly AnnotationPlanner _planner = new AnnotationPlanner();

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
            var plan = _planner.CreatePlan(result.Lane, result.ScanId, result.DocumentFingerprint, result.Revision,
                ToAnnotationSnapshot(wordSnapshot), result.Findings);
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
                    ToAnnotationSnapshot(wordSnapshot), fallbackFindings);
            }

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
                    item.TextWidthPoints, item.KeepWithNext, item.WidowControl, item.StyleName)).ToArray(),
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
                    t.IsNested, t.HeaderRowIndexes, t.HasVerticalBorders)).ToArray());

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
