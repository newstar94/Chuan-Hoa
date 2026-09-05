using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Scans document snapshots for academic / LaTeX / Typst publishing standards,
    /// complementing administrative decree rules without violating legal mandates.
    /// </summary>
    public sealed class LatexTypographicScanner
    {
        private static readonly Regex InlineMathRegex = new Regex(
            @"(?<!\\)\$(?!\$)(?<math>[^$\r\n]+?)\$(?!\$)",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex DisplayMathRegex = new Regex(
            @"\$\$(?<math>[^$\r\n]+?)\$\$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex TableCaptionRegex = new Regex(
            @"^\s*Bảng\s+\d+[\.\-\/:]*\s*",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex FigureCaptionRegex = new Regex(
            @"^\s*Hình\s+\d+[\.\-\/:]*\s*",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private readonly HeadingDetector _headingDetector = new HeadingDetector();

        public IReadOnlyList<AnnotationFinding> Scan(
            LocalScanSnapshot snapshot,
            LocalRulePack rules,
            CancellationToken cancellationToken = default)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (rules == null) throw new ArgumentNullException(nameof(rules));

            var findings = new List<AnnotationFinding>();
            var paragraphs = snapshot.Paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.Text)).ToList();

            // 1. Heading Recognition & Style Checks
            cancellationToken.ThrowIfCancellationRequested();
            var detectedHeadings = _headingDetector.Detect(paragraphs);
            CheckHeadingStyles(findings, detectedHeadings, paragraphs);
            CheckHeadingContinuity(findings, detectedHeadings);
            CheckKeepWithNext(findings, detectedHeadings);

            // 2. Pagination / Widow & Orphan Control
            cancellationToken.ThrowIfCancellationRequested();
            CheckWidowOrphanControl(findings, paragraphs, detectedHeadings);

            // 3. LaTeX Math Syntax
            cancellationToken.ThrowIfCancellationRequested();
            CheckMathSyntax(findings, paragraphs);

            // 4. Table Booktabs Styling & Captions
            cancellationToken.ThrowIfCancellationRequested();
            CheckTableBooktabs(findings, snapshot);
            CheckCaptions(findings, paragraphs);

            return findings;
        }

        private static void CheckHeadingStyles(
            ICollection<AnnotationFinding> findings,
            IReadOnlyList<DetectedHeading> headings,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs)
        {
            var pMap = paragraphs.ToDictionary(p => p.Index);

            foreach (var h in headings)
            {
                if (!pMap.TryGetValue(h.ParagraphIndex, out var p)) continue;

                // If not assigned to a Heading style or outline level is body text (>= 10)
                var hasHeadingStyle = p.StyleName != null &&
                    p.StyleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
                var hasValidOutline = p.OutlineLevel.HasValue && p.OutlineLevel.Value < 10;

                if (!hasHeadingStyle && !hasValidOutline)
                {
                    findings.Add(new AnnotationFinding(
                        "LATEX-SEC-STYLE-P" + p.Index,
                        "LATEX-SEC-STYLE",
                        "Warning",
                        "Đề mục cấp " + h.Level + " ('" + h.NumberText + " " + h.TitleText + "') chưa được gán Style Heading.",
                        "Gán Style Heading " + h.Level + " để kích hoạt Mục lục tự động và Bookmarks PDF.",
                        "Chuẩn xuất bản học thuật (LaTeX/Typst)",
                        Anchor(AnnotationAnchorKind.Paragraph, p, null, null, string.Empty)));
                }
            }
        }

        private void CheckHeadingContinuity(
            ICollection<AnnotationFinding> findings,
            IReadOnlyList<DetectedHeading> headings)
        {
            var issues = _headingDetector.AnalyzeContinuity(headings);
            foreach (var issue in issues)
            {
                var h = issue.Heading;
                findings.Add(new AnnotationFinding(
                    "LATEX-SEC-CONTINUITY-P" + h.ParagraphIndex,
                    "LATEX-SEC-CONTINUITY",
                    "Warning",
                    issue.CurrentIssue,
                    issue.Expected,
                    "Chuẩn đánh số thứ tự đề mục (LaTeX/Typst)",
                    new AnnotationAnchor(
                        AnnotationAnchorKind.Paragraph,
                        "wdMainTextStory",
                        h.ParagraphIndex,
                        null, null, string.Empty,
                        1, null, null, null)));
            }
        }

        private static void CheckKeepWithNext(
            ICollection<AnnotationFinding> findings,
            IReadOnlyList<DetectedHeading> headings)
        {
            foreach (var h in headings)
            {
                if (h.KeepWithNext == false)
                {
                    findings.Add(new AnnotationFinding(
                        "LATEX-PAGINATION-KEEP-P" + h.ParagraphIndex,
                        "LATEX-PAGINATION-KEEP",
                        "Warning",
                        "Tiêu đề mục chưa bật thuộc tính dính liền dòng tiếp theo (Keep with next).",
                        "Bật KeepWithNext để tránh tiêu đề bị ngắt trang rơi xuống đáy trang lẻ loi.",
                        "Chuẩn dàn trang chống mồ côi tiêu đề (LaTeX)",
                        new AnnotationAnchor(
                            AnnotationAnchorKind.Paragraph,
                            "wdMainTextStory",
                            h.ParagraphIndex,
                            null, null, string.Empty,
                            1, null, null, null)));
                }
            }
        }

        private static void CheckWidowOrphanControl(
            ICollection<AnnotationFinding> findings,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs,
            IReadOnlyList<DetectedHeading> headings)
        {
            var headingIndices = new HashSet<int>(headings.Select(h => h.ParagraphIndex));

            foreach (var p in paragraphs)
            {
                if (p.IsInTable) continue;
                if (headingIndices.Contains(p.Index)) continue;
                if (p.Text.Length < 120) continue; // Only for multiline body paragraphs

                if (p.WidowControl == false)
                {
                    findings.Add(new AnnotationFinding(
                        "LATEX-PAGINATION-WIDOW-P" + p.Index,
                        "LATEX-PAGINATION-WIDOW",
                        "Suggestion",
                        "Đoạn văn bản dài chưa bật kiểm soát ngắt dòng mồ côi (Widow/Orphan control).",
                        "Bật Widow/Orphan control để tránh 1 dòng đơn lẻ bị rớt sang trang mới.",
                        "Chuẩn vi mô ngắt dòng đoạn văn (LaTeX/Typst)",
                        Anchor(AnnotationAnchorKind.Paragraph, p, null, null, string.Empty)));
                }
            }
        }

        private static void CheckMathSyntax(
            ICollection<AnnotationFinding> findings,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs)
        {
            foreach (var p in paragraphs)
            {
                var text = p.Text;
                if (string.IsNullOrEmpty(text) || !text.Contains("$")) continue;

                // 1. Display math $$ ... $$
                foreach (Match m in DisplayMathRegex.Matches(text))
                {
                    if (m.Success && m.Length >= 4)
                    {
                        findings.Add(new AnnotationFinding(
                            "LATEX-MATH-SYNTAX-P" + p.Index + "-C" + m.Index,
                            "LATEX-MATH-SYNTAX",
                            "Warning",
                            "Phát hiện công thức toán học khối dạng mã LaTeX: '" + m.Value + "'.",
                            "Chuyển đổi thành công thức Word Equation (OMath) căn giữa chuẩn xuất bản.",
                            "Chuẩn biểu thức toán học (LaTeX Math)",
                            Anchor(AnnotationAnchorKind.TextSpan, p, m.Index, m.Length, m.Value)));
                    }
                }

                // 2. Inline math $ ... $
                foreach (Match m in InlineMathRegex.Matches(text))
                {
                    if (m.Success && m.Length >= 3)
                    {
                        findings.Add(new AnnotationFinding(
                            "LATEX-MATH-SYNTAX-P" + p.Index + "-C" + m.Index,
                            "LATEX-MATH-SYNTAX",
                            "Warning",
                            "Phát hiện ký hiệu toán học inline dạng mã LaTeX: '" + m.Value + "'.",
                            "Chuyển đổi thành công thức Word Equation (OMath) nội dòng sắc nét.",
                            "Chuẩn biểu thức toán học (LaTeX Math)",
                            Anchor(AnnotationAnchorKind.TextSpan, p, m.Index, m.Length, m.Value)));
                    }
                }
            }
        }

        private static void CheckTableBooktabs(
            ICollection<AnnotationFinding> findings,
            LocalScanSnapshot snapshot)
        {
            if (snapshot.Tables == null || snapshot.Tables.Count == 0) return;

            foreach (var t in snapshot.Tables)
            {
                if (t.HasVerticalBorders)
                {
                    // Find first paragraph in table to anchor the finding
                    var tableP = snapshot.Paragraphs.FirstOrDefault(p => p.TableIndex == t.Index);
                    if (tableP != null)
                    {
                        findings.Add(new AnnotationFinding(
                            "LATEX-TABLE-BOOKTABS-T" + t.Index,
                            "LATEX-TABLE-BOOKTABS",
                            "Suggestion",
                            "Bảng số " + t.Index + " đang có viền dọc ô lưới thô kệch.",
                            "Áp dụng chuẩn booktabs (xóa viền dọc, giữ 3 đường ngang thanh lịch) để bảng thoáng đãng, chuyên nghiệp.",
                            "Chuẩn bảng biểu học thuật (LaTeX booktabs)",
                            Anchor(AnnotationAnchorKind.Paragraph, tableP, null, null, string.Empty)));
                    }
                }
            }
        }

        private static void CheckCaptions(
            ICollection<AnnotationFinding> findings,
            IReadOnlyList<LocalParagraphSnapshot> paragraphs)
        {
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var p = paragraphs[i];
                var text = p.Text.Trim();

                // Check table caption placed below table:
                // If previous paragraph is in table, but current paragraph is NOT in table and starts with "Bảng \d+"
                if (!p.IsInTable && TableCaptionRegex.IsMatch(text))
                {
                    if (i > 0 && paragraphs[i - 1].IsInTable)
                    {
                        findings.Add(new AnnotationFinding(
                            "LATEX-CAPTION-POS-P" + p.Index,
                            "LATEX-CAPTION-POS",
                            "Warning",
                            "Chú thích '" + text.Substring(0, Math.Min(text.Length, 40)) + "...' đang đặt ở PHÍA DƯỚI bảng.",
                            "Đặt chú thích Bảng ở PHÍA TRÊN bảng theo quy chuẩn xuất bản học thuật và NĐ30.",
                            "Chuẩn vị trí chú thích bảng biểu (LaTeX/NĐ30)",
                            Anchor(AnnotationAnchorKind.Paragraph, p, null, null, string.Empty)));
                    }
                }
            }
        }

        private static AnnotationAnchor Anchor(
            AnnotationAnchorKind kind,
            LocalParagraphSnapshot paragraph,
            int? offset,
            int? length,
            string expectedText) =>
            new AnnotationAnchor(
                kind,
                paragraph.StoryType,
                paragraph.Index,
                offset,
                length,
                expectedText,
                paragraph.SectionIndex,
                paragraph.TableIndex,
                paragraph.RowIndex,
                paragraph.CellIndex);
    }
}
