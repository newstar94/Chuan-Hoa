using System;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using Xunit;

namespace ChuanHoa.Client.Core.Tests.Scanning
{
    public sealed class LatexTypographicScannerTests
    {
        private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");

        private static LocalRulePack Rules() => new("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
            210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
            Array.Empty<TextCorrectionRule>(), Array.Empty<TelexRule>(), Array.Empty<char>());

        private static LocalSectionSnapshot ValidSection()
        {
            const double pt = 72d / 25.4d;
            return new LocalSectionSnapshot(1, 210 * pt, 297 * pt, 20 * pt, 20 * pt, 30 * pt, 15 * pt, false);
        }

        private static LocalParagraphSnapshot P(int index, string text, bool bold = true, int? outlineLevel = null,
            bool? keepWithNext = null, bool? widowControl = null, string? styleName = null,
            bool inTable = false, int? tableIndex = null) =>
            new LocalParagraphSnapshot(index, text, "wdMainTextStory", 1, index * 50, "Times New Roman",
                bold: bold, outlineLevel: outlineLevel, keepWithNext: keepWithNext, widowControl: widowControl,
                styleName: styleName, isInTable: inTable, tableIndex: tableIndex);

        [Fact]
        public void Flags_hand_typed_heading_missing_heading_style()
        {
            var paragraphs = new[]
            {
                P(1, "1.1. Mục tiêu nghiên cứu", bold: true, outlineLevel: 10, styleName: "Normal")
            };

            var snapshot = new LocalScanSnapshot("sha256:h-style", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>());

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Contains(findings, f => f.RuleCode == "LATEX-SEC-STYLE" && f.Anchor.ParagraphIndex == 1);
        }

        [Fact]
        public void Flags_skipped_heading_numbers()
        {
            var paragraphs = new[]
            {
                P(1, "1.1. Mục tiêu A", bold: true),
                P(2, "1.3. Mục tiêu B", bold: true) // Skipped 1.2
            };

            var snapshot = new LocalScanSnapshot("sha256:h-cont", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>());

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Contains(findings, f => f.RuleCode == "LATEX-SEC-CONTINUITY" && f.Anchor.ParagraphIndex == 2);
        }

        [Fact]
        public void Flags_heading_missing_keep_with_next()
        {
            var paragraphs = new[]
            {
                P(1, "1. ĐẶT VẤN ĐỀ", bold: true, keepWithNext: false)
            };

            var snapshot = new LocalScanSnapshot("sha256:h-kwn", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>());

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Contains(findings, f => f.RuleCode == "LATEX-PAGINATION-KEEP" && f.Anchor.ParagraphIndex == 1);
        }

        [Fact]
        public void Flags_body_paragraph_missing_widow_control()
        {
            var longText = "Đây là một đoạn văn bản nghiên cứu tương đối dài được viết trong tài liệu để kiểm tra tính năng phát hiện đoạn văn chưa bật thuộc tính kiểm soát ngắt dòng mồ côi widow and orphan control của add-in.";
            var paragraphs = new[]
            {
                P(1, longText, bold: false, widowControl: false)
            };

            var snapshot = new LocalScanSnapshot("sha256:p-widow", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>());

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Contains(findings, f => f.RuleCode == "LATEX-PAGINATION-WIDOW" && f.Anchor.ParagraphIndex == 1);
        }

        [Fact]
        public void Flags_table_with_vertical_borders()
        {
            var paragraphs = new[]
            {
                P(1, "Nội dung ô 1", inTable: true, tableIndex: 1),
                P(2, "Nội dung ô 2", inTable: true, tableIndex: 1)
            };
            var tables = new[]
            {
                new LocalTableSnapshot(1, 2, 2, hasVerticalBorders: true)
            };

            var snapshot = new LocalScanSnapshot("sha256:t-borders", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>(), tables: tables);

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Contains(findings, f => f.RuleCode == "LATEX-TABLE-BOOKTABS");
        }

        [Fact]
        public void Flags_table_caption_placed_below_table()
        {
            var paragraphs = new[]
            {
                P(1, "Dữ liệu bảng ô 1", inTable: true, tableIndex: 1),
                P(2, "Bảng 1: Thống kê số liệu chi tiết", inTable: false) // Placed below table!
            };

            var snapshot = new LocalScanSnapshot("sha256:t-caption", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>());

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Contains(findings, f => f.RuleCode == "LATEX-CAPTION-POS" && f.Anchor.ParagraphIndex == 2);
        }

        [Fact]
        public void Flags_unrendered_latex_math_syntax()
        {
            var paragraphs = new[]
            {
                P(1, "Ta có phương trình năng lượng $E = mc^2$ theo thuyết tương đối hẹp."),
                P(2, "Tích phân toàn phần biểu diễn dưới dạng: $$\\int_{0}^{1} f(x)dx$$.")
            };

            var snapshot = new LocalScanSnapshot("sha256:math", 1, new[] { ValidSection() }, paragraphs,
                Array.Empty<AnnotationProtectedSpan>());

            var scanner = new LatexTypographicScanner();
            var findings = scanner.Scan(snapshot, Rules());

            Assert.Equal(2, findings.Count(f => f.RuleCode == "LATEX-MATH-SYNTAX"));
        }
    }
}
