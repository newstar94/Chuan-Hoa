using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Scanning;
using Xunit;

namespace ChuanHoa.Client.Core.Tests.Scanning
{
    public sealed class HeadingDetectorTests
    {
        private readonly HeadingDetector _detector = new HeadingDetector();

        private static LocalParagraphSnapshot P(
            int index,
            string text,
            bool bold = true,
            int? outlineLevel = null,
            bool? keepWithNext = null,
            string? styleName = null,
            bool inTable = false,
            string role = "Unknown",
            string storyType = "wdMainTextStory") =>
            new LocalParagraphSnapshot(
                index,
                text,
                storyType,
                1,
                index * 50,
                "Times New Roman",
                bold: bold,
                outlineLevel: outlineLevel,
                keepWithNext: keepWithNext,
                styleName: styleName,
                isInTable: inTable,
                role: role);

        [Fact]
        public void Detects_confident_decimal_headings_with_supported_separators()
        {
            var paragraphs = new[]
            {
                P(1, "1. TỔNG QUAN NGHIÊN CỨU", bold: false),
                P(2, "1.1. Mục tiêu cụ thể"),
                P(3, "1.1.1- Phương pháp thực hiện"),
                P(4, "Nội dung thân bài mô tả chi tiết phương pháp...", bold: false)
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(3, headings.Count);
            Assert.Equal(1, headings[0].Level);
            Assert.Equal("1", headings[0].NumberText);
            Assert.Equal("TỔNG QUAN NGHIÊN CỨU", headings[0].TitleText);
            Assert.Equal(2, headings[1].Level);
            Assert.Equal("1.1", headings[1].NumberText);
            Assert.Equal(3, headings[2].Level);
            Assert.Equal("1.1.1", headings[2].NumberText);
        }

        [Fact]
        public void Single_level_decimal_requires_metadata_or_an_uppercase_title()
        {
            var paragraphs = new[]
            {
                P(1, "1. Nội dung khoản", bold: true),
                P(2, "2. Tổng quan nghiên cứu", bold: true, outlineLevel: 1),
                P(3, "3. Phương pháp nghiên cứu", bold: false, styleName: "Heading 1")
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(new[] { 2, 3 }, headings.Select(heading => heading.ParagraphIndex));
        }

        [Fact]
        public void Detects_only_canonical_roman_numbers()
        {
            var paragraphs = new[]
            {
                P(1, "I. ĐẶT VẤN ĐỀ"),
                P(2, "IV. PHƯƠNG PHÁP"),
                P(3, "IX. KẾT QUẢ"),
                P(4, "IIII. Không canonical"),
                P(5, "IC. Không canonical"),
                P(6, "iiv. Không canonical")
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(new[] { "I", "IV", "IX" }, headings.Select(heading => heading.NumberText));
            Assert.All(headings, heading => Assert.Equal(HeadingNumberingKind.Roman, heading.Kind));
        }

        [Fact]
        public void Does_not_treat_legal_articles_points_bases_lists_or_components_as_headings()
        {
            var paragraphs = new[]
            {
                P(1, "Điều 1. Phạm vi điều chỉnh"),
                P(2, "1. Nội dung khoản", bold: true),
                P(3, "a) Nội dung điểm", bold: true),
                P(4, "Căn cứ Luật Tổ chức Chính phủ;"),
                P(5, "Xét đề nghị của Bộ trưởng;"),
                P(6, "QUYẾT ĐỊNH"),
                P(7, "PHỤ LỤC I"),
                P(8, "1.1. DÒNG CÓ ROLE CĂN CỨ", role: "legalBasis"),
                P(9, "2.1. DÒNG CÓ ROLE TRÍCH YẾU", role: "subject"),
                P(10, "1.1. TRONG BẢNG", inTable: true),
                P(11, "1.1. TRONG HEADER", storyType: "wdPrimaryHeaderStory")
            };

            Assert.Empty(_detector.Detect(paragraphs));
        }

        [Theory]
        [InlineData("1. Nội dung khoản")]
        [InlineData("1. Nội dung khoản.")]
        [InlineData("1. Nội dung khoản;")]
        [InlineData("1. Nội dung khoản:")]
        [InlineData("a) Nội dung điểm")]
        [InlineData("đ) Nội dung điểm")]
        [InlineData("Điều 2. Nội dung thi hành")]
        [InlineData("Khoản 1. Nội dung")]
        [InlineData("Điểm a) Nội dung")]
        [InlineData("Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam")]
        [InlineData("Theo đề nghị của Bộ trưởng Bộ Nội vụ")]
        [InlineData("Số: 12/QĐ-UBND")]
        [InlineData("Nơi nhận")]
        [InlineData("Kính gửi: Các cơ quan, đơn vị")]
        public void Negative_legal_and_list_corpus_is_not_a_heading(string text)
        {
            Assert.Empty(_detector.Detect(new[] { P(1, text, bold: true) }));
        }

        [Theory]
        [InlineData("01. TỔNG QUAN")]
        [InlineData("1.1. nội dung bắt đầu chữ thường")]
        [InlineData("1/2/2026 KẾ HOẠCH")]
        [InlineData("1. Mô tả một câu hoàn chỉnh.")]
        public void Ambiguous_decimal_text_is_not_a_heading(string text)
        {
            Assert.Empty(_detector.Detect(new[] { P(1, text, bold: true) }));
        }

        [Fact]
        public void Detection_context_can_supply_canonical_roles_and_logical_blocks()
        {
            var paragraphs = new[]
            {
                P(1, "1.1. SHOULD BE EXCLUDED"),
                P(2, "1. TỔNG QUAN", bold: false)
            };
            var context = new HeadingDetectionContext(
                new Dictionary<int, string> { [1] = "clause" },
                new Dictionary<int, string> { [2] = "document-b" });

            var headings = _detector.Detect(paragraphs, context);

            var heading = Assert.Single(headings);
            Assert.Equal(2, heading.ParagraphIndex);
            Assert.Equal("document-b", heading.LogicalBlockId);
        }

        [Fact]
        public void Detects_standard_unnumbered_headings_but_not_appendix_or_toc_components()
        {
            var paragraphs = new[]
            {
                P(1, "MỞ ĐẦU", bold: false),
                P(2, "KẾT LUẬN", bold: false),
                P(3, "TÀI LIỆU THAM KHẢO", bold: false),
                P(4, "PHỤ LỤC", bold: false),
                P(5, "MỤC LỤC", bold: false)
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(new[] { 1, 2, 3 }, headings.Select(heading => heading.ParagraphIndex));
            Assert.All(headings, heading => Assert.Equal(HeadingNumberingKind.Unnumbered, heading.Kind));
        }

        [Fact]
        public void Continuity_detects_gap_within_the_same_parent()
        {
            var headings = DetectAcademic(
                "1. TỔNG QUAN",
                "1.1. Cơ sở lý luận",
                "1.2. Cơ sở thực tiễn",
                "1.4. Đánh giá tác động");

            var issues = _detector.AnalyzeContinuity(headings);

            var issue = Assert.Single(issues);
            Assert.Equal(HeadingIssueKind.SkippedNumber, issue.IssueKind);
            Assert.Equal("1.3", issue.Expected);
        }

        [Fact]
        public void Continuity_keeps_sibling_state_separate_for_each_parent_prefix()
        {
            var headings = DetectAcademic(
                "1. PHẦN MỘT",
                "1.1. Mục tiêu một",
                "1.2. Mục tiêu hai",
                "2. PHẦN HAI",
                "2.1. Mục tiêu mới");

            Assert.Empty(_detector.AnalyzeContinuity(headings));
        }

        [Fact]
        public void Continuity_detects_duplicate_path()
        {
            var headings = DetectAcademic(
                "1. TỔNG QUAN",
                "1.1. Mục tiêu A",
                "1.1. Mục tiêu B");

            var issue = Assert.Single(_detector.AnalyzeContinuity(headings));

            Assert.Equal(HeadingIssueKind.DuplicateNumber, issue.IssueKind);
            Assert.Equal("1.2", issue.Expected);
        }

        [Fact]
        public void Continuity_detects_backward_number_without_losing_high_water_mark()
        {
            var headings = DetectAcademic(
                "1. TỔNG QUAN",
                "1.3. Mục tiêu C",
                "1.2. Mục tiêu B",
                "1.4. Mục tiêu D");

            var issue = Assert.Single(_detector.AnalyzeContinuity(headings));

            Assert.Equal(HeadingIssueKind.BackwardNumber, issue.IssueKind);
            Assert.Equal(3, issue.Heading.ParagraphIndex);
            Assert.Equal("1.4", issue.Expected);
        }

        [Fact]
        public void Continuity_detects_missing_parent()
        {
            var headings = DetectAcademic(
                "1. PHẦN MỘT",
                "2.1. Mục tiêu không có cha");

            var issue = Assert.Single(_detector.AnalyzeContinuity(headings));

            Assert.Equal(HeadingIssueKind.MissingParent, issue.IssueKind);
            Assert.Equal("Bổ sung đề mục cha 2 trước đề mục này.", issue.Expected);
        }

        [Fact]
        public void Continuity_detects_level_jump_before_missing_parent()
        {
            var headings = DetectAcademic(
                "1. TỔNG QUAN",
                "1.1.1. Nhảy trực tiếp đến cấp ba");

            var issue = Assert.Single(_detector.AnalyzeContinuity(headings));

            Assert.Equal(HeadingIssueKind.LevelJump, issue.IssueKind);
        }

        [Fact]
        public void Continuity_resets_at_explicit_logical_block_boundary()
        {
            var paragraphs = new[]
            {
                P(1, "1. TÀI LIỆU A", bold: false),
                P(2, "2. NỘI DUNG A", bold: false),
                P(3, "1. TÀI LIỆU B", bold: false),
                P(4, "2. NỘI DUNG B", bold: false)
            };
            var context = new HeadingDetectionContext(
                logicalBlockIdsByParagraphIndex: new Dictionary<int, string>
                {
                    [1] = "A", [2] = "A", [3] = "B", [4] = "B"
                });

            var headings = _detector.Detect(paragraphs, context);

            Assert.Empty(_detector.AnalyzeContinuity(headings));
            Assert.Equal(new[] { "A", "A", "B", "B" },
                headings.Select(heading => heading.LogicalBlockId));
        }

        [Fact]
        public void Continuity_options_can_override_blocks_for_compatibility_results()
        {
            var headings = _detector.Detect(new[]
            {
                P(1, "1. TÀI LIỆU A", bold: false),
                P(2, "1. TÀI LIỆU B", bold: false)
            });
            var options = new HeadingContinuityOptions(
                logicalBlockIdsByParagraphIndex: new Dictionary<int, string>
                {
                    [1] = "A", [2] = "B"
                });

            Assert.Empty(_detector.AnalyzeContinuity(headings, options));
        }

        [Fact]
        public void First_number_policy_is_opt_in_for_complete_sequences()
        {
            var headings = DetectAcademic("3. ĐOẠN TRÍCH");

            Assert.Empty(_detector.AnalyzeContinuity(headings));
            var issue = Assert.Single(_detector.AnalyzeContinuity(
                headings,
                new HeadingContinuityOptions(requireFirstNumberAtOne: true)));
            Assert.Equal(HeadingIssueKind.SkippedNumber, issue.IssueKind);
            Assert.Equal("1", issue.Expected);
        }

        [Fact]
        public void First_gap_policy_applies_to_each_parent_when_enabled()
        {
            var headings = DetectAcademic(
                "1. TỔNG QUAN",
                "1.2. Bắt đầu sai ở cấp hai");

            var issue = Assert.Single(_detector.AnalyzeContinuity(
                headings,
                new HeadingContinuityOptions(requireFirstNumberAtOne: true)));

            Assert.Equal(HeadingIssueKind.SkippedNumber, issue.IssueKind);
            Assert.Equal("1.1", issue.Expected);
        }

        [Fact]
        public void Roman_and_decimal_sequences_keep_independent_state()
        {
            var headings = _detector.Detect(new[]
            {
                P(1, "I. PHẦN LA MÃ"),
                P(2, "1. PHẦN THẬP PHÂN", bold: false),
                P(3, "II. PHẦN LA MÃ TIẾP")
            });

            Assert.Empty(_detector.AnalyzeContinuity(headings));
        }

        private IReadOnlyList<DetectedHeading> DetectAcademic(params string[] texts)
        {
            return _detector.Detect(texts.Select((text, index) => P(index + 1, text)));
        }
    }
}
