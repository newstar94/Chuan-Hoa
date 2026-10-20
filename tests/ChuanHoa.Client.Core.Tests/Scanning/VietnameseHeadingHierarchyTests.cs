using System.Linq;
using ChuanHoa.Client.Core.Scanning;
using Xunit;

namespace ChuanHoa.Client.Core.Tests.Scanning
{
    /// <summary>Đợt 5 — PHẦN → I. → 1. → 1.1. hierarchy of technical documents.</summary>
    public sealed class VietnameseHeadingHierarchyTests
    {
        private readonly HeadingDetector _detector = new HeadingDetector();

        private static LocalParagraphSnapshot P(int index, string text, bool bold = true) =>
            new LocalParagraphSnapshot(index, text, "wdMainTextStory", 1, index * 50,
                "Times New Roman", bold: bold);

        [Fact]
        public void Levels_follow_the_order_of_numbering_families()
        {
            var headings = _detector.Detect(new[]
            {
                P(1, "PHẦN I: THUYẾT MINH CHUNG"),
                P(2, "I. GIỚI THIỆU GÓI THẦU"),
                P(3, "1. Thông tin chung:"),
                P(4, "1.1. Tên gói thầu"),
                P(5, "Nội dung mô tả gói thầu.", bold: false)
            });

            Assert.Equal(new[] { 1, 2, 3, 4 }, headings.Select(item => item.ParagraphIndex));
            Assert.Equal(new[] { 1, 2, 3, 4 }, headings.Select(item => item.Level));
            Assert.Equal(HeadingNumberingKind.Part, headings[0].Kind);
            Assert.Equal("Thông tin chung", headings[2].TitleText);
        }

        [Fact]
        public void Roman_headings_are_level_one_and_decimals_level_two_beneath_them()
        {
            var headings = _detector.Detect(new[]
            {
                P(1, "I. CĂN CỨ LẬP BIỆN PHÁP"),
                P(2, "1. Hồ sơ mời thầu"),
                P(3, "II. BIỆN PHÁP THI CÔNG")
            });

            Assert.Equal(new[] { 1, 2, 1 }, headings.Select(item => item.Level));
        }

        [Fact]
        public void Bold_level_one_decimal_without_an_upper_heading_is_still_a_clause()
        {
            Assert.Empty(_detector.Detect(new[] { P(1, "1. Thông tin chung:") }));
        }

        [Fact]
        public void Decimal_numbering_restarts_under_each_roman_heading()
        {
            var headings = _detector.Detect(new[]
            {
                P(1, "I. CĂN CỨ"),
                P(2, "1. Hồ sơ mời thầu"),
                P(3, "2. Tiêu chuẩn áp dụng"),
                P(4, "II. BIỆN PHÁP"),
                P(5, "1. Công tác chuẩn bị"),
                P(6, "2. Công tác đất")
            });

            Assert.Empty(_detector.AnalyzeContinuity(headings));
        }

        [Fact]
        public void Duplicate_decimal_under_the_same_roman_heading_is_still_reported()
        {
            var headings = _detector.Detect(new[]
            {
                P(1, "I. CĂN CỨ"),
                P(2, "1. Hồ sơ mời thầu"),
                P(3, "1. Tiêu chuẩn áp dụng")
            });

            var issue = Assert.Single(_detector.AnalyzeContinuity(headings));
            Assert.Equal(HeadingIssueKind.DuplicateNumber, issue.IssueKind);
        }

        [Fact]
        public void Lower_case_part_keyword_in_a_sentence_is_not_a_heading()
        {
            Assert.Empty(_detector.Detect(new[] { P(1, "Phần 1 của báo cáo trình bày kết quả", bold: false) }));
        }
    }
}
