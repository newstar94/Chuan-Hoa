using System.Linq;
using ChuanHoa.Client.Core.Scanning;
using Xunit;

namespace ChuanHoa.Client.Core.Tests.Scanning
{
    public sealed class HeadingDetectorTests
    {
        private readonly HeadingDetector _detector = new HeadingDetector();

        private static LocalParagraphSnapshot P(int index, string text, bool bold = true, int? outlineLevel = null,
            bool? keepWithNext = null, string? styleName = null, bool inTable = false) =>
            new LocalParagraphSnapshot(index, text, "wdMainTextStory", 1, index * 50, "Times New Roman",
                bold: bold, outlineLevel: outlineLevel, keepWithNext: keepWithNext, styleName: styleName,
                isInTable: inTable);

        [Fact]
        public void Detects_decimal_headings_with_various_spacing_and_symbols()
        {
            var paragraphs = new[]
            {
                P(1, "1.   Tổng quan nghiên cứu"),
                P(2, "1.1. Mục tiêu cụ thể"),
                P(3, "1.1.1- Phương pháp thực hiện"),
                P(4, "Nội dung thân bài mô tả chi tiết phương pháp...")
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(3, headings.Count);
            Assert.Equal(1, headings[0].Level);
            Assert.Equal("1", headings[0].NumberText);
            Assert.Equal("Tổng quan nghiên cứu", headings[0].TitleText);

            Assert.Equal(2, headings[1].Level);
            Assert.Equal("1.1", headings[1].NumberText);
            Assert.Equal("Mục tiêu cụ thể", headings[1].TitleText);

            Assert.Equal(3, headings[2].Level);
            Assert.Equal("1.1.1", headings[2].NumberText);
            Assert.Equal("Phương pháp thực hiện", headings[2].TitleText);
        }

        [Fact]
        public void Detects_roman_and_article_headings()
        {
            var paragraphs = new[]
            {
                P(1, "I. ĐẶT VẤN ĐỀ"),
                P(2, "Điều 1. Phạm vi điều chỉnh"),
                P(3, "Điều 2. Đối tượng áp dụng")
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(3, headings.Count);
            Assert.Equal(HeadingNumberingKind.Roman, headings[0].Kind);
            Assert.Equal("I", headings[0].NumberText);
            Assert.Equal("ĐẶT VẤN ĐỀ", headings[0].TitleText);

            Assert.Equal(HeadingNumberingKind.Article, headings[1].Kind);
            Assert.Equal("Điều 1", headings[1].NumberText);
            Assert.Equal("Phạm vi điều chỉnh", headings[1].TitleText);

            Assert.Equal(HeadingNumberingKind.Article, headings[2].Kind);
            Assert.Equal("Điều 2", headings[2].NumberText);
        }

        [Fact]
        public void Detects_standard_unnumbered_headings()
        {
            var paragraphs = new[]
            {
                P(1, "MỞ ĐẦU"),
                P(2, "Nội dung văn bản mở đầu ở đây..."),
                P(3, "KẾT LUẬN")
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Equal(2, headings.Count);
            Assert.Equal("MỞ ĐẦU", headings[0].TitleText);
            Assert.Equal(HeadingNumberingKind.Unnumbered, headings[0].Kind);
            Assert.Equal("KẾT LUẬN", headings[1].TitleText);
        }

        [Fact]
        public void Ignores_body_text_ending_in_semicolon_or_colon()
        {
            var paragraphs = new[]
            {
                P(1, "1. Thực hiện đầy đủ các nhiệm vụ được giao sau đây:"),
                P(2, "a) Hoàn thành báo cáo đúng hạn;"),
                P(3, "b) Nộp lưu chiểu hồ sơ.")
            };

            var headings = _detector.Detect(paragraphs);

            Assert.Empty(headings);
        }

        [Fact]
        public void Continuity_detects_skipped_decimal_numbers()
        {
            var paragraphs = new[]
            {
                P(1, "1.1. Cơ sở lý luận"),
                P(2, "1.2. Cơ sở thực tiễn"),
                P(3, "1.4. Đánh giá tác động") // Missing 1.3
            };

            var headings = _detector.Detect(paragraphs);
            var issues = _detector.AnalyzeContinuity(headings);

            Assert.Single(issues);
            Assert.Equal(HeadingIssueKind.SkippedNumber, issues[0].IssueKind);
            Assert.Contains("1.4", issues[0].CurrentIssue);
            Assert.Equal("1.3", issues[0].Expected);
        }

        [Fact]
        public void Continuity_detects_duplicate_numbers()
        {
            var paragraphs = new[]
            {
                P(1, "1.1. Mục tiêu A"),
                P(2, "1.1. Mục tiêu B") // Duplicate 1.1
            };

            var headings = _detector.Detect(paragraphs);
            var issues = _detector.AnalyzeContinuity(headings);

            Assert.Single(issues);
            Assert.Equal(HeadingIssueKind.DuplicateNumber, issues[0].IssueKind);
        }
    }
}
