using ChuanHoa.Client.Core.Text;

namespace ChuanHoa.Client.Core.Tests;

public sealed class ParagraphIndentPolicyTests
{
    [Theory]
    [InlineData("- Tên gói thầu")]
    [InlineData("-\tTên gói thầu")]
    [InlineData("– Nguồn vốn")]
    [InlineData("— Thời gian thực hiện")]
    [InlineData("• Nội dung")]
    [InlineData("   - Nội dung")]
    public void Recognizes_dash_and_bullet_list_paragraphs(string text)
    {
        Assert.True(ParagraphIndentPolicy.IsDashListParagraph(text));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Nội dung thông thường")]
    [InlineData("-5 ngày")]
    [InlineData("2. Nội dung gói thầu")]
    public void Rejects_non_list_paragraphs(string? text)
    {
        Assert.False(ParagraphIndentPolicy.IsDashListParagraph(text));
    }

    [Fact]
    public void Uses_stable_hanging_indent_geometry()
    {
        Assert.Equal(10d, ParagraphIndentPolicy.ListMarkerMillimeters);
        Assert.Equal(15d, ParagraphIndentPolicy.ListTextMillimeters);
        Assert.Equal(10d, ParagraphIndentPolicy.BodyFirstLineMillimeters);
    }

    [Theory]
    [InlineData("- Căn cứ Luật", -5, 15, false)]
    [InlineData("- Căn cứ Luật", 10, 0, true)]
    [InlineData("- Cam kết", -5, 25, false)]
    [InlineData("- Cam kết", 0, 15, false)]
    [InlineData("Nội dung", -5, 15, false)]
    [InlineData("Nội dung", 10, 0, true)]
    [InlineData("Nội dung", 10, 15, false)]
    public void Checks_effective_marker_and_text_positions(string text, double first, double left, bool expected)
    {
        Assert.Equal(expected, ParagraphIndentPolicy.IsValidIndent(text, first, left, 10, 12.7));
    }

    [Theory]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\vĐộc lập - Tự do - Hạnh phúc\v---------------", 2)]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\vĐộc lập - Tự do - Hạnh phúc", 1)]
    [InlineData("Nội dung\v---------------", 0)]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\vĐộc lập - Tự do - Hạnh phúc\vNội dung", 0)]
    public void Splits_only_complete_combined_header_lines(string text, int expected)
    {
        Assert.Equal(expected, CombinedNationalHeader.GetBreakOffsets(text).Length);
    }
}
