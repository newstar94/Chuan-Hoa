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
}
