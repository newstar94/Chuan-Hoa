using ChuanHoa.Client.Core.Text;

namespace ChuanHoa.Client.Core.Tests;

public sealed class VietnameseTextConversionTests
{
    [Fact]
    public void Tcvn3LowerConvertsToUnicode()
    {
        var source = "V" + (char)0xB8 + "n b" + (char)0xB6 + "n";

        var result = VietnameseLegacyEncodingConverter.Convert(".VnTime", source);

        Assert.True(result.Converted);
        Assert.Equal("Ván bản", result.Text);
        Assert.Equal(0, result.UnmappedCharacters);
    }

    [Fact]
    public void Tcvn3UpperUsesUppercaseMapping()
    {
        var source = "VI" + (char)0xD6 + "T NAM";

        var result = VietnameseLegacyEncodingConverter.Convert(".VnTimeH", source);

        Assert.Equal("VIỆT NAM", result.Text);
    }

    [Fact]
    public void VniCombiningBytesComposeToNfc()
    {
        var source = "Ho" + (char)0xF8;

        var result = VietnameseLegacyEncodingConverter.Convert("VNI-Times", source);

        Assert.Equal("Hò".Normalize(System.Text.NormalizationForm.FormC), result.Text);
        Assert.Equal("Hò", result.Text);
    }

    [Fact]
    public void UnicodeFontAndTextRemainUnchanged()
    {
        const string source = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";

        var result = VietnameseLegacyEncodingConverter.Convert("Times New Roman", source);

        Assert.False(result.Converted);
        Assert.Equal(source, result.Text);
    }

    [Theory]
    [InlineData("hòa thủy khóa", "hoà thuỷ khoá")]
    [InlineData("HÒA THỦY KHÓA", "HOÀ THUỶ KHOÁ")]
    [InlineData("ngoài ra, quyết định chuyển giao", "ngoài ra, quyết định chuyển giao")]
    public void MainVowelStyleUsesOaUyPlacement(string source, string expected)
    {
        Assert.Equal(expected,
            VietnameseTonePlacementNormalizer.Normalize(source, VietnameseTonePlacementStyle.MainVowel));
    }

    [Theory]
    [InlineData("hoà thuỷ khoá", "hòa thủy khóa")]
    [InlineData("HOÀ THUỶ KHOÁ", "HÒA THỦY KHÓA")]
    public void FirstVowelStyleUsesOaUyPlacement(string source, string expected)
    {
        Assert.Equal(expected,
            VietnameseTonePlacementNormalizer.Normalize(source, VietnameseTonePlacementStyle.FirstVowel));
    }

    [Fact]
    public void TonePlacementDoesNotNormalizeIY()
    {
        const string source = "hy vọng, kỷ niệm, lý luận, mỹ thuật, tỷ lệ";

        Assert.Equal(source,
            VietnameseTonePlacementNormalizer.Normalize(source, VietnameseTonePlacementStyle.MainVowel));
        Assert.Equal(source,
            VietnameseTonePlacementNormalizer.Normalize(source, VietnameseTonePlacementStyle.FirstVowel));
    }

    [Theory]
    [InlineData("Đây  là   văn   bản.", "Đây là văn bản.")]
    [InlineData("Hà Nội , ngày 15 tháng 8 .", "Hà Nội, ngày 15 tháng 8.")]
    [InlineData("Công tác:chỉ đạo;điều hành", "Công tác: chỉ đạo; điều hành")]
    [InlineData("Số: 123/QĐ-UBND", "Số: 123/QĐ-UBND")]
    [InlineData("Tỷ lệ 1,5% hoặc 10.000 đồng", "Tỷ lệ 1,5% hoặc 10.000 đồng")]
    [InlineData("Nội dung ( trong ngoặc ) và [ ghi chú ]", "Nội dung (trong ngoặc) và [ghi chú]")]
    public void CleanWhitespaceAndPunctuationNormalizesProperly(string source, string expected)
    {
        Assert.Equal(expected, VietnameseTypographyCleaner.CleanWhitespaceAndPunctuation(source));
    }

    [Theory]
    [InlineData("\"Nội dung văn bản\"", "“Nội dung văn bản”")]
    [InlineData("Báo cáo \" kết quả thực hiện \" năm 2026", "Báo cáo “kết quả thực hiện” năm 2026")]
    public void NormalizeQuotationMarksConvertsToVietnameseCurlyQuotes(string source, string expected)
    {
        Assert.Equal(expected, VietnameseTypographyCleaner.NormalizeQuotationMarks(source));
    }
}
