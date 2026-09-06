using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Tests;

public sealed class HeaderLayoutPolicyTests
{
    [Fact]
    public void Draft_number_and_date_keep_organization_header_centered()
    {
        var texts = new[] { "SỞ Y TẾ TÂY NINH", "TRUNG TÂM Y TẾ KHU VỰC TÂN CHÂU",
            "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc",
            "Số:.../QĐ-TTYT", "Tân Châu, ngày … tháng 08 năm 2026", "QUYẾT ĐỊNH" };
        var paragraphs = texts.Select((text, i) => new LocalParagraphSnapshot(i + 1, text,
            "wdMainTextStory", 1, i * 100, "Times New Roman")).ToArray();
        var snapshot = new LocalScanSnapshot("draft", 1, Array.Empty<LocalSectionSnapshot>(),
            paragraphs, Array.Empty<AnnotationProtectedSpan>());
        var detector = new DocumentRoleDetector();
        var roles = detector.Detect(snapshot);
        Assert.Equal("organName", roles[2]);
        Assert.Equal("placeAndIssuedDate", roles[6]);
        Assert.Equal(1, HeaderLayoutPolicy.DateAlignment(detector.DetectBlocks(snapshot), 6));
    }
    [Fact]
    public void Organization_after_date_in_table_order_still_requires_center_alignment()
    {
        var roles = new[] { "nationalTitle", "nationalMotto", "placeAndIssuedDate", "organName", "typeName" };
        var texts = new[] { "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc",
            "Hà Nội, ngày 14 tháng 08 năm 2026", "CƠ QUAN BAN HÀNH", "QUYẾT ĐỊNH" };
        var paragraphs = roles.Select((role, i) => new LocalParagraphSnapshot(i + 1, texts[i],
            "wdMainTextStory", 1, i * 100, "Times New Roman", role: role)).ToArray();
        var snapshot = new LocalScanSnapshot("table-order", 1, Array.Empty<LocalSectionSnapshot>(),
            paragraphs, Array.Empty<AnnotationProtectedSpan>());
        Assert.Equal(1, HeaderLayoutPolicy.DateAlignment(new DocumentRoleDetector().DetectBlocks(snapshot), 3));
    }
    [Theory]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    public void Date_alignment_depends_on_its_own_header(bool hasOrgan, int expected)
    {
        var paragraphs = new List<LocalParagraphSnapshot>();
        void Add(string role, string text) => paragraphs.Add(new LocalParagraphSnapshot(
            paragraphs.Count + 1, text, "wdMainTextStory", 1, paragraphs.Count * 100,
            "Times New Roman", role: role));
        if (hasOrgan) Add("organName", "CƠ QUAN BAN HÀNH");
        Add("nationalTitle", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
        Add("nationalMotto", "Độc lập - Tự do - Hạnh phúc");
        Add("placeAndIssuedDate", "Hà Nội, ngày 14 tháng 08 năm 2026");
        var snapshot = new LocalScanSnapshot("header", 1, Array.Empty<LocalSectionSnapshot>(),
            paragraphs, Array.Empty<AnnotationProtectedSpan>());
        Assert.Equal(expected, HeaderLayoutPolicy.DateAlignment(
            new DocumentRoleDetector().DetectBlocks(snapshot), paragraphs.Count));
    }
}
