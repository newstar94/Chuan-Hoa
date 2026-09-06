using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Tests;

public sealed class HeaderLayoutPolicyTests
{
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
