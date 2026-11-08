using ChuanHoa.Client.Core.Text;

namespace ChuanHoa.Client.Core.Tests;

public sealed class BrokenLineJoinerTests
{
    [Fact]
    public void PdfWrappedSentenceIsJoinedUntilItsSemicolon()
    {
        var lines = new[]
        {
            "Căn cứ Quyết định số 4869/QĐ-UBND ngày 22/12/2025 của UBND tỉnh",
            "Quảng Ninh (nay là UBND thành phố) về việc giao dự toán thu ngân sách nhà",
            "nước, chi ngân sách địa phương, phân bổ dự toán chi ngân sách cấp tỉnh năm",
            "2026;",
            "Căn cứ Nghị quyết số 15/NQ-HĐND ngày 10/12/2025;",
        };

        Assert.Equal(
            "Căn cứ Quyết định số 4869/QĐ-UBND ngày 22/12/2025 của UBND tỉnh Quảng Ninh (nay là UBND thành phố) " +
            "về việc giao dự toán thu ngân sách nhà nước, chi ngân sách địa phương, phân bổ dự toán chi ngân sách " +
            "cấp tỉnh năm 2026;\nCăn cứ Nghị quyết số 15/NQ-HĐND ngày 10/12/2025;",
            Join(lines));
    }

    [Theory]
    [InlineData("Kết thúc câu.", "Câu tiếp theo")]
    [InlineData("Nơi nhận:", "Như Điều 3")]
    [InlineData("theo quy định;", "b) Nội dung")]
    [InlineData("Ông ấy hỏi: “Đúng không?”", "rồi đi")]
    [InlineData("Các nhiệm vụ gồm", "- Rà soát hồ sơ")]
    [InlineData("Các nhiệm vụ gồm", "+ Rà soát hồ sơ")]
    [InlineData("Các nhiệm vụ gồm", "1. Rà soát hồ sơ")]
    [InlineData("Các nhiệm vụ gồm", "1.2. Rà soát hồ sơ")]
    [InlineData("Các nhiệm vụ gồm", "a) Rà soát hồ sơ")]
    [InlineData("Các nhiệm vụ gồm", "II. NHIỆM VỤ")]
    [InlineData("Các nhiệm vụ gồm", "Điều 2. Tổ chức thực hiện")]
    [InlineData("Phạm vi điều chỉnh", "Chương II")]
    [InlineData("QUYẾT ĐỊNH", "Về việc giao dự toán")]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc")]
    [InlineData("Ủy ban nhân dân tỉnh", "ỦY BAN NHÂN DÂN")]
    [InlineData("", "Dòng sau dòng trống")]
    [InlineData("Dòng trước dòng trống", "   ")]
    [InlineData("1. Mục tiêu", "Nâng cao chất lượng dự toán")]
    [InlineData("Độc lập -", "Tự do")]
    [InlineData("Thời gian thực hiện", "Địa điểm tổ chức")]
    [InlineData("Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015", "Căn cứ Nghị định số 30/2020/NĐ-CP")]
    [InlineData("Căn cứ Luật Ngân sách nhà nước ngày 25 tháng 6 năm 2015 và các văn bản", "Xét đề nghị của Giám đốc Sở Tài chính")]
    public void RealBoundariesKeepTheirBreak(string previous, string next)
    {
        Assert.Equal(LineBreakDecision.Keep, BrokenLineJoiner.Decide(previous, next));
    }

    [Theory]
    [InlineData("Căn cứ Quyết định số 4869/QĐ-UBND ngày 22/12/2025 của UBND tỉnh", "Quảng Ninh (nay là UBND thành phố)")]
    [InlineData("(nay là UBND thành phố)", "về việc giao dự toán")]
    [InlineData("1. Mục tiêu", "của kế hoạch là nâng cao chất lượng")]
    [InlineData("dự toán chi ngân sách cấp tỉnh năm", "2026. Giao Sở Tài chính")]
    [InlineData("Thực hiện nhiệm vụ thu ngân sách theo phân cấp quy định tại", "Điều 5 Luật Ngân sách nhà nước")]
    [InlineData("các chỉ tiêu phát triển kinh tế - xã hội giai đoạn 2026 -", "2030 đã được thông qua")]
    [InlineData("ngân sách  ", "  nhà nước")]
    public void WrapsInsideASentenceJoinWithOneSpace(string previous, string next)
    {
        Assert.Equal(LineBreakDecision.JoinWithSpace, BrokenLineJoiner.Decide(previous, next));
    }

    [Theory]
    [InlineData("Quyết định số 4869/QĐ-", "UBND ngày 22/12/2025")]
    [InlineData("Quyết định số 4869/", "QĐ-UBND ngày 22/12/2025")]
    [InlineData("giao dự toán", ", phân bổ dự toán")]
    public void WrapsInsideATokenJoinWithoutSpace(string previous, string next)
    {
        Assert.Equal(LineBreakDecision.JoinWithoutSpace, BrokenLineJoiner.Decide(previous, next));
    }

    [Fact]
    public void BlankCountsCoverSpacesTabsAndNoBreakSpaces()
    {
        Assert.Equal(3, BrokenLineJoiner.TrailingBlankCount("nhà \t "));
        Assert.Equal(2, BrokenLineJoiner.LeadingBlankCount("  nước"));
        Assert.Equal(0, BrokenLineJoiner.TrailingBlankCount("nhà"));
    }

    private static string Join(IReadOnlyList<string> lines)
    {
        var builder = new System.Text.StringBuilder(lines[0]);
        for (var index = 1; index < lines.Count; index++)
        {
            switch (BrokenLineJoiner.Decide(lines[index - 1], lines[index]))
            {
                case LineBreakDecision.JoinWithSpace: builder.Append(' '); break;
                case LineBreakDecision.Keep: builder.Append('\n'); break;
            }
            builder.Append(lines[index]);
        }
        return builder.ToString();
    }
}
