using ChuanHoa.Client.Core.Licensing;

namespace ChuanHoa.Client.Core.Tests;

public sealed class ProductFeaturePolicyTests
{
    [Fact]
    public void Free_plan_contains_only_the_three_approved_capabilities() =>
        Assert.Equal(new[] { "FORMAT_SCAN", "SPELLING_SCAN", "TABLE_IMAGE_TOOLS" }, ProductFeaturePolicy.FreeFeatures);

    [Theory]
    [InlineData("btnLapDongTieuDe")]
    [InlineData("btnChuanHoaBang")]
    [InlineData("btnChuanHoaAnh")]
    [InlineData("btnCanDinhO")]
    [InlineData("btnCanGiuaO")]
    [InlineData("btnXoaKyTuThuaBangExcel")]
    public void All_table_and_image_commands_use_the_narrow_free_capability(string command) =>
        Assert.Equal("TABLE_IMAGE_TOOLS", ProductFeaturePolicy.RequiredFeature(command));

    [Theory]
    [InlineData("btnScaleTang", "DOCUMENT_TOOLS")]
    [InlineData("btnParagraph", "DOCUMENT_TOOLS")]
    [InlineData("btnAutoFixAll2026", "AUTOFIX")]
    [InlineData("btnSuaTatCaChinhTa", "AUTOFIX")]
    [InlineData("btnThietLap", "")]
    [InlineData("unknown", "DOCUMENT_TOOLS")]
    public void Other_commands_do_not_inherit_free_table_permissions(string command, string feature) =>
        Assert.Equal(feature, ProductFeaturePolicy.RequiredFeature(command));
}
