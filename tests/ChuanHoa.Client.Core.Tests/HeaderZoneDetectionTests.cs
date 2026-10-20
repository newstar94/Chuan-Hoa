using System.Text;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;

namespace ChuanHoa.Client.Core.Tests;

/// <summary>
/// Đợt 1 — header-zone recognition. Every snapshot goes through the real
/// detector with Role = "Unknown"; no role is pre-assigned.
/// </summary>
public sealed class HeaderZoneDetectionTests
{
    [Fact]
    public void Legal_basis_citing_the_constitution_is_not_the_national_title()
    {
        var roles = Detect(
            P(1, "ỦY BAN NHÂN DÂN TỈNH A"),
            P(2, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"),
            P(3, "Độc lập - Tự do - Hạnh phúc"),
            P(4, "Số: 01/QĐ-UBND"),
            P(5, "QUYẾT ĐỊNH"),
            P(6, "Về việc ban hành quy chế"),
            P(7, "Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam;"),
            P(8, "Căn cứ Luật Tổ chức chính quyền địa phương;"));

        Assert.Equal("nationalTitle", roles[2]);
        Assert.Equal("nationalMotto", roles[3]);
        Assert.NotEqual("nationalTitle", Role(roles, 7));
        Assert.Equal("legalBasis", roles[7]);
    }

    [Fact]
    public void Body_sentence_naming_the_country_is_not_the_national_title()
    {
        var roles = Detect(
            P(1, "THUYẾT MINH BIỆN PHÁP TỔ CHỨC THI CÔNG"),
            P(2, "I. CĂN CỨ LẬP BIỆN PHÁP"),
            P(3, "Nhà thầu cam kết tuân thủ pháp luật nước Cộng hòa xã hội chủ nghĩa Việt Nam trong suốt quá trình thi công."));

        Assert.DoesNotContain(roles, item => item.Value == "nationalTitle");
    }

    [Fact]
    public void Body_sentence_about_independence_and_happiness_is_not_the_motto()
    {
        var roles = Detect(
            P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"),
            P(2, "Độc lập - Tự do - Hạnh phúc"),
            P(3, "BÁO CÁO"),
            P(4, "Về công tác tu bổ đền thờ liệt sĩ"),
            P(5, "1. Khái quát chung"),
            P(6, "Các anh hùng liệt sĩ đã hy sinh vì độc lập dân tộc và hạnh phúc của nhân dân."));

        Assert.Equal("nationalMotto", roles[2]);
        Assert.Null(Role(roles, 6));
    }

    [Fact]
    public void Motto_is_only_recognized_directly_below_the_national_title()
    {
        var roles = Detect(
            P(1, "CÔNG TY CỔ PHẦN XÂY DỰNG A"),
            P(2, "Độc lập - Tự do - Hạnh phúc"),
            P(3, "Số: 12/BC-XD"));

        Assert.NotEqual("nationalMotto", Role(roles, 2));
    }

    [Fact]
    public void Decomposed_unicode_header_is_recognized()
    {
        var title = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM".Normalize(NormalizationForm.FormD);
        var motto = "Độc lập - Tự do - Hạnh phúc".Normalize(NormalizationForm.FormD);
        var roles = Detect(
            P(1, "ỦY BAN NHÂN DÂN XÃ B"),
            P(2, title),
            P(3, motto),
            P(4, "Số: 05/TB-UBND"),
            P(5, "THÔNG BÁO"));

        Assert.Equal("nationalTitle", roles[2]);
        Assert.Equal("nationalMotto", roles[3]);
    }

    [Fact]
    public void Header_with_nbsp_and_en_dashes_is_recognized()
    {
        var roles = Detect(
            P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"),
            P(2, "Độc lập – Tự do – Hạnh phúc"),
            P(3, "QUYẾT ĐỊNH"));

        Assert.Equal("nationalTitle", roles[1]);
        Assert.Equal("nationalMotto", roles[2]);
    }

    [Fact]
    public void Data_table_cells_are_never_type_names()
    {
        var snapshot = Snapshot(
            P(1, "THUYẾT MINH BIỆN PHÁP THI CÔNG"),
            P(2, "I. HỒ SƠ QUẢN LÝ CHẤT LƯỢNG"),
            Cell(3, "STT", table: 1, row: 1, cell: 1),
            Cell(4, "Biên bản", table: 1, row: 2, cell: 2),
            Cell(5, "Báo cáo", table: 1, row: 3, cell: 2),
            Cell(6, "Kế hoạch", table: 1, row: 4, cell: 2),
            Cell(7, "Quy định", table: 1, row: 5, cell: 2));

        var detector = new DocumentRoleDetector();
        var roles = detector.Detect(snapshot);

        Assert.DoesNotContain(roles, item => item.Value == "typeName");
        Assert.Equal(LocalDocumentTypeCodes.Unknown, detector.ResolveDocumentType(snapshot));
        Assert.Single(detector.DetectBlocks(snapshot));
    }

    [Fact]
    public void Leading_data_table_cell_is_not_a_type_name()
    {
        var roles = Detect(
            Cell(1, "Báo cáo", table: 1, row: 1, cell: 1),
            Cell(2, "Tuần 1", table: 1, row: 1, cell: 2),
            P(3, "Nội dung"));

        Assert.DoesNotContain(roles, item => item.Value == "typeName");
    }

    [Fact]
    public void Table_cell_starting_with_so_does_not_make_body_headings_organ_names()
    {
        var roles = Detect(
            P(1, "THUYẾT MINH BIỆN PHÁP THI CÔNG"),
            P(2, "III. TIÊU CHUẨN ÁP DỤNG"),
            P(3, "QUY CHUẨN KỸ THUẬT QUỐC GIA"),
            Cell(4, "Số 1 trong danh mục tiêu chuẩn", table: 1, row: 1, cell: 1));

        Assert.DoesNotContain(roles, item => item.Value == "organName" || item.Value == "superiorOrganName");
        Assert.DoesNotContain(roles, item => item.Value == "codeNumber");
    }

    [Fact]
    public void Organ_heading_without_national_title_or_type_is_a_technical_block()
    {
        var roles = Detect(
            P(1, "CÔNG TY CỔ PHẦN XÂY DỰNG A"),
            P(2, "Số: 12/2026"),
            P(3, "THUYẾT MINH BIỆN PHÁP THI CÔNG"),
            P(4, "I. GIỚI THIỆU CHUNG"));

        Assert.DoesNotContain(roles, item => item.Value == "organName");
    }

    [Fact]
    public void Header_components_after_the_body_starts_are_not_header_roles()
    {
        var roles = Detect(
            P(1, "QUYẾT ĐỊNH"),
            P(2, "Về việc phê duyệt"),
            P(3, "Căn cứ Luật Đấu thầu số 22/2023/QH15;"),
            P(4, "Điều 1. Phê duyệt kế hoạch."),
            P(5, "Hà Nội, ngày 01 tháng 09 năm 2026"),
            P(6, "Số: 99/QĐ-ABC"));

        Assert.NotEqual("placeAndIssuedDate", Role(roles, 5));
        Assert.NotEqual("codeNumber", Role(roles, 6));
    }

    [Fact]
    public void Header_in_layout_table_with_motto_in_row_below_is_recognized()
    {
        var roles = Detect(
            Cell(1, "ỦY BAN NHÂN DÂN", table: 1, row: 1, cell: 1),
            Cell(2, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", table: 1, row: 1, cell: 2),
            Cell(3, "Số: 01/KH-UBND", table: 1, row: 2, cell: 1),
            Cell(4, "Độc lập - Tự do - Hạnh phúc", table: 1, row: 2, cell: 2),
            P(5, "KẾ HOẠCH"),
            P(6, "Triển khai nhiệm vụ năm 2026"),
            P(7, "I. MỤC ĐÍCH"));

        Assert.Equal("nationalTitle", roles[2]);
        Assert.Equal("nationalMotto", roles[4]);
        Assert.Equal("codeNumber", roles[3]);
        Assert.Equal("organName", roles[1]);
        Assert.Equal("typeName", roles[5]);
    }

    [Fact]
    public void Lower_case_title_is_accepted_only_inside_the_header_zone()
    {
        var roles = Detect(
            P(1, "Cộng hòa xã hội chủ nghĩa Việt Nam"),
            P(2, "Độc lập - Tự do - Hạnh phúc"),
            P(3, "THÔNG BÁO"));

        Assert.Equal("nationalTitle", roles[1]);
        Assert.False(ComponentTextNormalizer.IsUppercaseNationalTitle("Cộng hòa xã hội chủ nghĩa Việt Nam"));
        Assert.True(ComponentTextNormalizer.IsUppercaseNationalTitle("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM"));
    }

    [Theory]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true)]
    [InlineData("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", true)]
    [InlineData("  CỘNG  HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM  ", true)]
    [InlineData("Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam;", false)]
    [InlineData("Quốc hội nước Cộng hòa xã hội chủ nghĩa Việt Nam khóa XV", false)]
    public void National_title_must_be_the_whole_paragraph(string text, bool expected)
    {
        Assert.Equal(expected, ComponentTextNormalizer.IsNationalTitle(text));
    }

    [Theory]
    [InlineData("Độc lập - Tự do - Hạnh phúc", true)]
    [InlineData("Độc lập-Tự do-Hạnh phúc", true)]
    [InlineData("ĐỘC LẬP – TỰ DO – HẠNH PHÚC", true)]
    [InlineData("Độc lập - Tự do - Hạnh phúc\v---------------", true)]
    [InlineData("…hy sinh vì độc lập dân tộc và hạnh phúc của nhân dân.", false)]
    [InlineData("Độc lập, tự chủ và hạnh phúc", false)]
    public void Motto_must_be_the_whole_paragraph(string text, bool expected)
    {
        Assert.Equal(expected, ComponentTextNormalizer.IsNationalMotto(text));
    }

    [Fact]
    public void Long_dash_run_does_not_time_out()
    {
        var text = new string('-', 5000) + "x";
        Assert.False(ComponentTextNormalizer.IsNationalMotto(text));
        Assert.Equal(text, ComponentTextNormalizer.Normalize(text));
    }

    [Fact]
    public void Motto_edit_plan_changes_only_separators_and_letters()
    {
        const string source = "Độc lập-Tự Do -  Hạnh phúc";
        var edits = NationalMottoEditPlanner.Plan(source);

        Assert.NotNull(edits);
        Assert.Equal(NationalMottoEditPlanner.Expected, NationalMottoEditPlanner.Apply(source, edits!));
        Assert.All(edits!, edit => Assert.True(edit.Length < source.Length / 2));
    }

    [Fact]
    public void Motto_edit_plan_never_touches_body_text()
    {
        Assert.Null(NationalMottoEditPlanner.Plan("Các anh hùng đã hy sinh vì độc lập dân tộc và hạnh phúc của nhân dân."));
        Assert.Null(NationalMottoEditPlanner.Plan("Độc lập - Tự do - Hạnh phúc và ấm no"));
        Assert.Empty(NationalMottoEditPlanner.Plan(NationalMottoEditPlanner.Expected)!);
    }

    [Fact]
    public void Motto_edit_plan_skips_decomposed_text()
    {
        Assert.Null(NationalMottoEditPlanner.Plan("Độc lập-Tự do-Hạnh phúc".Normalize(NormalizationForm.FormD)));
    }

    [Theory]
    [InlineData("Số 5 qđ abc", "Số: 05/QĐ-ABC", true)]
    [InlineData("Các anh hùng đã hy sinh vì độc lập dân tộc và hạnh phúc của nhân dân.", "Độc lập - Tự do - Hạnh phúc", false)]
    [InlineData("Số: 01/QĐ-UBND", "Số: 01/QĐ-UBND", true)]
    public void Whole_paragraph_replacement_guard(string original, string replacement, bool expected)
    {
        Assert.Equal(expected, ComponentTextNormalizer.IsSafeWholeParagraphReplacement(original, replacement));
    }

    [Theory]
    [InlineData("UBND XÃ HOÀNG HOA THÁM\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true)]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", false)]
    [InlineData("Ghi chú\tnội dung", false)]
    public void Detects_one_line_tab_separated_organ_and_title(string text, bool expected)
    {
        Assert.Equal(expected, ComponentTextNormalizer.IsTabSeparatedOrganAndNationalTitle(text));
    }

    [Fact]
    public void Header_zone_is_exposed_on_the_block()
    {
        var snapshot = Snapshot(
            P(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"),
            P(2, "Độc lập - Tự do - Hạnh phúc"),
            P(3, "QUYẾT ĐỊNH"),
            P(4, "Về việc phê duyệt"),
            P(5, "Căn cứ Luật Đấu thầu số 22/2023/QH15;"),
            P(6, "Điều 1. Phê duyệt."));

        var block = Assert.Single(new DocumentRoleDetector().DetectBlocks(snapshot));

        Assert.True(block.IsAdministrative);
        Assert.True(block.IsInHeaderZone(4));
        Assert.False(block.IsInHeaderZone(5));
    }

    private static Dictionary<int, string> Detect(params LocalParagraphSnapshot[] paragraphs) =>
        new DocumentRoleDetector().Detect(Snapshot(paragraphs));

    private static string? Role(IReadOnlyDictionary<int, string> roles, int index) =>
        roles.TryGetValue(index, out var role) ? role : null;

    private static LocalScanSnapshot Snapshot(params LocalParagraphSnapshot[] paragraphs) =>
        new("sha256:header-zone", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

    private static LocalParagraphSnapshot P(int index, string text) =>
        new(index, text, "wdMainTextStory", 1, index * 100, "Times New Roman",
            fontSizePoints: 13, bold: true, alignment: 1, pageNumber: 1);

    private static LocalParagraphSnapshot Cell(int index, string text, int table, int row, int cell) =>
        new(index, text, "wdMainTextStory", 1, index * 100, "Times New Roman",
            tableIndex: table, rowIndex: row, cellIndex: cell, isInTable: true,
            fontSizePoints: 13, bold: true, alignment: 1, pageNumber: 1);
}
