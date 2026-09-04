using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class DocumentRoleDetectorTests
{
    [Theory]
    [InlineData("NGHỊ QUYẾT", LocalDocumentTypeCodes.Resolution)]
    [InlineData("QUYẾT ĐỊNH", LocalDocumentTypeCodes.Decision)]
    [InlineData("CHỈ THỊ", LocalDocumentTypeCodes.Directive)]
    [InlineData("THÔNG TƯ", LocalDocumentTypeCodes.Circular)]
    [InlineData("THÔNG CÁO", LocalDocumentTypeCodes.Communique)]
    [InlineData("THÔNG BÁO", LocalDocumentTypeCodes.Notice)]
    [InlineData("HƯỚNG DẪN", LocalDocumentTypeCodes.Guidance)]
    [InlineData("CHƯƠNG TRÌNH", LocalDocumentTypeCodes.Program)]
    [InlineData("KẾ HOẠCH", LocalDocumentTypeCodes.Plan)]
    [InlineData("PHƯƠNG ÁN", LocalDocumentTypeCodes.Option)]
    [InlineData("ĐỀ ÁN", LocalDocumentTypeCodes.Scheme)]
    [InlineData("DỰ ÁN", LocalDocumentTypeCodes.Project)]
    [InlineData("BÁO CÁO", LocalDocumentTypeCodes.Report)]
    [InlineData("TỜ TRÌNH", LocalDocumentTypeCodes.Proposal)]
    [InlineData("QUY CHẾ", LocalDocumentTypeCodes.Regulation)]
    [InlineData("QUY ĐỊNH", LocalDocumentTypeCodes.Regulation)]
    [InlineData("GIẤY MỜI", LocalDocumentTypeCodes.Invitation)]
    [InlineData("CÔNG ĐIỆN", LocalDocumentTypeCodes.Telegram)]
    [InlineData("GIẤY GIỚI THIỆU", LocalDocumentTypeCodes.IntroductionLetter)]
    [InlineData("BIÊN BẢN", LocalDocumentTypeCodes.Minutes)]
    [InlineData("GIẤY NGHỈ PHÉP", LocalDocumentTypeCodes.LeavePermit)]
    [InlineData("GIẤY ỦY QUYỀN", LocalDocumentTypeCodes.AuthorizationLetter)]
    [InlineData("PHIẾU GỬI", LocalDocumentTypeCodes.SendingSlip)]
    [InlineData("PHIẾU CHUYỂN", LocalDocumentTypeCodes.TransferSlip)]
    [InlineData("PHIẾU BÁO", LocalDocumentTypeCodes.NotificationSlip)]
    [InlineData("KẾT LUẬN", LocalDocumentTypeCodes.Conclusion)]
    public void Resolves_all_named_document_types_from_content(string title, string expectedCode)
    {
        Assert.Equal(expectedCode, new DocumentRoleDetector().ResolveDocumentType(Snapshot(title)));
    }

    [Fact]
    public void Resolves_official_letter_without_a_type_heading_from_subject_marker()
    {
        Assert.Equal(LocalDocumentTypeCodes.OfficialLetter,
            new DocumentRoleDetector().ResolveDocumentType(Snapshot("V/v triển khai nhiệm vụ năm 2026")));
    }

    [Fact]
    public void Content_overrides_a_stale_manual_selection()
    {
        var snapshot = Snapshot("QUYẾT ĐỊNH", LocalDocumentTypeCodes.Report, true);

        Assert.Equal(LocalDocumentTypeCodes.Decision,
            new DocumentRoleDetector().ResolveDocumentType(snapshot));
    }

    [Fact]
    public void Unknown_content_does_not_reuse_non_manual_stale_context()
    {
        var snapshot = Snapshot("Nội dung chưa có tên loại văn bản", LocalDocumentTypeCodes.Decision, false);

        Assert.Equal(LocalDocumentTypeCodes.Unknown,
            new DocumentRoleDetector().ResolveDocumentType(snapshot));
    }

    [Fact]
    public void Accepts_a_small_footnote_marker_after_the_type_heading()
    {
        Assert.Equal(LocalDocumentTypeCodes.Decision,
            new DocumentRoleDetector().ResolveDocumentType(Snapshot("QUYẾT ĐỊNH 1")));
    }

    [Fact]
    public void Repeated_decision_formula_does_not_turn_article_one_into_a_subject()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            Paragraph(2, "Về việc phê duyệt kế hoạch lựa chọn nhà thầu"),
            Paragraph(3, "Căn cứ Luật Đấu thầu số 22/2023/QH15;"),
            Paragraph(4, "QUYẾT ĐỊNH"),
            Paragraph(5, "Điều 1. Phê duyệt kế hoạch lựa chọn nhà thầu.")
        };
        var snapshot = new LocalScanSnapshot("sha256:repeated-decision", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("typeName", roles[1]);
        Assert.Equal("subject", roles[2]);
        Assert.Equal("structuralTitle", roles[4]);
        Assert.False(roles.ContainsKey(5));
    }

    [Fact]
    public void Consecutive_centered_bold_paragraphs_extend_the_subject_until_a_blank_gap()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            Paragraph(2, "Về việc phê duyệt kế hoạch lựa chọn nhà thầu"),
            Paragraph(3, "dự án, dự toán mua sắm: Mua sắm phục vụ ăn bán trú"),
            Paragraph(5, "HIỆU TRƯỞNG TRƯỜNG THCS NGỌC HÀ"),
            Paragraph(6, "Căn cứ Luật Đấu thầu số 22/2023/QH15;")
        };
        var snapshot = new LocalScanSnapshot("sha256:multiline-subject", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("subject", roles[2]);
        Assert.Equal("subjectContinuation", roles[3]);
        Assert.False(roles.ContainsKey(5));
        Assert.Equal("legalBasis", roles[6]);
    }

    [Fact]
    public void Recognizes_legal_basis_only_inside_the_formal_preamble_window()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "BÁO CÁO"),
            Paragraph(2, "Về kết quả thẩm định hồ sơ mời thầu"),
            Paragraph(3, "Căn cứ Nghị định số 30/2020/NĐ-CP;"),
            Paragraph(4, "Theo đề nghị của Tổ thẩm định."),
            Paragraph(5, "1. NỘI DUNG THẨM ĐỊNH"),
            Paragraph(6, "a) Ý kiến thẩm định về cơ sở pháp lý:"),
            Paragraph(7, "Căn cứ các tài liệu được cung cấp, kết quả thẩm định được tổng hợp tại Bảng số 01.")
        };
        var snapshot = new LocalScanSnapshot("sha256:contextual-legal-basis", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("legalBasis", roles[3]);
        Assert.Equal("legalBasis", roles[4]);
        Assert.False(roles.ContainsKey(7));
    }

    [Fact]
    public void Does_not_treat_the_first_căn_cứ_phrase_as_a_component_without_a_formal_title()
    {
        var snapshot = new LocalScanSnapshot("sha256:body-only-căn-cứ", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            new[]
            {
                Paragraph(1, "2. Nội dung của hồ sơ mời thầu"),
                Paragraph(2, "Căn cứ các tài liệu được cung cấp, kết quả được tổng hợp tại Bảng số 02.")
            }, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.False(roles.ContainsKey(2));
    }

    [Fact]
    public void Narrative_căn_cứ_immediately_after_subject_is_not_a_formal_legal_basis()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "BÁO CÁO"),
            Paragraph(2, "Về kết quả thẩm định hồ sơ mời thầu"),
            Paragraph(3, "Căn cứ các tài liệu được cung cấp, kết quả thẩm định được tổng hợp tại Bảng số 01.")
        };
        var snapshot = new LocalScanSnapshot("sha256:narrative-after-subject", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.False(roles.ContainsKey(3));
    }

    [Theory]
    [InlineData("Căn cứ Luật Đấu thầu số 22/2023/QH15;")]
    [InlineData("Căn cứ Hợp đồng số 129/2026/HĐTV/CĐCS-BMC;")]
    [InlineData("Căn cứ chức năng, nhiệm vụ và thẩm quyền được giao;")]
    [InlineData("Theo đề nghị của Tổ thẩm định.")]
    public void Recognizes_supported_formal_legal_basis_sources_after_subject(string legalBasis)
    {
        var paragraphs = new[]
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            Paragraph(2, "Về việc phê duyệt hồ sơ mời thầu"),
            Paragraph(3, legalBasis)
        };
        var snapshot = new LocalScanSnapshot("sha256:formal-source", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("legalBasis", roles[3]);
    }

    [Fact]
    public void Detects_roles_independently_for_two_documents_in_one_word_file()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "CÔNG TY TNHH THỨ NHẤT", page: 1),
            Paragraph(2, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", page: 1),
            Paragraph(3, "Độc lập - Tự do - Hạnh phúc", page: 1),
            Paragraph(4, "Số: 01/QĐ-CT1", page: 1),
            Paragraph(5, "Hà Nội, ngày 01 tháng 09 năm 2026", page: 1),
            Paragraph(6, "QUYẾT ĐỊNH", page: 1),
            Paragraph(7, "Về việc phê duyệt kế hoạch", page: 1),
            Paragraph(8, "Căn cứ Luật Đấu thầu số 22/2023/QH15;", page: 1),
            Paragraph(20, "CÔNG TY TNHH THỨ HAI", page: 2),
            Paragraph(21, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", page: 2),
            Paragraph(22, "Độc lập - Tự do - Hạnh phúc", page: 2),
            Paragraph(23, "Số: 02/TB-CT2", page: 2),
            Paragraph(24, "Hà Nội, ngày 02 tháng 09 năm 2026", page: 2),
            Paragraph(25, "THÔNG BÁO", page: 2),
            Paragraph(26, "Về việc triển khai nhiệm vụ", page: 2),
            Paragraph(27, "Căn cứ Nghị định số 30/2020/NĐ-CP;", page: 2)
        };
        var snapshot = new LocalScanSnapshot("sha256:two-documents", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var detector = new DocumentRoleDetector();
        var blocks = detector.DetectBlocks(snapshot);
        var roles = detector.Detect(snapshot);

        Assert.Equal(2, blocks.Count);
        Assert.Equal(LocalDocumentTypeCodes.Decision, blocks[0].DocumentTypeCode);
        Assert.Equal(LocalDocumentTypeCodes.Notice, blocks[1].DocumentTypeCode);
        Assert.Equal("organName", roles[1]);
        Assert.Equal("typeName", roles[6]);
        Assert.Equal("legalBasis", roles[8]);
        Assert.Equal("organName", roles[20]);
        Assert.Equal("typeName", roles[25]);
        Assert.Equal("subject", roles[26]);
        Assert.Equal("legalBasis", roles[27]);
    }

    [Fact]
    public void Repeated_decision_formula_stays_inside_its_document_block()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", page: 1),
            Paragraph(2, "Độc lập - Tự do - Hạnh phúc", page: 1),
            Paragraph(3, "Số: 01/QĐ-ABC", page: 1),
            Paragraph(4, "QUYẾT ĐỊNH", page: 1),
            Paragraph(5, "Về việc phê duyệt", page: 1),
            Paragraph(6, "Căn cứ Luật Đấu thầu số 22/2023/QH15;", page: 1),
            Paragraph(7, "QUYẾT ĐỊNH", page: 1),
            Paragraph(8, "Điều 1. Phê duyệt kế hoạch.", page: 1)
        };
        var snapshot = new LocalScanSnapshot("sha256:one-document-operative-formula", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var detector = new DocumentRoleDetector();
        var blocks = detector.DetectBlocks(snapshot);
        var roles = detector.Detect(snapshot);

        Assert.Single(blocks);
        Assert.Equal("typeName", roles[4]);
        Assert.Equal("structuralTitle", roles[7]);
    }

    [Theory]
    [InlineData("Số: 129/2026/QĐ-TTĐ.BMC", "QĐ", "Số: 129/QĐ-TTĐ-BMC")]
    [InlineData("Số 5 qđ abc", "QĐ", "Số: 05/QĐ-ABC")]
    public void Normalizes_nd30_code_number_as_one_atomic_component(string source,
        string abbreviation, string expected)
    {
        Assert.Equal(expected,
            LocalAdministrativeTextNormalizer.NormalizeCodeNumber(source, false, abbreviation));
    }

    private static LocalScanSnapshot Snapshot(string text,
        string documentTypeCode = LocalDocumentTypeCodes.Unknown,
        bool selectedManually = false)
    {
        var paragraph = new LocalParagraphSnapshot(1, text, "wdMainTextStory", 1, 0,
            "Times New Roman", fontSizePoints: 14, bold: true, alignment: 1);
        return new LocalScanSnapshot("sha256:document-type", 1,
            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
            new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>(),
            documentTypeCode: documentTypeCode,
            documentTypeWasSelectedManually: selectedManually);
    }

    private static LocalParagraphSnapshot Paragraph(int index, string text, int page = 0) =>
        new(index, text, "wdMainTextStory", 1, index * 100,
            "Times New Roman", fontSizePoints: 14, bold: true, alignment: 1,
            pageNumber: page);
}
