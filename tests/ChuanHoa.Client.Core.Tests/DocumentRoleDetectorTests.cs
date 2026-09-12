using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class DocumentRoleDetectorTests
{
    [Fact]
    public void Business_basis_continues_formal_preamble_but_not_operative_body()
    {
        var texts = new[] { "QUYẾT ĐỊNH", "Về việc cử nhân viên đi công tác",
            "Căn cứ điều lệ tổ chức và hoạt động của Công ty.",
            "Căn cứ nhu cầu hoạt động kinh doanh;", "Căn cứ nhu cầu và năng lực cán bộ,",
            "QUYẾT ĐỊNH", "Điều 1. Cử nhân viên đi công tác.",
            "Căn cứ nhu cầu thực tế, bố trí công việc." };
        var snapshot = new LocalScanSnapshot("sha256:business-basis", 1,
            Array.Empty<LocalSectionSnapshot>(),
            texts.Select((text, i) => Paragraph(i + 1, text)).ToArray(),
            Array.Empty<AnnotationProtectedSpan>());
        var roles = new DocumentRoleDetector().Detect(snapshot);
        Assert.Equal("legalBasis", roles[3]);
        Assert.Equal("legalBasis", roles[4]);
        Assert.Equal("legalBasis", roles[5]);
        Assert.False(roles.TryGetValue(8, out var role) && role == "legalBasis");
    }

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
    public void Long_subject_without_ve_viec_is_detected_from_adjacency_and_style()
    {
        var longSubject = "Phê duyệt " + new string('x', 330);
        var paragraphs = new[]
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            Paragraph(2, longSubject),
            Paragraph(3, "Căn cứ Bộ luật Lao động số 45/2019/QH14;")
        };
        var snapshot = new LocalScanSnapshot("sha256:long-subject", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("subject", roles[2]);
        Assert.Equal("legalBasis", roles[3]);
    }

    [Fact]
    public void Subject_continuation_can_be_plain_when_its_other_style_signals_match()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            new LocalParagraphSnapshot(2, "Phê duyệt kế hoạch", "wdMainTextStory", 1, 200,
                "Times New Roman", fontSizePoints: 14, bold: false, alignment: 0),
            new LocalParagraphSnapshot(3, "lựa chọn nhà thầu năm 2026", "wdMainTextStory", 1, 300,
                "Times New Roman", fontSizePoints: 14, bold: false, alignment: 0),
            Paragraph(4, "Điều 1. Phê duyệt kế hoạch.")
        };
        var snapshot = new LocalScanSnapshot("sha256:plain-continuation", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("subjectContinuation", roles[3]);
        Assert.False(roles.ContainsKey(4));
    }

    [Fact]
    public void Plain_subject_is_detected_from_position_font_and_following_boundary()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            new LocalParagraphSnapshot(2, "Phê duyệt kế hoạch công tác năm 2026",
                "wdMainTextStory", 1, 200, "Times New Roman", fontSizePoints: 14,
                bold: false, alignment: 0),
            Paragraph(3, "Căn cứ Luật Tổ chức chính quyền địa phương;")
        };
        var snapshot = new LocalScanSnapshot("sha256:plain-subject", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("subject", roles[2]);
        Assert.Equal("legalBasis", roles[3]);
    }

    [Fact]
    public void Plain_body_sentence_after_type_is_not_swallowed_as_subject()
    {
        var paragraphs = new[]
        {
            Paragraph(1, "BÁO CÁO"),
            new LocalParagraphSnapshot(2, "Nội dung báo cáo được tổng hợp từ các đơn vị.",
                "wdMainTextStory", 1, 200, "Times New Roman", fontSizePoints: 13,
                bold: false, alignment: 3),
            new LocalParagraphSnapshot(3, "Kết quả thực hiện nhiệm vụ được trình bày dưới đây.",
                "wdMainTextStory", 1, 300, "Times New Roman", fontSizePoints: 13,
                bold: false, alignment: 3)
        };
        var snapshot = new LocalScanSnapshot("sha256:body-after-type", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.False(roles.TryGetValue(2, out var role) && role == "subject");
        Assert.False(roles.TryGetValue(3, out role) && role == "subjectContinuation");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Legal_basis_window_allows_one_or_two_neutral_bridge_paragraphs(int bridgeCount)
    {
        var paragraphs = new List<LocalParagraphSnapshot>
        {
            Paragraph(1, "QUYẾT ĐỊNH"),
            Paragraph(2, "Về việc phê duyệt kế hoạch"),
            Paragraph(3, "Căn cứ Luật Tổ chức chính quyền địa phương;")
        };
        for (var index = 0; index < bridgeCount; index++)
        {
            paragraphs.Add(new LocalParagraphSnapshot(4 + index, "tiếp tục nội dung viện dẫn của căn cứ trước",
                "wdMainTextStory", 1, 400 + index * 100, "Times New Roman",
                fontSizePoints: 13, bold: false, alignment: 3));
        }
        paragraphs.Add(Paragraph(4 + bridgeCount,
            "Căn cứ Nghị định số 30/2020/NĐ-CP của Chính phủ;"));
        var snapshot = new LocalScanSnapshot("sha256:legal-bridge-" + bridgeCount, 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs, Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("legalBasis", roles[3]);
        Assert.Equal("legalBasis", roles[4 + bridgeCount]);
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

    [Fact]
    public void Separates_report_decision_and_two_combined_header_commitments()
    {
        const string title = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
        const string header = title + "\vĐộc lập - Tự do - Hạnh phúc\v---------------";
        var texts = new[] {
            header, "Hà Nội, ngày 07 tháng 09 năm 2026", "BÁO CÁO THẨM ĐỊNH HỒ SƠ MỜI THẦU", "Nội dung báo cáo",
            title, "Độc lập - Tự do - Hạnh phúc", "Số: 01/QĐ-ABC", "QUYẾT ĐỊNH", "Về việc phê duyệt", "Điều 1. Nội dung",
            header, "Hà Nội, ngày 14 tháng 08 năm 2026", "BẢN CAM KẾT", "Tôi tên là: Nguyễn Văn A", "- Cam kết thực hiện",
            header, "Hà Nội, ngày 14 tháng 08 năm 2026", "BẢN CAM KẾT", "Tôi tên là: Nguyễn Văn B"
        };
        var snapshot = new LocalScanSnapshot("sha256:mixed", 1,
            Array.Empty<LocalSectionSnapshot>(), texts.Select((text, i) => Paragraph(i + 1, text)).ToArray(),
            Array.Empty<AnnotationProtectedSpan>());
        var blocks = new DocumentRoleDetector().DetectBlocks(snapshot);
        Assert.Equal(4, blocks.Count);
        Assert.Equal(LocalDocumentTypeCodes.Report, blocks[0].DocumentTypeCode);
        Assert.Equal(LocalDocumentTypeCodes.Decision, blocks[1].DocumentTypeCode);
        Assert.Equal(LocalDocumentTypeCodes.Unknown, blocks[2].DocumentTypeCode);
        Assert.Equal("standaloneTitle", blocks[2].Roles[13]);
        Assert.False(blocks[2].Roles.ContainsKey(14));
        Assert.Equal("standaloneTitle", blocks[3].Roles[18]);
    }

    private static LocalParagraphSnapshot Paragraph(int index, string text, int page = 0) =>
        new(index, text, "wdMainTextStory", 1, index * 100,
            "Times New Roman", fontSizePoints: 14, bold: true, alignment: 1,
            pageNumber: page);

    [Fact]
    public void Identity_card_label_is_not_an_administrative_document_number()
    {
        var roles = new DocumentRoleDetector().Detect(Snapshot("Số CCCD/Hộ chiếu: 000000000000, cấp ngày 01/01/2026"));
        Assert.False(roles.TryGetValue(1, out var role) && role == "codeNumber");
        Assert.Equal("codeNumber", new DocumentRoleDetector().Detect(Snapshot("Số: 126/QĐ-ABC"))[1]);
    }

    [Fact]
    public void Code_number_can_follow_the_organ_name_inside_the_same_table_cell()
    {
        var paragraph = new LocalParagraphSnapshot(1,
            "ỦY BAN NHÂN DÂN\vSố: 01/QĐ-UBND", "wdMainTextStory", 1, 0,
            "Times New Roman", tableIndex: 1, rowIndex: 1, cellIndex: 1,
            isInTable: true, fontSizePoints: 13);
        var snapshot = new LocalScanSnapshot("sha256:inline-table-number", 1,
            Array.Empty<LocalSectionSnapshot>(), new[] { paragraph },
            Array.Empty<AnnotationProtectedSpan>());

        Assert.Equal("codeNumber", new DocumentRoleDetector().Detect(snapshot)[1]);
    }

    [Fact]
    public void Detects_two_organ_headings_in_left_header_cell_when_code_number_is_absent()
    {
        var paragraphs = new[]
        {
            TableParagraph(1, "BỘ CHỦ QUẢN", 1),
            TableParagraph(2, "CƠ QUAN BAN HÀNH", 1),
            TableParagraph(3, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 2),
            TableParagraph(4, "Độc lập - Tự do - Hạnh phúc", 2),
            Paragraph(5, "QUYẾT ĐỊNH"),
            Paragraph(6, "Về việc phê duyệt kế hoạch"),
            Paragraph(7, "Căn cứ Luật Đấu thầu số 22/2023/QH15;")
        };
        var snapshot = new LocalScanSnapshot("sha256:header-without-code-number", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs,
            Array.Empty<AnnotationProtectedSpan>());

        var roles = new DocumentRoleDetector().Detect(snapshot);

        Assert.Equal("superiorOrganName", roles[1]);
        Assert.Equal("organName", roles[2]);
        Assert.Equal("nationalTitle", roles[3]);
        Assert.Equal("nationalMotto", roles[4]);
        Assert.DoesNotContain(roles, item => item.Value == "codeNumber");
    }

    private static LocalParagraphSnapshot TableParagraph(int index, string text, int column) =>
        new(index, text, "wdMainTextStory", 1, index * 100,
            "Times New Roman", tableIndex: 1, rowIndex: 1, cellIndex: column,
            fontSizePoints: 13, bold: true, alignment: 1, isInTable: true,
            pageNumber: 1);
}
