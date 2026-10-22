using System;
using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Lexicon;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class ContextualSpellingTests
{
    private static readonly LocalRulePack Rules = BundledRulePack.Load("1.0.0.178");

    [Theory]
    // Units after a number, in a compound unit or with an exponent.
    [InlineData("+ Vi sinh vật cố định đạm: 3,12 x 106 CFU/g;", "x,g")]
    [InlineData("+ Khối lượng tịnh: 25kg/bao.", "kg")]
    [InlineData("Đơn giá 15.000 đ/kg, giao trong 8 h kể từ khi nhận đơn.", "đ,h")]
    [InlineData("Diện tích sàn 120 m2, thể tích bể 30 m³.", "m")]
    [InlineData("Độ pH của đất phải đạt từ 6 đến 7.", "pH")]
    // Products and dimensions, also between the dimension names.
    [InlineData("Kích thước (dài x rộng x cao): 20 x 30 x 40 cm.", "x")]
    [InlineData("Trọng lượng được tính theo công thức P = m x g.", "m,x,g")]
    // Lettered parts of a legal or document structure.
    [InlineData("Trường hợp bên A chấm dứt hợp đồng theo điểm b khoản 1 Điều này.", "b")]
    [InlineData("Áp dụng điểm b, c và đ khoản 2 Điều 5 của Luật.", "b,c,đ")]
    [InlineData("Thực hiện theo mẫu tại phụ lục c kèm theo.", "c")]
    // Inline list markers, also lower-case roman numerals.
    [InlineData("Hồ sơ gồm: a) Đơn đề nghị; b) Bản sao giấy phép.", "b")]
    [InlineData("Các quyền gồm: (i) quyền sở hữu; (ii) quyền sử dụng; (iii) quyền định đoạt.", "ii,iii")]
    [InlineData("d.Bước 4: Bố trí nhân lực:", "d")]
    [InlineData("b . Không gây cản trở giao thông:", "b")]
    [InlineData("B Giằng thép theo thiết kế.", "B")]
    // Slash, dotted and labelled abbreviations.
    [InlineData("V/v đề nghị hỗ trợ kinh phí", "V,v")]
    [InlineData("Kính gửi đ/c Nguyễn Văn Bình.", "đ,c")]
    [InlineData("Chi tiết xem tr. 15 của báo cáo, các mục khác v.v.", "tr,v")]
    [InlineData("Sđt: 0912345678.", "Sđt")]
    [InlineData("Vd: thôn Đông, xã Long Cốc.", "Vd")]
    public void Lexicon_accepts_tokens_that_their_context_makes_correct(string text, string tokens)
    {
        var findings = ScanSpelling(Paragraph(1, text));

        foreach (var token in tokens.Split(','))
            Assert.DoesNotContain(findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
                item.Anchor.ExpectedText == token);
    }

    [Theory]
    [InlineData("Nội dung b cần được xem xét lại.", "b")]
    [InlineData("Chúng tôi x nhận hồ sơ của ông.", "x")]
    [InlineData("Gửi kèm g cho đơn vị.", "g")]
    [InlineData("Thời hạn là 3 ngàyy làm việc.", "ngàyy")]
    [InlineData("Đơn giá tính theo pcs.", "pcs")]
    public void Lexicon_still_reports_the_same_tokens_outside_such_a_context(string text, string token)
    {
        var findings = ScanSpelling(Paragraph(1, text));

        Assert.Contains(findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            item.Anchor.ExpectedText == token);
    }

    [Fact]
    public void A_unit_column_and_the_units_it_establishes_are_accepted_across_the_document()
    {
        var paragraphs = new[]
        {
            Cell(1, "Tên hàng hóa", 1, 1), Cell(2, "Đơn vị tính", 1, 2), Cell(3, "Số lượng", 1, 3),
            Cell(4, "Ống nhựa", 2, 1), Cell(5, "pcs", 2, 2), Cell(6, "120", 2, 3),
            Cell(7, "Phân bón", 3, 1), Cell(8, "Kg", 3, 2), Cell(9, "99.600", 3, 3),
            Paragraph(10, "Đơn giá tính theo pcs.")
        };

        var findings = ScanSpelling(paragraphs);

        Assert.DoesNotContain(findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            (item.Anchor.ExpectedText == "pcs" || item.Anchor.ExpectedText == "Kg"));
    }

    [Fact]
    public void A_unit_column_is_recognised_by_its_header_not_by_its_position()
    {
        var paragraphs = new[]
        {
            Cell(1, "Tên hàng hóa", 1, 1), Cell(2, "Ghi chú", 1, 2),
            Cell(3, "Ống nhựa", 2, 1), Cell(4, "pcs", 2, 2)
        };

        var findings = ScanSpelling(paragraphs);

        Assert.Contains(findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            item.Anchor.ExpectedText == "pcs");
    }

    [Fact]
    public void A_foreign_term_the_document_repeats_is_accepted_but_a_one_off_or_unaccented_word_is_not()
    {
        var findings = ScanSpelling(
            Paragraph(1, "Lan can bằng inox, tay vịn inox."),
            Paragraph(2, "Cổng inox sơn tĩnh điện, ốp đá granite."),
            Paragraph(3, "Hồ sơ khong đầy đủ, nhà thầu khong bổ sung, chủ đầu tư khong chấp nhận."));

        Assert.DoesNotContain(findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            item.Anchor.ExpectedText == "inox");
        Assert.Contains(findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            item.Anchor.ExpectedText == "granite");
        Assert.Equal(3, findings.Count(item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            item.Anchor.ExpectedText == "khong"));
    }

    [Theory]
    [InlineData("3,12 x 106", 2, ContextualTokenKind.Multiplication)]
    [InlineData("theo điểm b khoản 1", 4, ContextualTokenKind.StructuralLabel)]
    [InlineData("gồm (ii) quyền", 3, ContextualTokenKind.ListMarker)]
    [InlineData("Kính gửi đ/c Bình", 6, ContextualTokenKind.Abbreviation)]
    [InlineData("25 kg", 2, ContextualTokenKind.MeasurementUnit)]
    [InlineData("khổ A4", 2, ContextualTokenKind.CodeOrFormula)]
    [InlineData("nội dung b cần", 4, ContextualTokenKind.None)]
    public void Classifier_names_the_context_that_makes_a_token_correct(string text, int tokenIndex,
        ContextualTokenKind expected)
    {
        var tokens = VietnameseWordTokenizer.TokenizeParagraph(text);

        Assert.Equal(expected, ContextualTokenClassifier.Classify(tokens, tokenIndex));
    }

    private static IReadOnlyList<AnnotationFinding> ScanSpelling(params LocalParagraphSnapshot[] paragraphs) =>
        new LocalDocumentScanner().ScanSpelling(new LocalScanSnapshot("sha256:context", 1,
            new[] { ValidSection() }, paragraphs, Array.Empty<AnnotationProtectedSpan>()), Rules).Findings;

    private static LocalParagraphSnapshot Paragraph(int index, string text) =>
        new(index, text, "wdMainTextStory", 1, index * 200, "Times New Roman");

    private static LocalParagraphSnapshot Cell(int index, string text, int row, int cell) =>
        new(index, text, "wdMainTextStory", 1, index * 200, "Times New Roman",
            tableIndex: 1, rowIndex: row, cellIndex: cell, isInTable: true);

    private static LocalSectionSnapshot ValidSection()
    {
        const double point = 72d / 25.4d;
        return new LocalSectionSnapshot(1, 210 * point, 297 * point, 20 * point, 20 * point, 30 * point, 15 * point, false);
    }
}
