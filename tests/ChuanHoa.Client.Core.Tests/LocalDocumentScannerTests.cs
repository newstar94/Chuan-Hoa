using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class LocalDocumentScannerTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-09-02T00:00:00Z");

    [Fact]
    public void Format_scan_honors_a_pre_cancelled_operation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            new LocalDocumentScanner().ScanFormat(
                Snapshot(ValidSection(), new LocalParagraphSnapshot(
                    1, "Nội dung", "wdMainTextStory", 1, 0, "Times New Roman")),
                Rules(),
                cancellation.Token));
    }

    [Fact]
    public void Spelling_scan_honors_a_pre_cancelled_operation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            new LocalDocumentScanner().ScanSpelling(
                Snapshot(ValidSection(), new LocalParagraphSnapshot(
                    1, "Nội dung", "wdMainTextStory", 1, 0, "Times New Roman")),
                Rules(),
                cancellation.Token));
    }

    [Fact]
    public void Format_scan_returns_exact_font_and_section_findings()
    {
        var snapshot = Snapshot(new LocalSectionSnapshot(1, 700, 900, 10, 10, 10, 10, true),
            new LocalParagraphSnapshot(1, "Nội dung", "wdMainTextStory", 1, 0, "Arial"));

        var result = new LocalDocumentScanner().ScanFormat(snapshot, Rules());

        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M1-K1" && item.Anchor.Kind == AnnotationAnchorKind.Section);
        Assert.DoesNotContain(result.Findings, item => item.RuleCode == "ND30-PL1-M1-K2");
        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M1-K3");
    }

    [Fact]
    public void Spelling_scan_has_exact_anchor_and_does_not_include_removed_tone_or_iy_rules()
    {
        var paragraph = new LocalParagraphSnapshot(1, "Đơn vị sát nhập  và viết sai , rồi.\u200B", "wdMainTextStory", 1, 0, "Times New Roman");

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), Rules());

        var typo = Assert.Single(result.Findings, item => item.RuleCode == "LOCAL-TYPO-DICT");
        Assert.Equal("sát nhập", typo.Anchor.ExpectedText);
        Assert.Equal(paragraph.Text.IndexOf("sát nhập", StringComparison.Ordinal), typo.Anchor.StartOffset);
        Assert.Contains(result.Findings, item => item.RuleCode == "LOCAL-TYPO-SPACE");
        Assert.Contains(result.Findings, item => item.RuleCode == "LOCAL-TYPO-PUNCT");
        Assert.Contains(result.Findings, item => item.RuleCode == "LOCAL-TYPO-HIDDEN");
        Assert.DoesNotContain(result.Findings, item => item.RuleCode.Contains("TONE", StringComparison.Ordinal) || item.RuleCode.Contains("IY", StringComparison.Ordinal));
    }

    [Fact]
    public void Spelling_scan_detects_contextual_confusion_pair_ban_giao_tai_nieu()
    {
        var paragraph = new LocalParagraphSnapshot(1, "Bàn giao tài niệu", "wdMainTextStory", 1, 0, "Times New Roman");
        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), Rules());

        var finding = Assert.Single(result.Findings, item => item.RuleCode == "LOCAL-TYPO-DICT");
        Assert.Equal("tài niệu", finding.Anchor.ExpectedText);
        Assert.Contains("tài liệu", finding.Expected);
    }

    [Fact]
    public void Spelling_scan_ignores_words_in_personal_dictionary()
    {
        const string customWord = "chuyengiarieng";
        ChuanHoa.Client.Core.Lexicon.PersonalDictionaryManager.Instance.AddUserWord(customWord);
        try
        {
            var paragraph = new LocalParagraphSnapshot(1, "Báo cáo " + customWord + " đầy đủ.", "wdMainTextStory", 1, 0, "Times New Roman");
            var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), Rules());
            Assert.DoesNotContain(result.Findings, item => item.Anchor.ExpectedText == customWord);
        }
        finally
        {
            ChuanHoa.Client.Core.Lexicon.PersonalDictionaryManager.Instance.RemoveUserWord(customWord);
        }
    }

    [Fact]
    public void Spelling_scan_honors_document_scoped_ignore_using_snapshot_fingerprint()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "ChuanHoaTestDict_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            var manager = new ChuanHoa.Client.Core.Lexicon.PersonalDictionaryManager(tempDir);
            const string fingerprint = "sha256:document-a";
            Assert.True(manager.IgnoreWordForDocument(fingerprint, "nhậnn").Succeeded);
            var paragraph = new LocalParagraphSnapshot(1, "Văn bản nhậnn.", "wdMainTextStory", 1, 0,
                "Times New Roman");
            var snapshot = new LocalScanSnapshot(fingerprint, 1, new[] { ValidSection() }, new[] { paragraph },
                Array.Empty<AnnotationProtectedSpan>());
            var rules = new LocalRulePack("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
                210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
                Array.Empty<TextCorrectionRule>(), Array.Empty<TelexRule>(), Array.Empty<char>(),
                lexicon: new[] { "văn", "bản", "nhận" });

            var findings = new LocalDocumentScanner(manager).ScanSpelling(snapshot, rules).Findings;
            Assert.DoesNotContain(findings, item => item.Anchor.ExpectedText == "nhậnn");

            var otherSnapshot = new LocalScanSnapshot("sha256:document-b", 1,
                new[] { ValidSection() }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());
            Assert.Contains(new LocalDocumentScanner(manager).ScanSpelling(otherSnapshot, rules).Findings,
                item => item.Anchor.ExpectedText == "nhậnn");
        }
        finally
        {
            if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void Format_scan_checks_component_body_code_number_and_place_date_locally()
    {
        var paragraphs = new[]
        {
            new LocalParagraphSnapshot(1, "Cộng hòa xã hội chủ nghĩa Việt Nam", "wdMainTextStory", 1, 0,
                "Arial", fontSizePoints: 11, bold: false, italic: true, alignment: 0),
            new LocalParagraphSnapshot(2, "Đây là một đoạn nội dung đủ dài để được nhận diện là nội dung chính của văn bản.",
                "wdMainTextStory", 1, 40, "Times New Roman", fontSizePoints: 11, alignment: 0,
                firstLineIndentPoints: 0, spaceAfterPoints: 0),
            new LocalParagraphSnapshot(3, "Số: 5/qđ-abc", "wdMainTextStory", 1, 130, "Times New Roman"),
            new LocalParagraphSnapshot(4, "Hà Nội ngày 5 tháng 2 năm 2026", "wdMainTextStory", 1, 150, "Times New Roman")
        };
        var snapshot = new LocalScanSnapshot("sha256:test", 1, new[] { ValidSection() }, paragraphs,
            Array.Empty<AnnotationProtectedSpan>());

        var result = new LocalDocumentScanner().ScanFormat(snapshot, Rules());

        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M2-K1-QH");
        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M2-K6E-ALIGN");
        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M2-K3-PAD" && item.Anchor.ExpectedText == "5");
        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M2-K3-CASE");
        Assert.Contains(result.Findings, item => item.RuleCode == "ND30-PL1-M2-K4-COMMA");
        Assert.Equal(2, result.Findings.Count(item => item.RuleCode == "ND30-PL1-M2-K4-PAD"));
    }

    [Fact]
    public void Spelling_scan_checks_sentence_capitalization_abbreviations_and_safe_telex_pattern()
    {
        var paragraph = new LocalParagraphSnapshot(1, "nội dung đầu. câu sau doof. kt. GIÁM ĐỐC", "wdMainTextStory", 1, 0, "Times New Roman");
        var rules = new LocalRulePack("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
            210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
            Array.Empty<TextCorrectionRule>(),
            new[] { new TelexRule(@"\b(?=[A-Za-z]*[sfrxj]\b)(?=[A-Za-z]*(?:aw|aa|dd|ee|oo|ow|uw))[A-Za-z]{3,}\b", "từ tiếng Việt tương ứng") },
            Array.Empty<char>());

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), rules);

        Assert.Equal(2, result.Findings.Count(item => item.RuleCode == "ND30-PL2-M1"));
        Assert.Contains(result.Findings, item => item.RuleCode == "LOCAL-TYPO-TELEX" && item.Anchor.ExpectedText == "doof");
    }

    [Fact]
    public void Spelling_scan_combines_context_corrections_with_unknown_word_detection()
    {
        var paragraph = new LocalParagraphSnapshot(1,
            "Quyết định xố, ự án và nhậnn.", "wdMainTextStory", 1, 0, "Times New Roman");
        var rules = new LocalRulePack("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
            210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
            new[]
            {
                new TextCorrectionRule("quyết định xố", "quyết định số"),
                new TextCorrectionRule("ự án", "dự án")
            },
            Array.Empty<TelexRule>(), Array.Empty<char>(),
            lexicon: new[] { "quyết", "định", "số", "dự", "án", "và", "nhận" });

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), rules);

        Assert.Equal(2, result.Findings.Count(item => item.RuleCode == "LOCAL-TYPO-DICT"));
        var unknown = Assert.Single(result.Findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON");
        Assert.Equal("nhậnn", unknown.Anchor.ExpectedText);
        Assert.Contains("“nhận”", unknown.Expected, StringComparison.Ordinal);
    }

    [Fact]
    public void Lexicon_only_returns_an_unambiguous_edit_distance_correction()
    {
        var checker = new VietnameseLexiconSpellChecker(new[] { "nhận", "dự", "tự", "án", "số", "sở" });

        Assert.Equal("nhận", checker.FindDeterministicCorrection("nhậnn"));
        Assert.Null(checker.FindDeterministicCorrection("ự"));
        Assert.True(checker.IsKnown("Dự"));
    }

    [Theory]
    [InlineData("a. Quy mô đầu tư:")]
    [InlineData("b. Giải pháp thiết kế:")]
    [InlineData("a) Nội dung")]
    [InlineData("1. Nội dung")]
    [InlineData("1) Nội dung")]
    [InlineData("(1) Nội dung")]
    [InlineData("- Nội dung")]
    [InlineData("– Nội dung")]
    [InlineData("+ Nội dung")]
    [InlineData("• Nội dung")]
    [InlineData("1. a) Nội dung")]
    public void Spelling_scan_accepts_lowercase_list_markers_when_content_starts_with_uppercase(string text)
    {
        var paragraph = new LocalParagraphSnapshot(1, text, "wdMainTextStory", 1, 0, "Times New Roman");

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), Rules());

        Assert.DoesNotContain(result.Findings, item => item.RuleCode == "ND30-PL2-M1");
    }

    [Theory]
    [InlineData("a. quy mô đầu tư:", "q")]
    [InlineData("b) giải pháp thiết kế:", "g")]
    [InlineData("1. nội dung", "n")]
    [InlineData("(1) nội dung", "n")]
    [InlineData("- nội dung", "n")]
    [InlineData("1. a) nội dung", "n")]
    public void Spelling_scan_reports_content_after_list_marker_once_and_anchors_the_content_letter(string text, string expectedLetter)
    {
        var paragraph = new LocalParagraphSnapshot(1, text, "wdMainTextStory", 1, 0, "Times New Roman");

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), Rules());

        var finding = Assert.Single(result.Findings, item => item.RuleCode == "ND30-PL2-M1");
        Assert.Equal(expectedLetter, finding.Anchor.ExpectedText);
        Assert.Equal(text.IndexOf(expectedLetter, StringComparison.Ordinal), finding.Anchor.StartOffset);
        Assert.Equal("Chữ cái đầu nội dung sau ký hiệu chỉ mục chưa viết hoa.", finding.CurrentIssue);
        Assert.Equal("Viết hoa chữ cái đầu nội dung sau ký hiệu chỉ mục.", finding.Expected);
    }

    [Theory]
    [InlineData("a) Nội dung")]
    [InlineData("b) Giải pháp")]
    [InlineData("c) Cách thức")]
    [InlineData("a. Nội dung")]
    [InlineData("1. Nội dung")]
    [InlineData("(1) Nội dung")]
    [InlineData("1. a) Nội dung")]
    public void Spelling_lexicon_ignores_leading_list_marker_but_checks_the_content(string text)
    {
        var paragraph = new LocalParagraphSnapshot(1, text, "wdMainTextStory", 1, 0, "Times New Roman");
        var rules = new LocalRulePack("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
            210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
            Array.Empty<TextCorrectionRule>(), Array.Empty<TelexRule>(), Array.Empty<char>(),
            lexicon: new[] { "nội", "dung", "giải", "pháp", "cách", "thức" });

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), rules);

        Assert.DoesNotContain(result.Findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            (item.Anchor.ExpectedText == "a" || item.Anchor.ExpectedText == "b" || item.Anchor.ExpectedText == "c"));
        Assert.DoesNotContain(result.Findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON");
    }

    [Fact]
    public void Spelling_lexicon_still_checks_unknown_content_after_a_list_marker()
    {
        var paragraph = new LocalParagraphSnapshot(1, "b) nhậnn", "wdMainTextStory", 1, 0, "Times New Roman");
        var rules = new LocalRulePack("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
            210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
            Array.Empty<TextCorrectionRule>(), Array.Empty<TelexRule>(), Array.Empty<char>(),
            lexicon: new[] { "nhận" });

        var result = new LocalDocumentScanner().ScanSpelling(Snapshot(ValidSection(), paragraph), rules);

        Assert.DoesNotContain(result.Findings, item => item.Anchor.ExpectedText == "b");
        Assert.Contains(result.Findings, item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
            item.Anchor.ExpectedText == "nhậnn");
    }

    private static LocalScanSnapshot Snapshot(LocalSectionSnapshot section, LocalParagraphSnapshot paragraph) =>
        new("sha256:test", 1, new[] { section }, new[] { paragraph }, Array.Empty<AnnotationProtectedSpan>());

    private static LocalSectionSnapshot ValidSection()
    {
        const double point = 72d / 25.4d;
        return new LocalSectionSnapshot(1, 210 * point, 297 * point, 20 * point, 20 * point, 30 * point, 15 * point, false);
    }

    private static LocalRulePack Rules() => new("TEST", "1.0.0", Now.AddDays(-1), Now.AddDays(30), "1.0.0.0",
        210, 297, 20, 25, 20, 25, 30, 35, 15, 20, "Times New Roman",
        new[] { new TextCorrectionRule("sát nhập", "sáp nhập") }, Array.Empty<TelexRule>(), new[] { '\u200B' });
}
