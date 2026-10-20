using System.Collections.Generic;
using System.Linq;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Tests;

/// <summary>Đợt 2 — D2 comment budget for long documents.</summary>
public sealed class AnnotationBudgetTests
{
    private const string Fingerprint = "sha256:budget";

    [Fact]
    public void Default_plan_groups_findings_of_one_paragraph_into_one_comment()
    {
        var document = Document(Paragraph(1, 0, "Nội dung đoạn một\r"));
        var findings = new[]
        {
            Finding("a", "ND30-PL1-M1-K4-FONT", 1),
            Finding("b", "ND30-PL1-M2-K6E-ALIGN", 1),
            Finding("c", "ND30-PL1-M2-K6E-INDENT", 1)
        };

        var plan = new AnnotationPlanner().CreatePlan("format", "scan", Fingerprint, 1, document,
            findings, AnnotationPlanOptions.Default);

        var comment = Assert.Single(plan.Comments);
        Assert.Equal("a", comment.FindingId);
        Assert.StartsWith("Đoạn này có 3 lỗi:", comment.CommentText);
        Assert.Equal(3, plan.AnnotatedFindingCount);
        Assert.Empty(plan.Omitted);
    }

    [Fact]
    public void Rule_repeated_in_many_paragraphs_gets_one_summary_comment()
    {
        var paragraphs = Enumerable.Range(1, 30)
            .Select(index => Paragraph(index, index * 100, "Đoạn " + index + "\r")).ToArray();
        var findings = Enumerable.Range(1, 30)
            .Select(index => Finding("font-" + index, "ND30-PL1-M1-K4-FONT", index)).ToArray();

        var plan = new AnnotationPlanner().CreatePlan("format", "scan", Fingerprint, 1,
            Document(paragraphs), findings, AnnotationPlanOptions.Default);

        var comment = Assert.Single(plan.Comments);
        Assert.Equal("font-1", comment.FindingId);
        Assert.StartsWith("Lỗi lặp lại ở 30 đoạn", comment.CommentText);
        Assert.Equal(29, plan.Omitted.Count);
    }

    [Fact]
    public void Comments_are_capped_and_the_rest_is_omitted_not_lost()
    {
        var paragraphs = Enumerable.Range(1, 500)
            .Select(index => Paragraph(index, index * 100, "Đoạn " + index + "\r")).ToArray();
        // A different rule per paragraph so nothing is summarized.
        var findings = Enumerable.Range(1, 500)
            .Select(index => Finding("f-" + index, "LOCAL-TYPO-DICT-" + index, index)).ToArray();

        var plan = new AnnotationPlanner().CreatePlan("spelling", "scan", Fingerprint, 1,
            Document(paragraphs), findings, AnnotationPlanOptions.Default);

        Assert.Equal(200, plan.Comments.Count);
        Assert.Equal(300, plan.Omitted.Count);
        Assert.Equal("f-1", plan.Comments[0].FindingId);
        Assert.Equal(500, plan.Comments.Count + plan.Omitted.Count);
        Assert.True(plan.Comments.Zip(plan.Comments.Skip(1), (left, right) => left.Start < right.Start).All(item => item));
    }

    [Fact]
    public void Unlimited_plan_keeps_one_comment_per_finding()
    {
        var document = Document(Paragraph(1, 0, "Nội dung\r"));
        var findings = new[]
        {
            Finding("a", "ND30-PL1-M1-K4-FONT", 1),
            Finding("b", "ND30-PL1-M2-K6E-ALIGN", 1)
        };

        var plan = new AnnotationPlanner().CreatePlan("format", "scan", Fingerprint, 1, document, findings);

        Assert.Equal(2, plan.Comments.Count);
        Assert.Empty(plan.Omitted);
    }

    [Fact]
    public void Protected_span_lookup_matches_linear_scan()
    {
        var document = new AnnotationDocumentSnapshot(Fingerprint, 1,
            new[] { Paragraph(1, 0, "abcdefghij\r"), Paragraph(2, 20, "klmnopqrst\r") },
            new[]
            {
                new AnnotationProtectedSpan("MainTextStory", 2, 50),
                new AnnotationProtectedSpan("MainTextStory", 5, 1)
            });
        var findings = new[] { Finding("a", "ND30-PL1-M1-K4-FONT", 2) };

        var plan = new AnnotationPlanner().CreatePlan("format", "scan", Fingerprint, 1, document,
            findings, AnnotationPlanOptions.Default);

        Assert.Empty(plan.Comments);
        Assert.Equal(AnnotationResolutionCode.ProtectedSpan, Assert.Single(plan.Unresolved).Code);
    }

    [Fact]
    public void Report_lists_every_finding_and_marks_omitted_ones()
    {
        var paragraphs = Enumerable.Range(1, 3)
            .Select(index => Paragraph(index, index * 100, "Đoạn " + index + "\r")).ToArray();
        var findings = Enumerable.Range(1, 3)
            .Select(index => Finding("f-" + index, "RULE-" + index, index)).ToArray();
        var plan = new AnnotationPlanner().CreatePlan("format", "scan", Fingerprint, 1,
            Document(paragraphs), findings, new AnnotationPlanOptions(1, true, 0));

        var html = AnnotationReportBuilder.BuildHtml("A<b>.docx", "format",
            new System.DateTime(2026, 10, 20, 9, 0, 0), findings, plan);

        Assert.Contains("A&lt;b&gt;.docx", html);
        Assert.Contains("RULE-3", html);
        Assert.Equal(2, plan.Omitted.Count);
        Assert.Equal(2, CountOccurrences(html, "class=\"omitted\""));
    }

    private static int CountOccurrences(string text, string token)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(token, index, System.StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += token.Length;
        }
        return count;
    }

    private static AnnotationDocumentSnapshot Document(params AnnotationParagraphSnapshot[] paragraphs) =>
        new(Fingerprint, 1, paragraphs, new List<AnnotationProtectedSpan>());

    private static AnnotationParagraphSnapshot Paragraph(int index, int start, string text) =>
        new("MainTextStory", index, 1, start, text);

    private static AnnotationFinding Finding(string id, string ruleCode, int paragraphIndex) =>
        new(id, ruleCode, "Warning", "Sai định dạng.", "Sửa đúng quy định.", "NĐ30",
            new AnnotationAnchor(AnnotationAnchorKind.Paragraph, "MainTextStory", paragraphIndex,
                null, null, string.Empty), AnnotationSourceFamily.Nd30);
}
