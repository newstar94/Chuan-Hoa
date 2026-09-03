using System.Collections.Generic;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Tests;

public sealed class AnnotationPlannerTests
{
    private const string Fingerprint = "sha256:document";

    [Fact]
    public void Resolves_exact_text_span_and_builds_red_marker_and_two_line_comment()
    {
        var document = Document(Paragraph("MainTextStory", 4, 1, 100, "Nội  dung\r"));
        var finding = Finding(
            "f-1",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 4, 3, 2, "  "));

        var plan = Plan(document, finding);

        var comment = Assert.Single(plan.Comments);
        Assert.Equal(103, comment.Start);
        Assert.Equal(2, comment.Length);
        Assert.Equal("[CHUẨN HÓA:FORMAT:f-1]", comment.Marker);
        Assert.Equal(
            "Hiện tại: Phần này sai định dạng.\nYêu cầu đúng: Dùng Times New Roman cỡ 14.",
            comment.CommentText.Replace("\r\n", "\n"));
        Assert.DoesNotContain("Mã quy tắc", comment.CommentText);
        Assert.DoesNotContain("Mức độ", comment.CommentText);
        Assert.DoesNotContain("Căn cứ:", comment.CommentText);
        var visual = Assert.Single(plan.VisualRanges);
        Assert.Equal(103, visual.Start);
        Assert.Equal(2, visual.Length);
        Assert.Empty(plan.Unresolved);
    }

    [Fact]
    public void Does_not_fall_back_to_another_occurrence_when_offset_is_stale()
    {
        var document = Document(Paragraph("MainTextStory", 1, 1, 0, "sai và sai\r"));
        var finding = Finding(
            "f-stale",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 1, 3, "sai"));

        var plan = Plan(document, finding);

        Assert.Empty(plan.Comments);
        Assert.Empty(plan.VisualRanges);
        Assert.Equal(AnnotationResolutionCode.ExpectedTextMismatch, Assert.Single(plan.Unresolved).Code);
    }

    [Fact]
    public void Exact_offset_selects_the_requested_repeated_occurrence()
    {
        var document = Document(Paragraph("MainTextStory", 1, 1, 20, "sai và sai\r"));
        var finding = Finding(
            "f-second",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 7, 3, "sai"));

        var plan = Plan(document, finding);

        Assert.Equal(27, Assert.Single(plan.Comments).Start);
    }

    [Fact]
    public void Paragraph_level_finding_marks_only_the_paragraph_without_terminator()
    {
        var document = Document(Paragraph("MainTextStory", 2, 1, 50, "Sai cả đoạn\r"));
        var finding = Finding(
            "f-paragraph",
            new AnnotationAnchor(AnnotationAnchorKind.Paragraph, "MainTextStory", 2, null, null, ""));

        var visual = Assert.Single(Plan(document, finding).VisualRanges);

        Assert.Equal(50, visual.Start);
        Assert.Equal("Sai cả đoạn".Length, visual.Length);
    }

    [Fact]
    public void Section_and_document_findings_get_comments_but_no_fake_red_text()
    {
        var document = Document(
            Paragraph("MainTextStory", 1, 1, 0, "Mở đầu\r"),
            Paragraph("MainTextStory", 2, 2, 20, "Section 2\r"));
        var findings = new[]
        {
            Finding("f-section", new AnnotationAnchor(AnnotationAnchorKind.Section, "MainTextStory", null, null, null, "", 2)),
            Finding("f-document", new AnnotationAnchor(AnnotationAnchorKind.Document, "MainTextStory", null, null, null, ""))
        };

        var plan = Plan(document, findings);

        Assert.Equal(2, plan.Comments.Count);
        Assert.Empty(plan.VisualRanges);
        Assert.Equal(20, plan.Comments[0].Start);
        Assert.Equal(0, plan.Comments[1].Start);
    }

    [Fact]
    public void Resolves_table_cell_and_header_story_with_full_coordinates()
    {
        var document = Document(
            new AnnotationParagraphSnapshot("MainTextStory", 7, 1, 80, "Sai trong ô\a", 2, 3, 4),
            Paragraph("PrimaryHeaderStory", 1, 1, 500, "Sai header\r"));
        var findings = new[]
        {
            Finding("f-cell", new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 7, 0, 3, "Sai", 1, 2, 3, 4)),
            Finding("f-header", new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "PrimaryHeaderStory", 1, 4, 6, "header", 1))
        };

        var plan = Plan(document, findings);

        Assert.Equal(2, plan.Comments.Count);
        Assert.Contains(plan.Comments, item => item.StoryType == "MainTextStory" && item.Start == 80);
        Assert.Contains(plan.Comments, item => item.StoryType == "PrimaryHeaderStory" && item.Start == 504);
    }

    [Fact]
    public void Rejects_field_content_control_or_other_protected_span()
    {
        var document = new AnnotationDocumentSnapshot(
            Fingerprint,
            3,
            new[] { Paragraph("MainTextStory", 1, 1, 100, "Vùng field\r") },
            new[] { new AnnotationProtectedSpan("MainTextStory", 100, 4) });
        var finding = Finding(
            "f-protected",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 0, 4, "Vùng"));

        var unresolved = Assert.Single(Plan(document, finding).Unresolved);

        Assert.Equal(AnnotationResolutionCode.ProtectedSpan, unresolved.Code);
    }

    [Fact]
    public void Merges_overlapping_red_ranges_but_keeps_each_explanatory_comment()
    {
        var document = Document(Paragraph("MainTextStory", 1, 1, 0, "abcdef\r"));
        var findings = new[]
        {
            Finding("f-a", new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 1, 3, "bcd")),
            Finding("f-b", new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 3, 3, "def"))
        };

        var plan = Plan(document, findings);

        Assert.Equal(2, plan.Comments.Count);
        var visual = Assert.Single(plan.VisualRanges);
        Assert.Equal(1, visual.Start);
        Assert.Equal(5, visual.Length);
    }

    [Fact]
    public void Deduplicates_same_finding_id_for_idempotent_server_replays()
    {
        var document = Document(Paragraph("MainTextStory", 1, 1, 0, "abc\r"));
        var finding = Finding(
            "same-id",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 0, 1, "a"));

        var plan = Plan(document, finding, finding);

        Assert.Single(plan.Comments);
        Assert.Single(plan.VisualRanges);
    }

    [Fact]
    public void Conflicting_duplicate_finding_id_fails_closed()
    {
        var document = Document(Paragraph("MainTextStory", 1, 1, 0, "abc\r"));
        var first = Finding(
            "same-id",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 0, 1, "a"));
        var second = Finding(
            "same-id",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 1, 1, "b"));

        var plan = Plan(document, first, second);

        Assert.Empty(plan.Comments);
        Assert.Empty(plan.VisualRanges);
        Assert.Equal(AnnotationResolutionCode.ConflictingDuplicateFinding, Assert.Single(plan.Unresolved).Code);
    }

    [Fact]
    public void Ambiguous_paragraph_coordinates_are_not_annotated()
    {
        var document = Document(
            Paragraph("MainTextStory", 1, 1, 0, "abc\r"),
            Paragraph("MainTextStory", 1, 2, 20, "abc\r"));
        var finding = Finding(
            "f-ambiguous",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 0, 1, "a"));

        var plan = Plan(document, finding);

        Assert.Empty(plan.Comments);
        Assert.Equal(AnnotationResolutionCode.ParagraphNotFound, Assert.Single(plan.Unresolved).Code);
    }

    [Fact]
    public void Ownership_policy_never_claims_user_comments_or_another_scan_lane()
    {
        Assert.True(AnnotationOwnershipPolicy.IsOwnedComment("[CHUẨN HÓA:FORMAT:f-1]\nChi tiết", "format"));
        Assert.False(AnnotationOwnershipPolicy.IsOwnedComment("[CHUẨN HÓA:SPELLING:f-1]\nChi tiết", "format"));
        Assert.False(AnnotationOwnershipPolicy.IsOwnedComment("Nhận xét của người dùng", "format"));
    }

    [Fact]
    public void Clearing_restores_only_addin_red_and_preserves_user_changes()
    {
        Assert.Equal(0, AnnotationOwnershipPolicy.ColorAfterClearing(AnnotationOwnershipPolicy.WordRedColor, 0));
        Assert.Equal(0x00FF00, AnnotationOwnershipPolicy.ColorAfterClearing(0x00FF00, 0));
        Assert.Equal(AnnotationOwnershipPolicy.WordRedColor,
            AnnotationOwnershipPolicy.ColorAfterClearing(
                AnnotationOwnershipPolicy.WordRedColor,
                AnnotationOwnershipPolicy.WordRedColor));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Fails_closed_when_fingerprint_or_revision_changed(bool fingerprintChanged)
    {
        var document = Document(Paragraph("MainTextStory", 1, 1, 0, "abc\r"));
        var finding = Finding(
            "f-mismatch",
            new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "MainTextStory", 1, 0, 1, "a"));
        var planner = new AnnotationPlanner();

        var plan = planner.CreatePlan(
            "format",
            "scan-1",
            fingerprintChanged ? "sha256:other" : Fingerprint,
            fingerprintChanged ? 3 : 4,
            document,
            new[] { finding });

        Assert.Empty(plan.Comments);
        Assert.Empty(plan.VisualRanges);
        Assert.Equal(
            fingerprintChanged ? AnnotationResolutionCode.DocumentMismatch : AnnotationResolutionCode.RevisionMismatch,
            Assert.Single(plan.Unresolved).Code);
    }

    private static AnnotationPlan Plan(AnnotationDocumentSnapshot document, params AnnotationFinding[] findings)
    {
        return Plan(document, (IReadOnlyList<AnnotationFinding>)findings);
    }

    private static AnnotationPlan Plan(AnnotationDocumentSnapshot document, IReadOnlyList<AnnotationFinding> findings)
    {
        return new AnnotationPlanner().CreatePlan("format", "scan-1", Fingerprint, 3, document, findings);
    }

    private static AnnotationFinding Finding(string id, AnnotationAnchor anchor)
    {
        return new AnnotationFinding(
            id,
            "RULE-01",
            "Lỗi",
            "Phần này sai định dạng.",
            "Dùng Times New Roman cỡ 14.",
            "NĐ 30/2020/NĐ-CP, Phụ lục I",
            anchor);
    }

    private static AnnotationDocumentSnapshot Document(params AnnotationParagraphSnapshot[] paragraphs)
    {
        return new AnnotationDocumentSnapshot(
            Fingerprint,
            3,
            paragraphs,
            new List<AnnotationProtectedSpan>());
    }

    private static AnnotationParagraphSnapshot Paragraph(
        string storyType,
        int index,
        int section,
        int start,
        string text)
    {
        return new AnnotationParagraphSnapshot(storyType, index, section, start, text);
    }
}
