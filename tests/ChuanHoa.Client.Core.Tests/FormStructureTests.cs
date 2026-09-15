using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Annotations;

namespace ChuanHoa.Client.Core.Tests;

public sealed class FormStructureTests
{
    [Theory]
    [InlineData("ĐƠN XIN NGHỈ VIỆC")]
    [InlineData("ĐƠN ĐỀ NGHỊ HỖ TRỢ")]
    [InlineData("GIẤY CAM KẾT")]
    [InlineData("BẢN ĐĂNG KÝ THAM GIA")]
    public void Recognizes_form_title_without_turning_following_text_into_subject(string title)
    {
        var roles = new DocumentRoleDetector().Detect(Snapshot(title, "Tôi tên là:...................."));
        Assert.Equal("standaloneTitle", roles[1]);
        Assert.False(roles.TryGetValue(2, out var role) && role == "subject");
    }

    [Fact]
    public void Inline_first_recipient_can_continue_with_automatic_list()
    {
        var paragraphs = new[] {
            new LocalParagraphSnapshot(1, "Kính gửi: Ban giám đốc", "wdMainTextStory", 1, 0, "Times New Roman"),
            new LocalParagraphSnapshot(2, "Phòng nhân sự", "wdMainTextStory", 1, 50, "Times New Roman", listMarker:"-"),
            new LocalParagraphSnapshot(3, "Tôi tên là:....................", "wdMainTextStory", 1, 100, "Times New Roman") };
        var roles = new DocumentRoleDetector().Detect(new LocalScanSnapshot("form", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs, Array.Empty<AnnotationProtectedSpan>()));
        Assert.Equal("recipientSalutationList", roles[2]);
        Assert.False(roles.ContainsKey(3));
    }

    [Fact]
    public void Ordinary_sentence_is_not_a_form_title()
    {
        Assert.False(new DocumentRoleDetector().Detect(Snapshot("Đơn xin nghỉ việc của tôi đã được chấp thuận.")).ContainsKey(1));
    }

    private static LocalScanSnapshot Snapshot(params string[] texts) => new("form", 1,
        Array.Empty<LocalSectionSnapshot>(), texts.Select((text,i) => new LocalParagraphSnapshot(
            i+1, text, "wdMainTextStory", 1, i*100, "Times New Roman")).ToArray(), Array.Empty<AnnotationProtectedSpan>());
}
