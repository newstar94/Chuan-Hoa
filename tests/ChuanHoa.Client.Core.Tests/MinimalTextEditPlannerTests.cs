using ChuanHoa.Client.Core.Text;

namespace ChuanHoa.Client.Core.Tests;

public sealed class MinimalTextEditPlannerTests
{
    [Fact]
    public void IdenticalTextNeedsNoEdit()
    {
        Assert.Empty(MinimalTextEditPlanner.Plan("Không đổi.", "Không đổi.")!);
    }

    [Fact]
    public void RepeatedSpacesAreRemovedWithoutTouchingWords()
    {
        var edits = MinimalTextEditPlanner.Plan("Văn   bản", "Văn bản")!;

        // The first space is kept; the two after it are deleted.
        var edit = Assert.Single(edits);
        Assert.Equal(4, edit.Offset);
        Assert.Equal(2, edit.Length);
        Assert.Equal("  ", edit.ExpectedText);
        Assert.Equal(string.Empty, edit.Replacement);
    }

    [Fact]
    public void MissingSpaceAfterPunctuationIsAnInsertion()
    {
        var edit = Assert.Single(MinimalTextEditPlanner.Plan("a,b", "a, b")!);

        Assert.Equal(2, edit.Offset);
        Assert.Equal(0, edit.Length);
        Assert.Equal(" ", edit.Replacement);
    }

    [Fact]
    public void SingleCharacterSubstitutionKeepsItsPosition()
    {
        var edit = Assert.Single(MinimalTextEditPlanner.Plan("ngày 15", "ngày 15")!);

        Assert.Equal(4, edit.Offset);
        Assert.Equal(1, edit.Length);
        Assert.Equal(" ", edit.ExpectedText);
        Assert.Equal(" ", edit.Replacement);
    }

    [Theory]
    [InlineData("Văn bản  có từ ĐẬM và  nghiêng ,ở giữa.")]
    [InlineData("Đây  là   văn   bản.")]
    [InlineData("Hà Nội , ngày 15 tháng 8 .")]
    [InlineData("Công tác:chỉ đạo;điều hành")]
    [InlineData("Nội dung ( trong ngoặc ) và [ ghi chú ]")]
    [InlineData("   Kính gửi:   Ủy ban nhân dân   ")]
    [InlineData("Báo cáo \" kết quả thực hiện \" năm 2026")]
    [InlineData("Kính gửi PGS.TS.Nguyễn Văn A và ThS.Trần Thị B.")]
    [InlineData("Căn cứ Quyết định  số 4869/QĐ-UBND ngày 22/12/2025 ,về việc giao dự toán.")]
    public void TypographyCleanupEditsOnlyWhitespacePunctuationAndQuotes(string source)
    {
        var target = VietnameseTypographyCleaner.NormalizeQuotationMarks(
            VietnameseTypographyCleaner.CleanWhitespaceAndPunctuation(source));

        var edits = MinimalTextEditPlanner.Plan(source, target)!;

        Assert.NotEmpty(edits);
        Assert.Equal(target, MinimalTextEditPlanner.ApplyTo(source, edits));
        // Letters and digits are never rewritten, so their character formatting
        // (bold, italic, colour) survives when the edits are applied in Word.
        foreach (var edit in edits)
        {
            Assert.DoesNotContain(edit.ExpectedText, char.IsLetterOrDigit);
            Assert.DoesNotContain(edit.Replacement, char.IsLetterOrDigit);
        }
    }

    [Fact]
    public void EditsAreOrderedDisjointAndMatchTheSource()
    {
        var random = new Random(20261108);
        const string alphabet = "ab  ,. \"“”đ";
        for (var iteration = 0; iteration < 3000; iteration++)
        {
            var source = RandomText(random, alphabet, random.Next(0, 40));
            var target = random.Next(4) == 0
                ? RandomText(random, alphabet, random.Next(0, 40))
                : Mutate(random, source, alphabet);

            var edits = MinimalTextEditPlanner.Plan(source, target);

            Assert.NotNull(edits);
            Assert.Equal(target, MinimalTextEditPlanner.ApplyTo(source, edits!));
            var previousEnd = -1;
            foreach (var edit in edits!)
            {
                Assert.True(edit.Offset > previousEnd, $"Edits overlap or touch for '{source}' -> '{target}'.");
                Assert.Equal(source.Substring(edit.Offset, edit.Length), edit.ExpectedText);
                Assert.True(edit.Length > 0 || edit.Replacement.Length > 0);
                previousEnd = edit.Offset + edit.Length;
            }
        }
    }

    [Fact]
    public void EditDistanceAboveTheLimitIsNotPlanned()
    {
        Assert.Null(MinimalTextEditPlanner.Plan("aaaaaaaa", "bbbbbbbb", maximumEditDistance: 4));
        Assert.NotNull(MinimalTextEditPlanner.Plan("aaaaaaaa", "aaaaaaab", maximumEditDistance: 4));
    }

    private static string RandomText(Random random, string alphabet, int length)
    {
        var characters = new char[length];
        for (var index = 0; index < length; index++)
            characters[index] = alphabet[random.Next(alphabet.Length)];
        return new string(characters);
    }

    private static string Mutate(Random random, string source, string alphabet)
    {
        var builder = new System.Text.StringBuilder(source);
        var changes = random.Next(1, 6);
        for (var change = 0; change < changes; change++)
        {
            var position = random.Next(builder.Length + 1);
            switch (random.Next(3))
            {
                case 0:
                    builder.Insert(position, alphabet[random.Next(alphabet.Length)]);
                    break;
                case 1 when position < builder.Length:
                    builder.Remove(position, 1);
                    break;
                default:
                    if (position < builder.Length) builder[position] = alphabet[random.Next(alphabet.Length)];
                    break;
            }
        }
        return builder.ToString();
    }
}
