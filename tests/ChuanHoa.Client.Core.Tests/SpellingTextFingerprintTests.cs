using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class SpellingTextFingerprintTests
{
    [Fact]
    public void FormattingPositionsAndCommentMarksDoNotChangeTheFingerprint()
    {
        var plain = Snapshot("sha256:a", 0, null, "Căn cứ Quyết định số 4869/QĐ-UBND;", "Điều 1. Giao dự toán.");
        // Same words after annotation: other formatting, offsets moved by a comment
        // reference mark (U+0005) and a different raw document fingerprint.
        var annotated = Snapshot("sha256:b", 7, true, "Căn cứ Quyết định số 4869/QĐ-UBND;\u0005", "Điều 1. Giao dự toán.");

        Assert.Equal(SpellingTextFingerprint.Compute(plain), SpellingTextFingerprint.Compute(annotated));
    }

    [Theory]
    [InlineData("Căn cứ Quyết định số 4869/QĐ-UBND.")]
    [InlineData("Căn cứ Quyết  định số 4869/QĐ-UBND;")]
    [InlineData("Căn cứ Quyết định số 4869/QĐ-UBND;\v")]
    public void AnyChangeToTheWordsChangesTheFingerprint(string edited)
    {
        var original = Snapshot("sha256:a", 0, null, "Căn cứ Quyết định số 4869/QĐ-UBND;", "Điều 1. Giao dự toán.");
        var changed = Snapshot("sha256:a", 0, null, edited, "Điều 1. Giao dự toán.");

        Assert.NotEqual(SpellingTextFingerprint.Compute(original), SpellingTextFingerprint.Compute(changed));
    }

    [Fact]
    public void SplittingAParagraphChangesTheFingerprint()
    {
        var joined = Snapshot("sha256:a", 0, null, "của UBND tỉnh Quảng Ninh");
        var split = Snapshot("sha256:a", 0, null, "của UBND tỉnh", "Quảng Ninh");

        Assert.NotEqual(SpellingTextFingerprint.Compute(joined), SpellingTextFingerprint.Compute(split));
    }

    [Fact]
    public void AlreadyCheckedSpellingIsAnEmptyResultBoundToTheSnapshot()
    {
        var snapshot = Snapshot("sha256:checked", 0, null, "Nội dung đã sửa chính tả.");
        var rules = BundledRulePack.Load("1.0.0.194");

        var result = new LocalDocumentScanner().AlreadyCheckedSpelling(snapshot, rules);

        Assert.Equal("spelling", result.Lane);
        Assert.Empty(result.Findings);
        Assert.Equal(snapshot.DocumentFingerprint, result.DocumentFingerprint);
        Assert.Equal(snapshot.Revision, result.Revision);
        Assert.Equal(rules.PackId, result.RulePackId);
    }

    private static LocalScanSnapshot Snapshot(string fingerprint, int offset, bool? bold, params string[] texts)
    {
        var paragraphs = new List<LocalParagraphSnapshot>();
        var start = offset;
        for (var index = 0; index < texts.Length; index++)
        {
            paragraphs.Add(new LocalParagraphSnapshot(index + 1, texts[index], "wdMainTextStory", 1, start,
                "Times New Roman", bold: bold));
            start += texts[index].Length + 1;
        }
        return new LocalScanSnapshot(fingerprint, 1, Array.Empty<LocalSectionSnapshot>(), paragraphs,
            Array.Empty<AnnotationProtectedSpan>());
    }
}
