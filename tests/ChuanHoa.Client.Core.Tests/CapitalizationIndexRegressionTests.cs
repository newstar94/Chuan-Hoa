using System.Reflection;
using System.Text;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Lexicon;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class CapitalizationIndexRegressionTests
{
    private const BindingFlags PrivateStatic = BindingFlags.Static | BindingFlags.NonPublic;

    [Fact]
    public void Indexed_candidates_preserve_exhaustive_culture_matching_and_original_offsets()
    {
        var rules = BundledRulePack.Load("1.0.0.192");
        var samples = rules.Capitalizations.Where((_, i) => i % 173 == 0).Select(r =>
            "Nội dung tại (" + r.Expected.ToLowerInvariant() + "), không phải x" + r.Expected.ToLowerInvariant()).ToList();
        samples.Add("Nội dung tại hà\u00ad nội và thành phố hồ chí minh.");
        samples.Add("Nội dung tại hà\u0000 nội.");
        var paragraphs = samples.Concat(samples.Select(s => s.Normalize(NormalizationForm.FormD)))
            .Select((text, index) => new LocalParagraphSnapshot(index + 1, text, "wdMainTextStory", 1,
                index * 1000, "Times New Roman")).ToArray();
        var actual = new List<AnnotationFinding>();
        typeof(CanonicalRuleScanner).GetMethod("CheckConfiguredCapitalizations", PrivateStatic)!
            .Invoke(null, new object?[] { actual, paragraphs, rules, new Dictionary<int, string>(), CancellationToken.None, null });
        var wholePhrase = typeof(CanonicalRuleScanner).GetMethod("WholePhraseMatches", PrivateStatic)!;
        var expected = new List<string>();
        var categories = new HashSet<string> { "administrative", "geographic", "terrain", "region", "organ", "specialOrgan", "holiday", "lunarYear" };
        foreach (var rule in rules.Capitalizations.Where(r => categories.Contains(r.Category)))
        foreach (var p in paragraphs)
        {
            var matches = (IEnumerable<Tuple<int, int>>)wholePhrase.Invoke(null,
                new object[] { NormalizedTextOffsetMap.Create(p.Text), rule.Expected })!;
            foreach (var match in matches)
            {
                var original = p.Text.Substring(match.Item1, match.Item2);
                if (original != rule.Expected)
                    expected.Add($"{p.Index}|{match.Item1}|{match.Item2}|{original}|Viết “{rule.Expected}”.");
            }
        }
        Assert.Equal(expected.OrderBy(s => s), actual.Select(f =>
            $"{f.Anchor.ParagraphIndex}|{f.Anchor.StartOffset}|{f.Anchor.Length}|{f.Anchor.ExpectedText}|{f.Expected}").OrderBy(s => s));
    }

    [Fact]
    public void Cancellation_is_observed_inside_capitalization_rules()
    {
        using var cancellation = new CancellationTokenSource();
        var calls = 0;
        var method = typeof(CanonicalRuleScanner).GetMethod("CheckConfiguredCapitalizations", PrivateStatic)!;
        var error = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, new object[]
        {
            new List<AnnotationFinding>(), new[] { new LocalParagraphSnapshot(1, "Nội dung kiểm tra.", "wdMainTextStory", 1, 0, "Times New Roman") },
            BundledRulePack.Load("1.0.0.192"), new Dictionary<int, string>(), cancellation.Token,
            (Action)(() => { if (++calls == 10) cancellation.Cancel(); })
        }));
        Assert.IsType<OperationCanceledException>(error.InnerException);
        Assert.Equal(10, calls);
    }

    [Fact]
    public void Spelling_checkpoint_can_cancel_on_the_calling_thread()
    {
        var owner = Environment.CurrentManagedThreadId;
        using var cancellation = new CancellationTokenSource();
        var snapshot = new LocalScanSnapshot("cancel", 1, Array.Empty<LocalSectionSnapshot>(),
            new[] { new LocalParagraphSnapshot(1, "Nội dung kiểm tra.", "wdMainTextStory", 1, 0, "Times New Roman") },
            Array.Empty<AnnotationProtectedSpan>());
        Assert.Throws<OperationCanceledException>(() => new LocalDocumentScanner().ScanSpelling(snapshot,
            BundledRulePack.Load("1.0.0.192"), cancellation.Token, () =>
            {
                Assert.Equal(owner, Environment.CurrentManagedThreadId);
                cancellation.Cancel();
            }));
    }
}
