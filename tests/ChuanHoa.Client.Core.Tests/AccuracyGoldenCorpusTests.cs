using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class AccuracyGoldenCorpusTests
{
    private sealed record RoleCase(string Id, string[] Input,
        IReadOnlyDictionary<int, string> ExpectedRoles,
        IReadOnlyDictionary<int, string> ForbiddenRoles,
        string DocumentType, string Regime, string ReasonSource);

    [Fact]
    public void Synthetic_public_safe_role_corpus_has_no_role_misses_or_forbidden_assignments()
    {
        var cases = new[]
        {
            new RoleCase("decision",
                new[] { "QUYẾT ĐỊNH", "Phê duyệt kế hoạch lựa chọn nhà thầu",
                    "Căn cứ Luật Đấu thầu số 22/2023/QH15;", "Điều 1. Phê duyệt kế hoạch." },
                new Dictionary<int, string> { [1] = "typeName", [2] = "subject", [3] = "legalBasis" },
                new Dictionary<int, string> { [4] = "subjectContinuation" },
                LocalDocumentTypeCodes.Decision, "ND30", "Synthetic NĐ30 decision"),
            new RoleCase("official-letter",
                new[] { "Số: 10/UBND-VP", "V/v triển khai nhiệm vụ", "Kính gửi: Bộ Nội vụ" },
                new Dictionary<int, string> { [1] = "codeNumber", [2] = "officialLetterSubject" },
                new Dictionary<int, string> { [3] = "legalBasis" },
                LocalDocumentTypeCodes.OfficialLetter, "ND30", "Synthetic NĐ30 official letter"),
            new RoleCase("party-document",
                new[] { "ĐẢNG CỘNG SẢN VIỆT NAM", "QUYẾT ĐỊNH", "Ban hành quy chế làm việc",
                    "- Căn cứ Điều lệ Đảng;", "Điều 1. Ban hành quy chế." },
                new Dictionary<int, string> { [1] = "partyTitle", [2] = "typeName", [3] = "subject", [4] = "legalBasis" },
                new Dictionary<int, string> { [5] = "subjectContinuation" },
                LocalDocumentTypeCodes.Decision, "PARTY_HD05", "Synthetic HĐ05 decision")
        };

        var falseNegatives = 0;
        var falsePositives = 0;
        foreach (var item in cases)
        {
            Assert.False(string.IsNullOrWhiteSpace(item.ReasonSource));
            var paragraphs = item.Input.Select((text, index) =>
                new LocalParagraphSnapshot(index + 1, text, "wdMainTextStory", 1,
                    index * 100, "Times New Roman", fontSizePoints: 14,
                    bold: true, alignment: 1)).ToArray();
            var snapshot = new LocalScanSnapshot("golden:" + item.Id, 1,
                Array.Empty<LocalSectionSnapshot>(), paragraphs,
                Array.Empty<ChuanHoa.Client.Core.Annotations.AnnotationProtectedSpan>(),
                regimeCode: item.Regime);
            var detector = new DocumentRoleDetector();
            var roles = detector.Detect(snapshot);
            falseNegatives += item.ExpectedRoles.Count(expected =>
                !roles.TryGetValue(expected.Key, out var actual) || actual != expected.Value);
            falsePositives += item.ForbiddenRoles.Count(forbidden =>
                roles.TryGetValue(forbidden.Key, out var actual) && actual == forbidden.Value);
            Assert.Equal(item.DocumentType, detector.ResolveDocumentType(snapshot));
        }

        Assert.Equal(0, falseNegatives);
        Assert.Equal(0, falsePositives);
    }

    [Fact]
    public void Large_synthetic_role_scan_is_deterministic()
    {
        var paragraphs = Enumerable.Range(1, 1500).Select(index =>
            new LocalParagraphSnapshot(index,
                index % 50 == 1 ? "QUYẾT ĐỊNH" :
                index % 50 == 2 ? "Phê duyệt kế hoạch công tác" :
                index % 50 == 3 ? "Căn cứ Luật Tổ chức chính quyền địa phương;" :
                "Nội dung thực hiện nhiệm vụ theo kế hoạch đã được phê duyệt.",
                "wdMainTextStory", 1, index * 100, "Times New Roman",
                fontSizePoints: 13, bold: index % 50 <= 2, alignment: 3)).ToArray();
        var snapshot = new LocalScanSnapshot("golden:performance", 1,
            Array.Empty<LocalSectionSnapshot>(), paragraphs,
            Array.Empty<ChuanHoa.Client.Core.Annotations.AnnotationProtectedSpan>());
        var detector = new DocumentRoleDetector();

        var first = detector.Detect(snapshot); // warm-up and deterministic baseline
        var allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        var second = detector.Detect(snapshot);
        stopwatch.Stop();
        var allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;

        Assert.Equal(first.OrderBy(item => item.Key), second.OrderBy(item => item.Key));
        Assert.NotEmpty(first);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5),
            $"Role scan exceeded 5 seconds: {stopwatch.Elapsed.TotalMilliseconds:F1} ms");
        Assert.True(allocated < 128 * 1024 * 1024,
            $"Role scan allocated too much memory: {allocated} bytes");
        Console.WriteLine($"ACCURACY_PERFORMANCE paragraphs={paragraphs.Length} " +
            $"elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1} findings={second.Count} " +
            $"allocatedBytes={allocated}");
    }
}
