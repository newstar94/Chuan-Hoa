using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Lexicon;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests
{
    public sealed class SpellCacheAndFeedbackTests
    {
        [Fact]
        public void ParagraphContentCache_CachesAndEvictsLru()
        {
            var cache = new ParagraphContentCache(capacity: 2);

            var finding1 = new AnnotationFinding("f1", "ND30", "warning", "test", "exp", "citation",
                new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "wdMainTextStory", 1, 0, 5, "Hello"));
            var finding2 = new AnnotationFinding("f2", "ND30", "warning", "test", "exp", "citation",
                new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "wdMainTextStory", 2, 0, 5, "World"));
            var finding3 = new AnnotationFinding("f3", "ND30", "warning", "test", "exp", "citation",
                new AnnotationAnchor(AnnotationAnchorKind.TextSpan, "wdMainTextStory", 3, 0, 5, "Third"));

            var hash1 = ParagraphContentCache.ComputeParagraphHash("Paragraph 1");
            var hash2 = ParagraphContentCache.ComputeParagraphHash("Paragraph 2");
            var hash3 = ParagraphContentCache.ComputeParagraphHash("Paragraph 3");

            cache.Set(hash1, new[] { finding1 });
            cache.Set(hash2, new[] { finding2 });

            Assert.True(cache.TryGet(hash1, out var cached1));
            Assert.Single(cached1);

            // Accessing hash1 makes hash2 the least recently used
            cache.Set(hash3, new[] { finding3 });

            // hash2 should be evicted because capacity is 2
            Assert.False(cache.TryGet(hash2, out _));
            Assert.True(cache.TryGet(hash1, out _));
            Assert.True(cache.TryGet(hash3, out _));
        }

        [Fact]
        public void LocalFeedbackStore_AdjustsPreferenceScores()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "FeedbackTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var store = new LocalFeedbackStore(tempDir);

                // Initial adjustment with no feedback
                Assert.Equal(0.0, store.GetPreferenceAdjustment("bàn dao", "bàn giao"));

                // User accepts 3 times
                store.RecordFeedback("bàn dao", "bàn giao", true);
                store.RecordFeedback("bàn dao", "bàn giao", true);
                store.RecordFeedback("bàn dao", "bàn giao", true);

                var boost = store.GetPreferenceAdjustment("bàn dao", "bàn giao");
                Assert.True(boost > 0, "Accepted suggestion should receive positive preference boost");

                // User rejects another suggestion 3 times
                store.RecordFeedback("sử lý", "xử lí", false);
                store.RecordFeedback("sử lý", "xử lí", false);
                store.RecordFeedback("sử lý", "xử lí", false);

                var penalty = store.GetPreferenceAdjustment("sử lý", "xử lí");
                Assert.True(penalty < 0, "Rejected suggestion should receive negative penalty");
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
        }

        [Fact]
        public async Task IpcClient_FailsSafelyWhenEngineOffline()
        {
            var client = new VietnameseEngineIpcClient();

            // When engine process is not running on the pipe, PingAsync must complete within timeout and return false
            var isAlive = await client.PingAsync(timeoutMs: 150);
            Assert.False(isAlive);

            // CheckAsync must fail-closed and return empty list without throwing
            var findings = await client.CheckAsync("Đồng chí gửi bàn dao hồ sơ", paragraphIndex: 1, timeoutMs: 150);
            Assert.NotNull(findings);
            Assert.Empty(findings);
        }
    }
}
