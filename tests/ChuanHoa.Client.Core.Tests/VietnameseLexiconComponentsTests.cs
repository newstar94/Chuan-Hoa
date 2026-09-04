using System;
using System.IO;
using System.Linq;
using System.Text;
using ChuanHoa.Client.Core.Lexicon;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests
{
    public sealed class VietnameseLexiconComponentsTests
    {
        [Fact]
        public void ConfusionSets_GeneratesExpectedPhoneticAlternates()
        {
            // Test s/x
            var alternatesSX = VietnameseConfusionSets.GeneratePhoneticAlternates("xử").ToList();
            Assert.Contains("sử", alternatesSX);

            // Test ch/tr
            var alternatesChTr = VietnameseConfusionSets.GeneratePhoneticAlternates("chân").ToList();
            Assert.Contains("trân", alternatesChTr);

            // Test d/gi/r
            var alternatesD = VietnameseConfusionSets.GeneratePhoneticAlternates("dao").ToList();
            Assert.Contains("giao", alternatesD);
            Assert.Contains("rao", alternatesD);

            // Test hỏi / ngã
            var alternatesHoiNga = VietnameseConfusionSets.GeneratePhoneticAlternates("nghĩ").ToList();
            Assert.Contains("nghỉ", alternatesHoiNga);

            var alternatesHoiNga2 = VietnameseConfusionSets.GeneratePhoneticAlternates("hướng dẩn").ToList();
            Assert.Contains("hướng dẫn", alternatesHoiNga2);
        }

        [Fact]
        public void ConfusionSets_ContainsAdministrativeConfusionPairs()
        {
            Assert.True(VietnameseConfusionSets.AdministrativeConfusionPairs.TryGetValue("bàn dao", out var banGiao));
            Assert.Equal("bàn giao", banGiao);

            Assert.True(VietnameseConfusionSets.AdministrativeConfusionPairs.TryGetValue("sử lý", out var xuLy));
            Assert.Equal("xử lý", xuLy);

            Assert.True(VietnameseConfusionSets.AdministrativeConfusionPairs.TryGetValue("điều khoảng", out var dieuKhoan));
            Assert.Equal("điều khoản", dieuKhoan);
        }

        [Fact]
        public void Tokenizer_PreservesExactOffsetsAndClassifiesCorrectly()
        {
            const string text = "Đồng chí gửi email: contact@chuanhoa.gov.vn hoặc truy cập https://chuanhoa.gov.vn trước ngày 02/09/2026.";
            var tokens = VietnameseWordTokenizer.TokenizeParagraph(text, 1);

            Assert.NotEmpty(tokens);

            // Verify every token's substring exactly matches text at specified offset
            foreach (var token in tokens)
            {
                Assert.Equal(token.Text, text.Substring(token.StartOffset, token.Length));
                Assert.Equal(1, token.ParagraphIndex);
            }

            // Verify URL / Email classification
            Assert.Contains(tokens, t => t.Kind == VietnameseTokenKind.UrlOrEmail && t.Text == "contact@chuanhoa.gov.vn");
            Assert.Contains(tokens, t => t.Kind == VietnameseTokenKind.UrlOrEmail && t.Text == "https://chuanhoa.gov.vn");
            Assert.Contains(tokens, t => t.Kind == VietnameseTokenKind.Word && t.Text == "Đồng");
            Assert.Contains(tokens, t => t.Kind == VietnameseTokenKind.Number && t.Text == "02");
        }

        [Fact]
        public void Tokenizer_PreservesOriginalOffsetsForDecomposedVietnameseText()
        {
            const string text = "Ho\u0300a bi\u0300nh";
            var words = VietnameseWordTokenizer.TokenizeParagraph(text, 7)
                .Where(token => token.Kind == VietnameseTokenKind.Word)
                .ToArray();

            Assert.Equal(2, words.Length);
            Assert.All(words, token => Assert.Equal(token.Text,
                text.Substring(token.StartOffset, token.Length)));
            Assert.Equal("Hòa", words[0].Text.Normalize(NormalizationForm.FormC));
            Assert.Equal(7, words[0].ParagraphIndex);
        }

        [Fact]
        public void Nfc_offset_map_returns_the_exact_decomposed_source_span()
        {
            const string source = "A Ho\u0300a bi\u0300nh Z";
            var mapped = NormalizedTextOffsetMap.Create(source);
            var normalizedStart = mapped.Normalized.IndexOf("Hòa bình", StringComparison.Ordinal);

            var sourceSpan = mapped.MapSpan(normalizedStart, "Hòa bình".Length);

            Assert.Equal(source.IndexOf("Ho\u0300a", StringComparison.Ordinal), sourceSpan.Item1);
            Assert.Equal("Ho\u0300a bi\u0300nh", source.Substring(sourceSpan.Item1, sourceSpan.Item2));
        }

        [Fact]
        public void PersonalDictionaryManager_SupportsAddingAndDocumentIgnoring()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "ChuanHoaTestDict_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var dictManager = new PersonalDictionaryManager(tempDir);

                // Initially unknown
                Assert.False(dictManager.IsKnownOrIgnored("PC04"));

                // Add user word
                Assert.True(dictManager.AddUserWord("PC04").Succeeded);
                Assert.True(dictManager.IsKnownOrIgnored("pc04")); // Case-insensitive
                Assert.True(dictManager.IsKnownOrIgnored("PC04"));

                // Document specific ignore
                const string docId = "doc-123";
                Assert.False(dictManager.IsKnownOrIgnored("CSGT", docId));
                Assert.True(dictManager.IgnoreWordForDocument(docId, "CSGT").Succeeded);
                Assert.True(dictManager.IsKnownOrIgnored("csgt", docId));
                Assert.False(dictManager.IsKnownOrIgnored("csgt", "other-doc")); // Scoped to docId

                // Remove user word
                Assert.True(dictManager.RemoveUserWord("PC04").Succeeded);
                Assert.False(dictManager.IsKnownOrIgnored("PC04"));

                Assert.Equal(PersonalDictionaryStatus.Success,
                    dictManager.ClearDocumentIgnores(docId).Status);
                Assert.False(dictManager.IsKnownOrIgnored("CSGT", docId));
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
        public void PersonalDictionaryManager_ClearsAllDocumentIgnoresWithoutChangingUserWords()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "ChuanHoaTestDict_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var manager = new PersonalDictionaryManager(tempDir);
                Assert.True(manager.AddUserWord("PC04").Succeeded);
                Assert.True(manager.IgnoreWordForDocument("doc-a", "CSGT").Succeeded);
                Assert.True(manager.IgnoreWordForDocument("doc-b", "KHLCNT").Succeeded);

                Assert.Equal(PersonalDictionaryStatus.Success, manager.ClearAllDocumentIgnores().Status);
                Assert.False(manager.IsKnownOrIgnored("CSGT", "doc-a"));
                Assert.False(manager.IsKnownOrIgnored("KHLCNT", "doc-b"));
                Assert.True(manager.IsKnownOrIgnored("PC04"));
                Assert.Equal(PersonalDictionaryStatus.NoChange, manager.ClearAllDocumentIgnores().Status);
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void PersonalDictionaryManager_NormalizesNfcAndPersistsAtomically()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "ChuanHoaTestDict_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var manager = new PersonalDictionaryManager(tempDir);
                var decomposed = "Ho\u0300a";
                Assert.Equal(PersonalDictionaryStatus.Success, manager.AddUserWord(decomposed).Status);
                Assert.Equal(PersonalDictionaryStatus.Duplicate, manager.AddUserWord("Hòa").Status);
                Assert.Equal(PersonalDictionaryStatus.Success, manager.AddUserWord("thuật ngữ riêng").Status);
                Assert.Contains("Hòa", manager.GetUserWordsResult().Words);
                Assert.Empty(Directory.GetFiles(tempDir, "*.tmp"));
                Assert.True(File.Exists(Path.Combine(tempDir, "user_custom_dictionary.txt.last-good")));

                var reloaded = new PersonalDictionaryManager(tempDir);
                Assert.True(reloaded.IsKnownOrIgnored("Hòa"));
                Assert.All(reloaded.GetUserWords(), value =>
                    Assert.True(value.IsNormalized(NormalizationForm.FormC)));
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void PersonalDictionaryManager_LoadsLastGoodWhenPrimaryFileIsInvalidUtf8()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "ChuanHoaTestDict_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var manager = new PersonalDictionaryManager(tempDir);
                Assert.True(manager.AddUserWord("thuật ngữ một").Succeeded);
                Assert.True(manager.AddUserWord("thuật ngữ hai").Succeeded);

                var primaryPath = Path.Combine(tempDir, "user_custom_dictionary.txt");
                File.WriteAllBytes(primaryPath, new byte[] { 0xC3, 0x28 });

                var recovered = new PersonalDictionaryManager(tempDir);
                var snapshot = recovered.GetUserWordsResult();
                Assert.Equal(PersonalDictionaryStatus.IoError, snapshot.Persistence.Status);
                Assert.True(recovered.IsKnownOrIgnored("thuật ngữ một"));
                Assert.False(recovered.IsKnownOrIgnored("thuật ngữ hai"));
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void PersonalDictionaryManager_ReturnsTypedValidationAndIoFailures()
        {
            var tempRoot = Path.Combine(Path.GetTempPath(), "ChuanHoaTestDict_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            try
            {
                var manager = new PersonalDictionaryManager(Path.Combine(tempRoot, "dictionary"));
                Assert.Equal(PersonalDictionaryStatus.InvalidInput, manager.AddUserWord("\r\n").Status);
                Assert.Equal(PersonalDictionaryStatus.LimitExceeded,
                    manager.AddUserWord(new string('a', PersonalDictionaryManager.MaximumEntryLength + 1)).Status);
                Assert.Equal(PersonalDictionaryStatus.InvalidInput,
                    manager.ClearDocumentIgnores(string.Empty).Status);

                var blockedPath = Path.Combine(tempRoot, "blocked");
                File.WriteAllText(blockedPath, "not a directory");
                var blocked = new PersonalDictionaryManager(blockedPath);
                var result = blocked.AddUserWord("thuật ngữ");
                Assert.Equal(PersonalDictionaryStatus.IoError, result.Status);
                Assert.False(blocked.IsKnownOrIgnored("thuật ngữ"));
            }
            finally
            {
                if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, true);
            }
        }
    }
}
