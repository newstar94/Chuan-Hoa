using System;
using System.IO;
using System.Linq;
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
        public void CandidateGenerator_IncludesOriginalWordAndResolvesTelexTypo()
        {
            var generator = new FastCandidateGenerator();
            var candidates = generator.GenerateCandidates("ngỉ");

            Assert.NotEmpty(candidates);
            Assert.True(candidates.Count <= 8);

            // Candidate 0 should be the original word (to allow NO_CHANGE decision in ranking model)
            Assert.Equal("ngỉ", candidates[0].Text);

            // Should suggest "nghỉ"
            Assert.Contains(candidates, c => c.Text.Equals("nghỉ", StringComparison.OrdinalIgnoreCase));
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
                dictManager.AddUserWord("PC04");
                Assert.True(dictManager.IsKnownOrIgnored("pc04")); // Case-insensitive
                Assert.True(dictManager.IsKnownOrIgnored("PC04"));

                // Document specific ignore
                const string docId = "doc-123";
                Assert.False(dictManager.IsKnownOrIgnored("CSGT", docId));
                dictManager.IgnoreWordForDocument(docId, "CSGT");
                Assert.True(dictManager.IsKnownOrIgnored("csgt", docId));
                Assert.False(dictManager.IsKnownOrIgnored("csgt", "other-doc")); // Scoped to docId

                // Remove user word
                dictManager.RemoveUserWord("PC04");
                Assert.False(dictManager.IsKnownOrIgnored("PC04"));
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
        }
    }
}
