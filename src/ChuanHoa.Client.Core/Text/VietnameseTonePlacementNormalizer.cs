using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Text
{
    public enum VietnameseTonePlacementStyle
    {
        MainVowel,
        FirstVowel
    }

    public static class VietnameseTonePlacementNormalizer
    {
        private const string OConsonants = "bcdđghklmnpqrstvx";
        private const string UConsonants = "bcdđghklmnprstvx";

        private static readonly IReadOnlyDictionary<string, string> ToMain = BuildMap(new[]
        {
            "òa=oà", "óa=oá", "ỏa=oả", "õa=oã", "ọa=oạ",
            "òe=oè", "óe=oé", "ỏe=oẻ", "õe=oẽ", "ọe=oẹ",
            "ùy=uỳ", "úy=uý", "ủy=uỷ", "ũy=uỹ", "ụy=uỵ"
        });

        private static readonly IReadOnlyDictionary<string, string> ToFirst = Reverse(ToMain);
        private static readonly Regex WordRegex = new Regex(
            @"[A-Za-zÀ-ỹĐđ]+", RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static string Normalize(string text, VietnameseTonePlacementStyle style)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            var map = style == VietnameseTonePlacementStyle.MainVowel ? ToMain : ToFirst;
            return WordRegex.Replace(text, match => NormalizeWord(match.Value, map));
        }

        private static string NormalizeWord(string word, IReadOnlyDictionary<string, string> map)
        {
            if (word.Length < 2) return word;
            var lower = word.ToLowerInvariant();
            foreach (var pair in map)
            {
                var index = lower.LastIndexOf(pair.Key, StringComparison.Ordinal);
                if (index < 0 || index + pair.Key.Length != lower.Length) continue;

                var isWordInitialForm = index == 0 &&
                    (pair.Key == "òa" || pair.Key == "óa" || pair.Key == "ọa" || pair.Key == "ủy" ||
                     pair.Key == "oà" || pair.Key == "oá" || pair.Key == "oạ" || pair.Key == "uỷ");
                if (!isWordInitialForm)
                {
                    if (index == 0) continue;
                    var previous = lower[index - 1];
                    var allowed = pair.Key.EndsWith("y", StringComparison.Ordinal) ? UConsonants : OConsonants;
                    if (allowed.IndexOf(previous) < 0) continue;
                }

                var replacement = MatchCase(word.Substring(index, pair.Key.Length), pair.Value);
                return word.Substring(0, index) + replacement;
            }
            return word;
        }

        private static string MatchCase(string source, string replacement)
        {
            if (source == source.ToUpperInvariant()) return replacement.ToUpperInvariant();
            if (char.IsUpper(source[0]))
                return char.ToUpperInvariant(replacement[0]) + replacement.Substring(1);
            return replacement;
        }

        private static IReadOnlyDictionary<string, string> BuildMap(IEnumerable<string> pairs)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var pair in pairs)
            {
                var separator = pair.IndexOf('=');
                result.Add(pair.Substring(0, separator), pair.Substring(separator + 1));
            }
            return result;
        }

        private static IReadOnlyDictionary<string, string> Reverse(IReadOnlyDictionary<string, string> source)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var pair in source) result.Add(pair.Value, pair.Key);
            return result;
        }
    }
}
