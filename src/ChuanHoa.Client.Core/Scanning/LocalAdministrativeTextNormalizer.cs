using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Deterministic text corrections that are safe for 1-Click. These methods do
    /// not guess prose; they only normalize fixed administrative components.
    /// </summary>
    public static class LocalAdministrativeTextNormalizer
    {
        public const string NationalMotto = "Độc lập - Tự do - Hạnh phúc";

        private static readonly Regex CodeNumber = new Regex(
            @"^\s*Số\s*:?[\s]*(?<number>\d+)\s*(?:[/\-])?\s*(?<notation>[^\r\n]*)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
            TimeSpan.FromMilliseconds(200));

        public static string NormalizeCodeNumber(string value, bool party,
            string expectedTypeAbbreviation)
        {
            var match = CodeNumber.Match(value ?? string.Empty);
            if (!match.Success) return value ?? string.Empty;

            var number = match.Groups["number"].Value;
            int numeric;
            if (int.TryParse(number, NumberStyles.None, CultureInfo.InvariantCulture, out numeric) &&
                numeric < 10)
                number = numeric.ToString("00", CultureInfo.InvariantCulture);

            // Groups are upper-cased except an official mixed-case type abbreviation
            // (TTr, QyĐ, CTr, TrS in Phụ lục III), which is kept exactly as written.
            var parts = Regex.Split(match.Groups["notation"].Value.Trim(), @"[\s/\-.]+")
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .Select(part => DocumentTypeAbbreviationPolicy.IsOfficialMixedCase(part)
                    ? part
                    : part.ToUpperInvariant())
                .ToList();
            // The expected abbreviation comes from the rule pack in its official case.
            var expected = (expectedTypeAbbreviation ?? string.Empty).Trim();

            if (party)
            {
                if (expected.Length > 0)
                {
                    var expectedIndex = parts.FindIndex(part =>
                        string.Equals(part, expected, StringComparison.OrdinalIgnoreCase));
                    if (expectedIndex >= 0)
                    {
                        parts = parts.Skip(expectedIndex).ToList();
                        parts[0] = expected;
                    }
                    else ReplaceLeadingTypeGroup(parts, expected);
                }
                var partyNotation = parts.Count <= 1
                    ? string.Join(string.Empty, parts)
                    : parts[0] + "/" + string.Join("-", parts.Skip(1));
                return "Số " + number + "-" + partyNotation;
            }

            if (expected.Length > 0)
            {
                var expectedIndex = parts.FindIndex(part =>
                    string.Equals(part, expected, StringComparison.OrdinalIgnoreCase));
                if (expectedIndex >= 0)
                {
                    parts = parts.Skip(expectedIndex).ToList();
                    parts[0] = expected;
                }
                else ReplaceLeadingTypeGroup(parts, expected);
            }
            var administrativeNotation = string.Join("-", parts);
            return "Số: " + number + "/" + administrativeNotation;
        }

        private static readonly Regex VbqpplCodeNumber = new Regex(
            @"^\s*Số\s*:?\s*(?<number>\d+)\s*[/\-]\s*(?<year>\d{4})\s*[/\-]\s*(?<notation>[^\r\n]*)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
            TimeSpan.FromMilliseconds(200));

        /// <summary>
        /// VBQPPL form "Số: 05/2025/QĐ-UBND" (Phụ lục I NĐ 187/2025/NĐ-CP, Mục 1, Phần II,
        /// khoản 3 điểm d, đ; Mẫu số 17-20). The issuing year is part of the notation and is
        /// kept. Text without a four-digit year is returned unchanged: the year cannot be
        /// inferred safely.
        /// </summary>
        public static string NormalizeVbqpplCodeNumber(string value, string expectedTypeAbbreviation)
        {
            var match = VbqpplCodeNumber.Match(value ?? string.Empty);
            if (!match.Success) return value ?? string.Empty;

            var number = match.Groups["number"].Value;
            int numeric;
            if (int.TryParse(number, NumberStyles.None, CultureInfo.InvariantCulture, out numeric) &&
                numeric < 10)
                number = numeric.ToString("00", CultureInfo.InvariantCulture);

            var parts = Regex.Split(match.Groups["notation"].Value.Trim(), @"[\s/\-.]+")
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .Select(part => part.ToUpper(CultureInfo.GetCultureInfo("vi-VN")))
                .ToList();
            if (parts.Count == 0) return value ?? string.Empty;
            var expected = (expectedTypeAbbreviation ?? string.Empty).Trim();
            if (expected.Length > 0)
            {
                expected = expected.ToUpper(CultureInfo.GetCultureInfo("vi-VN"));
                // "05/2025/UBND" has only the organ group: add the type group instead
                // of overwriting the organ.
                if (parts.Count >= 2) parts[0] = expected;
                else if (!string.Equals(parts[0], expected, StringComparison.Ordinal)) parts.Insert(0, expected);
            }
            return "Số: " + number + "/" + match.Groups["year"].Value + "/" + string.Join("-", parts);
        }

        private static void ReplaceLeadingTypeGroup(IList<string> parts, string expected)
        {
            // A frequent legacy form repeats the year between the registration
            // number and type abbreviation (129/2026/QĐ-...). The ND30 notation
            // starts with the type abbreviation, so discard only leading numeric
            // groups before replacing the first textual group.
            while (parts.Count > 0 && parts[0].All(char.IsDigit)) parts.RemoveAt(0);
            if (parts.Count == 0) parts.Add(expected);
            else parts[0] = expected;
        }
    }
}
