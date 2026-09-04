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

            var notation = match.Groups["notation"].Value.ToUpperInvariant().Trim();
            var parts = Regex.Split(notation, @"[\s/\-.]+")
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .ToList();
            var expected = (expectedTypeAbbreviation ?? string.Empty).Trim().ToUpperInvariant();

            if (party)
            {
                if (expected.Length > 0)
                {
                    var expectedIndex = parts.FindIndex(part =>
                        string.Equals(part, expected, StringComparison.OrdinalIgnoreCase));
                    if (expectedIndex >= 0) parts = parts.Skip(expectedIndex).ToList();
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
                if (expectedIndex >= 0) parts = parts.Skip(expectedIndex).ToList();
                else ReplaceLeadingTypeGroup(parts, expected);
            }
            var administrativeNotation = string.Join("-", parts);
            return "Số: " + number + "/" + administrativeNotation;
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
