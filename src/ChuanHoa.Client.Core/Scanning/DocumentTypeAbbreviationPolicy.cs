using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Resolves the type-abbreviation group of the "Số, ký hiệu" component.
    /// Phụ lục III NĐ 30/2020/NĐ-CP fixes four abbreviations that keep lower-case
    /// letters (TTr, QyĐ, CTr, TrS); they are written exactly as listed and are
    /// never upper-cased. HD 05-HD/VPTW gives "Quy định" the same QĐ as
    /// "Quyết định", so the party regime overrides the ND30 QyĐ.
    /// </summary>
    public static class DocumentTypeAbbreviationPolicy
    {
        private static readonly string[] OfficialMixedCaseAbbreviations = { "TTr", "QyĐ", "CTr", "TrS" };

        private static readonly IReadOnlyDictionary<string, DocumentTypeAbbreviationRule> PartyOverrides =
            new Dictionary<string, DocumentTypeAbbreviationRule>(StringComparer.OrdinalIgnoreCase)
            {
                ["Quy định"] = new DocumentTypeAbbreviationRule("Quy định", "QĐ")
            };

        /// <summary>
        /// Returns the abbreviation rule for the detected type name, or null when the
        /// rule pack does not list that type. Matching ignores case and spacing.
        /// </summary>
        public static DocumentTypeAbbreviationRule? Resolve(LocalRulePack rules, string? typeNameText, bool party)
        {
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            var typeName = Regex.Replace((typeNameText ?? string.Empty).Normalize(NormalizationForm.FormC), @"\s+", " ")
                .Trim().Trim('.', ':').Trim();
            if (typeName.Length == 0) return null;
            if (party && PartyOverrides.TryGetValue(typeName, out var partyRule)) return partyRule;
            return rules.DocumentTypeAbbreviations.FirstOrDefault(item =>
                string.Equals(item.TypeName.Normalize(NormalizationForm.FormC), typeName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// True when <paramref name="group"/> is, letter for letter and case for case,
        /// an official abbreviation that contains lower-case letters.
        /// </summary>
        public static bool IsOfficialMixedCase(string? group, LocalRulePack? rules = null)
        {
            if (string.IsNullOrEmpty(group)) return false;
            var value = group!.Normalize(NormalizationForm.FormC);
            return OfficialMixedCaseAbbreviations.Contains(value, StringComparer.Ordinal) ||
                (rules != null && rules.DocumentTypeAbbreviations.Any(item =>
                    HasLowercase(item.Abbreviation) &&
                    string.Equals(item.Abbreviation.Normalize(NormalizationForm.FormC), value, StringComparison.Ordinal)));
        }

        public static bool HasLowercase(string? value) => (value ?? string.Empty).Any(char.IsLower);
    }
}
