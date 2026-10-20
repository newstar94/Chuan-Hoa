using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml.Linq;

namespace ChuanHoa.Client.Core.Rules
{
/// <summary>Immutable application resources protected by the signed installer/assembly.
/// No network, identity, device binding or expiring usage grant.</summary>
public static class BundledRulePack
{
    public static LocalRulePack Load(string clientReleaseId)
    {
        var xml = BuildRulePack(DateTimeOffset.MinValue, DateTimeOffset.MaxValue, "1.0.0.0");
        return LocalRulePackParser.Parse(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)),
            DateTimeOffset.UtcNow, clientReleaseId);
    }

    private static XElement BuildRulePack(DateTimeOffset now, DateTimeOffset expires, string clientReleaseId)
    {
        var corrections = LoadCorrections();
        var lexicon = LoadLexicon();
        var advisoryProfile = new XElement("profile", new XAttribute("code", "AcademicTypography"), new XAttribute("enabled", "false"), new XAttribute("detectorPolicyVersion", "1"), new XElement("enabledRules"), new XElement("autoFixRules"), new XElement("thresholds", new XAttribute("headingConfidenceMinimum", ".90"), new XAttribute("bodyConfidenceMinimum", ".95"), new XAttribute("captionMaxBlankParagraphs", "1"), new XAttribute("mathMinimumSignalCount", "1")));
        return new XElement("rulePack",
            new XAttribute("schema", LocalRulePackParser.Schema),
            new XAttribute("packId", "BUNDLED-FREE-ND30-V2"),
            new XAttribute("version", "2.0.0"),
            new XAttribute("notBeforeUtc", DateTimeOffset.MinValue.ToString("O", CultureInfo.InvariantCulture)),
            new XAttribute("expiresAtUtc", expires.ToString("O", CultureInfo.InvariantCulture)),
            new XAttribute("minimumClientReleaseId", clientReleaseId),
            new XElement("format",
                new XAttribute("a4WidthMm", "210"), new XAttribute("a4HeightMm", "297"),
                new XAttribute("topMinMm", "20"), new XAttribute("topMaxMm", "25"),
                new XAttribute("bottomMinMm", "20"), new XAttribute("bottomMaxMm", "25"),
                new XAttribute("leftMinMm", "30"), new XAttribute("leftMaxMm", "35"),
                new XAttribute("rightMinMm", "15"), new XAttribute("rightMaxMm", "20"),
                new XAttribute("bodyFontName", "Times New Roman"),
                new XAttribute("bodyFontMinPoints", "13"), new XAttribute("bodyFontMaxPoints", "14"),
                new XAttribute("mottoLineMinRatio", "0.95"), new XAttribute("mottoLineMaxRatio", "1.05"),
                new XAttribute("organLineMinRatio", "0.3"), new XAttribute("organLineMaxRatio", "0.55"),
                new XAttribute("subjectLineMinRatio", "0.3"), new XAttribute("subjectLineMaxRatio", "0.55"),
                new XAttribute("partyTitleLineMinRatio", "0.8"), new XAttribute("partyTitleLineMaxRatio", "1.2"),
                new XAttribute("bodyFirstLineIndentMinMm", "10"), new XAttribute("bodyFirstLineIndentMaxMm", "12.7"),
                new XAttribute("bodySpaceAfterMinPoints", "6"), new XAttribute("bodyAlignment", "3")),
            new XElement("corrections", corrections.Select(item =>
                new XElement("correction", new XAttribute("wrong", item.Key), new XAttribute("replacement", item.Value)))),
            new XElement("lexicon", lexicon.Select(item =>
                new XElement("word", new XAttribute("value", item)))),
            new XElement("telex",
                new XElement("rule",
                    new XAttribute("pattern", @"\b(?=[A-Za-z]*[sfrxj]\b)(?=[A-Za-z]*(?:aw|aa|dd|ee|oo|ow|uw))[A-Za-z]{3,}\b"),
                    new XAttribute("replacement", "từ tiếng Việt tương ứng"))),
            new XElement("hiddenCharacters",
                new XElement("character", new XAttribute("codePoint", "200B")),
                new XElement("character", new XAttribute("codePoint", "200C")),
                new XElement("character", new XAttribute("codePoint", "200D")),
                new XElement("character", new XAttribute("codePoint", "FEFF")),
                new XElement("character", new XAttribute("codePoint", "00A0"))),
            new XElement("capitalizations",
                LoadCapitalizationEntries().Select(item => new XElement("entry",
                    new XAttribute("category", item.Category), new XAttribute("expected", item.Expected)))),
            new XElement("documentTypeAbbreviations",
                DocumentTypeEntries().Select(item => new XElement("entry",
                    new XAttribute("typeName", item.TypeName), new XAttribute("abbreviation", item.Abbreviation)))),
            new XElement("advisoryProfiles", advisoryProfile));
    }


    private static Stream Open(string name) =>
        typeof(BundledRulePack).Assembly.GetManifestResourceStream("ChuanHoa.Rules." + name)
        ?? throw new InvalidOperationException("BUNDLED_RULE_RESOURCE_MISSING: " + name);

    private static T ReadJson<T>(string name)
    {
        using (var stream = Open(name))
            return (T)new DataContractJsonSerializer(typeof(T),
                new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true }).ReadObject(stream);
    }

    private static IReadOnlyDictionary<string, string> LoadCorrections() =>
        ReadJson<Dictionary<string, string>>("typo_dictionary.json")
            .Where(item => !string.IsNullOrWhiteSpace(item.Key) && !string.IsNullOrWhiteSpace(item.Value) &&
                !string.Equals(item.Key, item.Value, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);

    private static IReadOnlyList<string> LoadLexicon()
    {
        var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in new[] { "vi-DauCu.dic", "vi-DauMoi.dic" })
        using (var stream = Open(name))
        using (var reader = new StreamReader(stream, Encoding.UTF8))
        {
            string? raw;
            while ((raw = reader.ReadLine()) != null)
            {
                var value = raw.Trim();
                if (value.Length == 0 || value.All(char.IsDigit)) continue;
                var flag = value.IndexOf('/');
                if (flag >= 0) value = value.Substring(0, flag);
                if (value.Length > 0) words.Add(value.Normalize(NormalizationForm.FormC));
            }
        }
        if (words.Count < 1000) throw new InvalidOperationException("BUNDLED_LEXICON_INCOMPLETE");
        return words.OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    [DataContract]
    private sealed class AdministrativeUnit
    {
        [DataMember(Name = "status")] public string Status { get; set; } = string.Empty;
        [DataMember(Name = "canonicalName")] public string CanonicalName { get; set; } = string.Empty;
    }

    private static IEnumerable<(string Category, string Expected)> LoadCapitalizationEntries()
    {
        foreach (var item in ReadJson<AdministrativeUnit[]>("administrative_units.json"))
            if (item.Status == "Active" && !string.IsNullOrWhiteSpace(item.CanonicalName))
                yield return ("administrative", item.CanonicalName);
        foreach (var expected in ReadJson<string[]>("special_capitalizations.json"))
            if (!string.IsNullOrWhiteSpace(expected)) yield return ("specialOrgan", expected);
        yield return ("geographic", "Việt Nam");
        yield return ("geographic", "Đông Nam Á");
        yield return ("terrain", "sông Hồng");
        yield return ("terrain", "dãy Trường Sơn");
        yield return ("region", "Tây Nguyên");
        yield return ("region", "Đồng bằng sông Cửu Long");
        yield return ("organ", "Bộ Nội vụ");
        yield return ("organ", "Chính phủ");
        yield return ("organ", "Quốc hội");
        yield return ("specialOrgan", "Ban Chấp hành Trung ương");
        yield return ("holiday", "Ngày Quốc khánh");
        yield return ("holiday", "Ngày Quốc tế Lao động");
        yield return ("lunarYear", "Giáp Thìn");
        yield return ("lunarYear", "Ất Tỵ");
        yield return ("lunarYear", "Bính Ngọ");
    }

    private static IEnumerable<(string TypeName, string Abbreviation)> DocumentTypeEntries()
    {
        yield return ("Nghị quyết", "NQ");
        yield return ("Quyết định", "QĐ");
        yield return ("Chỉ thị", "CT");
        yield return ("Thông tư", "TT");
        yield return ("Thông báo", "TB");
        yield return ("Kế hoạch", "KH");
        yield return ("Báo cáo", "BC");
        yield return ("Tờ trình", "TTr");
        yield return ("Công văn", "CV");
    }


}
}
