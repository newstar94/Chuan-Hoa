using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ChuanHoa.Client.Core.Rules
{
    public sealed class TextCorrectionRule
    {
        public TextCorrectionRule(string wrong, string replacement) { Wrong = wrong; Replacement = replacement; }
        public string Wrong { get; }
        public string Replacement { get; }
    }

    public sealed class TelexRule
    {
        public TelexRule(string pattern, string replacement) { Pattern = pattern; Replacement = replacement; }
        public string Pattern { get; }
        public string Replacement { get; }
    }

    public sealed class CapitalizationRule
    {
        public CapitalizationRule(string category, string expected)
        {
            Category = category ?? string.Empty;
            Expected = expected ?? string.Empty;
        }
        public string Category { get; }
        public string Expected { get; }
    }

    public sealed class DocumentTypeAbbreviationRule
    {
        public DocumentTypeAbbreviationRule(string typeName, string abbreviation)
        {
            TypeName = typeName ?? string.Empty;
            Abbreviation = abbreviation ?? string.Empty;
        }
        public string TypeName { get; }
        public string Abbreviation { get; }
    }

    public sealed class LocalRulePack
    {
        public LocalRulePack(string packId, string version, DateTimeOffset notBeforeUtc, DateTimeOffset expiresAtUtc,
            string minimumClientReleaseId, double a4WidthMm, double a4HeightMm, double topMinMm, double topMaxMm,
            double bottomMinMm, double bottomMaxMm, double leftMinMm, double leftMaxMm, double rightMinMm,
            double rightMaxMm, string bodyFontName, IReadOnlyList<TextCorrectionRule> corrections,
            IReadOnlyList<TelexRule> telexRules, IReadOnlyList<char> hiddenCharacters,
            double bodyFontMinPoints = 13, double bodyFontMaxPoints = 14,
            double bodyFirstLineIndentMinMm = 10, double bodyFirstLineIndentMaxMm = 12.7,
            double bodySpaceAfterMinPoints = 6, int bodyAlignment = 3,
            IReadOnlyList<CapitalizationRule>? capitalizations = null,
            IReadOnlyList<DocumentTypeAbbreviationRule>? documentTypeAbbreviations = null,
            double mottoLineMinRatio = 0.95, double mottoLineMaxRatio = 1.05,
            double organLineMinRatio = 0.3, double organLineMaxRatio = 0.55,
            double subjectLineMinRatio = 0.3, double subjectLineMaxRatio = 0.55,
            double partyTitleLineMinRatio = 0.8, double partyTitleLineMaxRatio = 1.2,
            IReadOnlyList<string>? lexicon = null)
        {
            PackId = packId; Version = version; NotBeforeUtc = notBeforeUtc; ExpiresAtUtc = expiresAtUtc;
            MinimumClientReleaseId = minimumClientReleaseId; A4WidthMm = a4WidthMm; A4HeightMm = a4HeightMm;
            TopMinMm = topMinMm; TopMaxMm = topMaxMm; BottomMinMm = bottomMinMm; BottomMaxMm = bottomMaxMm;
            LeftMinMm = leftMinMm; LeftMaxMm = leftMaxMm; RightMinMm = rightMinMm; RightMaxMm = rightMaxMm;
            BodyFontName = bodyFontName; Corrections = corrections; TelexRules = telexRules; HiddenCharacters = hiddenCharacters;
            BodyFontMinPoints = bodyFontMinPoints; BodyFontMaxPoints = bodyFontMaxPoints;
            BodyFirstLineIndentMinMm = bodyFirstLineIndentMinMm; BodyFirstLineIndentMaxMm = bodyFirstLineIndentMaxMm;
            BodySpaceAfterMinPoints = bodySpaceAfterMinPoints; BodyAlignment = bodyAlignment;
            Capitalizations = capitalizations ?? Array.Empty<CapitalizationRule>();
            DocumentTypeAbbreviations = documentTypeAbbreviations ?? Array.Empty<DocumentTypeAbbreviationRule>();
            MottoLineMinRatio = mottoLineMinRatio; MottoLineMaxRatio = mottoLineMaxRatio;
            OrganLineMinRatio = organLineMinRatio; OrganLineMaxRatio = organLineMaxRatio;
            SubjectLineMinRatio = subjectLineMinRatio; SubjectLineMaxRatio = subjectLineMaxRatio;
            PartyTitleLineMinRatio = partyTitleLineMinRatio; PartyTitleLineMaxRatio = partyTitleLineMaxRatio;
            Lexicon = lexicon ?? Array.Empty<string>();
        }

        public string PackId { get; }
        public string Version { get; }
        public DateTimeOffset NotBeforeUtc { get; }
        public DateTimeOffset ExpiresAtUtc { get; }
        public string MinimumClientReleaseId { get; }
        public double A4WidthMm { get; }
        public double A4HeightMm { get; }
        public double TopMinMm { get; }
        public double TopMaxMm { get; }
        public double BottomMinMm { get; }
        public double BottomMaxMm { get; }
        public double LeftMinMm { get; }
        public double LeftMaxMm { get; }
        public double RightMinMm { get; }
        public double RightMaxMm { get; }
        public string BodyFontName { get; }
        public IReadOnlyList<TextCorrectionRule> Corrections { get; }
        public IReadOnlyList<TelexRule> TelexRules { get; }
        public IReadOnlyList<char> HiddenCharacters { get; }
        public double BodyFontMinPoints { get; }
        public double BodyFontMaxPoints { get; }
        public double BodyFirstLineIndentMinMm { get; }
        public double BodyFirstLineIndentMaxMm { get; }
        public double BodySpaceAfterMinPoints { get; }
        public int BodyAlignment { get; }
        public IReadOnlyList<CapitalizationRule> Capitalizations { get; }
        public IReadOnlyList<DocumentTypeAbbreviationRule> DocumentTypeAbbreviations { get; }
        public double MottoLineMinRatio { get; }
        public double MottoLineMaxRatio { get; }
        public double OrganLineMinRatio { get; }
        public double OrganLineMaxRatio { get; }
        public double SubjectLineMinRatio { get; }
        public double SubjectLineMaxRatio { get; }
        public double PartyTitleLineMinRatio { get; }
        public double PartyTitleLineMaxRatio { get; }
        public IReadOnlyList<string> Lexicon { get; }
    }

    public static class LocalRulePackParser
    {
        public const string Schema = "chuanhoa.local-rule-pack.v1";
        public static LocalRulePack Parse(byte[] payload, DateTimeOffset nowUtc, string clientReleaseId)
        {
            if (payload == null || payload.Length == 0) throw new FormatException("Rule pack payload is empty.");
            var root = XElement.Parse(Encoding.UTF8.GetString(payload), LoadOptions.None);
            if (root.Name.LocalName != "rulePack" || Attr(root, "schema") != Schema)
                throw new FormatException("Rule pack schema is unsupported.");
            var notBefore = Time(root, "notBeforeUtc");
            var expires = Time(root, "expiresAtUtc");
            if (nowUtc < notBefore) throw new InvalidOperationException("RULE_PACK_NOT_ACTIVE");
            if (nowUtc >= expires) throw new InvalidOperationException("RULE_PACK_EXPIRED");
            var minimumClient = Attr(root, "minimumClientReleaseId");
            if (CompareVersion(clientReleaseId, minimumClient) < 0) throw new InvalidOperationException("RULE_PACK_CLIENT_TOO_OLD");
            var format = root.Element("format") ?? throw new FormatException("Format rules are missing.");
            var corrections = root.Element("corrections")?.Elements("correction")
                .Select(item => new TextCorrectionRule(Attr(item, "wrong"), Attr(item, "replacement"))).ToArray()
                ?? Array.Empty<TextCorrectionRule>();
            var telex = root.Element("telex")?.Elements("rule")
                .Select(item => new TelexRule(Attr(item, "pattern"), Attr(item, "replacement"))).ToArray()
                ?? Array.Empty<TelexRule>();
            var hidden = root.Element("hiddenCharacters")?.Elements("character")
                .Select(item => (char)int.Parse(Attr(item, "codePoint"), NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToArray()
                ?? Array.Empty<char>();
            var capitalizations = root.Element("capitalizations")?.Elements("entry")
                .Select(item => new CapitalizationRule(Attr(item, "category"), Attr(item, "expected"))).ToArray()
                ?? Array.Empty<CapitalizationRule>();
            var abbreviations = root.Element("documentTypeAbbreviations")?.Elements("entry")
                .Select(item => new DocumentTypeAbbreviationRule(Attr(item, "typeName"), Attr(item, "abbreviation"))).ToArray()
                ?? Array.Empty<DocumentTypeAbbreviationRule>();
            var lexicon = root.Element("lexicon")?.Elements("word")
                .Select(item => Attr(item, "value").Normalize(NormalizationForm.FormC))
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray() ?? Array.Empty<string>();
            var mottoLine = LineRatioRange(format, "mottoLineMinRatio", "mottoLineMaxRatio", .95, 1.05);
            var organLine = LineRatioRange(format, "organLineMinRatio", "organLineMaxRatio", .3, .55);
            var subjectLine = LineRatioRange(format, "subjectLineMinRatio", "subjectLineMaxRatio", .3, .55);
            var partyTitleLine = LineRatioRange(format, "partyTitleLineMinRatio", "partyTitleLineMaxRatio", .8, 1.2);
            return new LocalRulePack(Attr(root, "packId"), Attr(root, "version"), notBefore, expires, minimumClient,
                Number(format, "a4WidthMm"), Number(format, "a4HeightMm"), Number(format, "topMinMm"),
                Number(format, "topMaxMm"), Number(format, "bottomMinMm"), Number(format, "bottomMaxMm"),
                Number(format, "leftMinMm"), Number(format, "leftMaxMm"), Number(format, "rightMinMm"),
                Number(format, "rightMaxMm"), Attr(format, "bodyFontName"), corrections, telex, hidden,
                OptionalNumber(format, "bodyFontMinPoints", 13), OptionalNumber(format, "bodyFontMaxPoints", 14),
                OptionalNumber(format, "bodyFirstLineIndentMinMm", 10), OptionalNumber(format, "bodyFirstLineIndentMaxMm", 12.7),
                OptionalNumber(format, "bodySpaceAfterMinPoints", 6), OptionalInteger(format, "bodyAlignment", 3),
                capitalizations, abbreviations,
                mottoLine.Item1, mottoLine.Item2, organLine.Item1, organLine.Item2,
                subjectLine.Item1, subjectLine.Item2, partyTitleLine.Item1, partyTitleLine.Item2,
                lexicon);
        }

        private static int CompareVersion(string actual, string required)
        {
            Version left, right;
            return Version.TryParse(actual, out left) && Version.TryParse(required, out right)
                ? left.CompareTo(right) : string.CompareOrdinal(actual, required);
        }
        private static string Attr(XElement element, string name) =>
            (string)element.Attribute(name) ?? throw new FormatException("Missing rule pack attribute: " + name + ".");
        private static double Number(XElement element, string name) =>
            double.Parse(Attr(element, name), NumberStyles.Float, CultureInfo.InvariantCulture);
        private static double OptionalNumber(XElement element, string name, double fallback)
        {
            var value = (string)element.Attribute(name);
            return string.IsNullOrWhiteSpace(value) ? fallback :
                double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
        }
        private static int OptionalInteger(XElement element, string name, int fallback)
        {
            var value = (string)element.Attribute(name);
            return string.IsNullOrWhiteSpace(value) ? fallback :
                int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
        }
        private static Tuple<double, double> LineRatioRange(XElement element, string minimumName,
            string maximumName, double fallbackMinimum, double fallbackMaximum)
        {
            var minimum = OptionalNumber(element, minimumName, fallbackMinimum);
            var maximum = OptionalNumber(element, maximumName, fallbackMaximum);
            if (minimum < .1d || maximum > 2d || minimum > maximum)
                throw new FormatException("Invalid signed Line Shape ratio range: " + minimumName + "/" + maximumName + ".");
            return Tuple.Create(minimum, maximum);
        }
        private static DateTimeOffset Time(XElement element, string name) =>
            DateTimeOffset.ParseExact(Attr(element, name), "O", CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal);
    }
}
