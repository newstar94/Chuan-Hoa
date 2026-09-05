using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

    public static class AcademicTypographyRuleCodes
    {
        public const string SectionStyle = "LATEX-SEC-STYLE";
        public const string SectionContinuity = "LATEX-SEC-CONTINUITY";
        public const string PaginationKeep = "LATEX-PAGINATION-KEEP";
        public const string PaginationWidow = "LATEX-PAGINATION-WIDOW";
        public const string TableBooktabs = "LATEX-TABLE-BOOKTABS";
        public const string CaptionPosition = "LATEX-CAPTION-POS";
        public const string MathSyntax = "LATEX-MATH-SYNTAX";

        private static readonly IReadOnlyList<string> CanonicalCodes = Array.AsReadOnly(new[]
        {
            SectionStyle,
            SectionContinuity,
            PaginationKeep,
            PaginationWidow,
            TableBooktabs,
            CaptionPosition,
            MathSyntax
        });

        private static readonly HashSet<string> CanonicalCodeSet =
            new HashSet<string>(CanonicalCodes, StringComparer.Ordinal);

        private static readonly HashSet<string> AutoFixCodeSet = new HashSet<string>(new[]
        {
            PaginationKeep,
            PaginationWidow
        }, StringComparer.Ordinal);

        public static IReadOnlyList<string> All => CanonicalCodes;

        public static bool IsCanonical(string code) =>
            !string.IsNullOrWhiteSpace(code) && CanonicalCodeSet.Contains(code);

        public static bool CanAutoFix(string code) =>
            !string.IsNullOrWhiteSpace(code) && AutoFixCodeSet.Contains(code);
    }

    public enum AdvisoryProfileStatus
    {
        DisabledNotConfigured = 0,
        DisabledByPolicy = 1,
        DisabledLegacyV1 = 2,
        DisabledMalformed = 3,
        DisabledUnsupportedPolicyVersion = 4,
        Enabled = 5
    }

    public enum AdvisoryProfileDiagnosticCode
    {
        AdvisoryProfileNotConfigured = 0,
        ProfileDisabledByPolicy = 1,
        LegacyRulePackHasNoAdvisoryProfile = 2,
        AdvisoryProfileMalformed = 3,
        DetectorPolicyVersionUnsupported = 4,
        ProfileEnabled = 5
    }

    public sealed class AcademicTypographyAdvisoryThresholds
    {
        public AcademicTypographyAdvisoryThresholds(
            double headingConfidenceMinimum,
            double bodyConfidenceMinimum,
            int captionMaxBlankParagraphs,
            int mathMinimumSignalCount)
        {
            HeadingConfidenceMinimum = headingConfidenceMinimum;
            BodyConfidenceMinimum = bodyConfidenceMinimum;
            CaptionMaxBlankParagraphs = captionMaxBlankParagraphs;
            MathMinimumSignalCount = mathMinimumSignalCount;
        }

        public double HeadingConfidenceMinimum { get; }
        public double BodyConfidenceMinimum { get; }
        public int CaptionMaxBlankParagraphs { get; }
        public int MathMinimumSignalCount { get; }

        internal static AcademicTypographyAdvisoryThresholds ConservativeDefaults { get; } =
            new AcademicTypographyAdvisoryThresholds(.90d, .95d, 1, 1);
    }

    public sealed class AcademicTypographyAdvisoryProfile
    {
        public const string ProfileCode = "AcademicTypography";
        public const int SupportedDetectorPolicyVersion = 1;

        internal AcademicTypographyAdvisoryProfile(
            AdvisoryProfileStatus status,
            AdvisoryProfileDiagnosticCode diagnosticCode,
            int detectorPolicyVersion,
            IReadOnlyList<string>? enabledRuleCodes = null,
            IReadOnlyList<string>? autoFixRuleCodes = null,
            AcademicTypographyAdvisoryThresholds? thresholds = null)
        {
            Status = status;
            DiagnosticCode = diagnosticCode;
            DetectorPolicyVersion = detectorPolicyVersion;
            EnabledRuleCodes = status == AdvisoryProfileStatus.Enabled
                ? enabledRuleCodes ?? Array.Empty<string>()
                : Array.Empty<string>();
            AutoFixRuleCodes = status == AdvisoryProfileStatus.Enabled
                ? autoFixRuleCodes ?? Array.Empty<string>()
                : Array.Empty<string>();
            Thresholds = thresholds ?? AcademicTypographyAdvisoryThresholds.ConservativeDefaults;
        }

        public string Code => ProfileCode;
        public bool Enabled => Status == AdvisoryProfileStatus.Enabled;
        public AdvisoryProfileStatus Status { get; }
        public AdvisoryProfileDiagnosticCode DiagnosticCode { get; }
        public int DetectorPolicyVersion { get; }
        public IReadOnlyList<string> EnabledRuleCodes { get; }
        public IReadOnlyList<string> AutoFixRuleCodes { get; }
        public AcademicTypographyAdvisoryThresholds Thresholds { get; }

        public bool IsRuleEnabled(string ruleCode) =>
            Enabled && EnabledRuleCodes.Contains(ruleCode, StringComparer.Ordinal);

        public bool IsAutoFixEnabled(string ruleCode) =>
            Enabled && AutoFixRuleCodes.Contains(ruleCode, StringComparer.Ordinal);

        internal static AcademicTypographyAdvisoryProfile LegacyV1() =>
            Disabled(AdvisoryProfileStatus.DisabledLegacyV1,
                AdvisoryProfileDiagnosticCode.LegacyRulePackHasNoAdvisoryProfile);

        internal static AcademicTypographyAdvisoryProfile NotConfigured() =>
            Disabled(AdvisoryProfileStatus.DisabledNotConfigured,
                AdvisoryProfileDiagnosticCode.AdvisoryProfileNotConfigured);

        internal static AcademicTypographyAdvisoryProfile Malformed() =>
            Disabled(AdvisoryProfileStatus.DisabledMalformed,
                AdvisoryProfileDiagnosticCode.AdvisoryProfileMalformed);

        internal static AcademicTypographyAdvisoryProfile UnsupportedPolicyVersion(int detectorPolicyVersion) =>
            Disabled(AdvisoryProfileStatus.DisabledUnsupportedPolicyVersion,
                AdvisoryProfileDiagnosticCode.DetectorPolicyVersionUnsupported, detectorPolicyVersion);

        private static AcademicTypographyAdvisoryProfile Disabled(
            AdvisoryProfileStatus status,
            AdvisoryProfileDiagnosticCode diagnosticCode,
            int detectorPolicyVersion = 0) =>
            new AcademicTypographyAdvisoryProfile(status, diagnosticCode, detectorPolicyVersion);
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
             IReadOnlyList<string>? lexicon = null,
             AcademicTypographyAdvisoryProfile? academicTypography = null)
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
             AcademicTypography = academicTypography ?? AcademicTypographyAdvisoryProfile.NotConfigured();
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
         public AcademicTypographyAdvisoryProfile AcademicTypography { get; }
     }

     public static class LocalRulePackParser
     {
         public const string Schema = "chuanhoa.local-rule-pack.v2";
         public const string LegacySchema = "chuanhoa.local-rule-pack.v1";
         public static LocalRulePack Parse(byte[] payload, DateTimeOffset nowUtc, string clientReleaseId)
         {
             if (payload == null || payload.Length == 0) throw new FormatException("Rule pack payload is empty.");
             var root = XElement.Parse(Encoding.UTF8.GetString(payload), LoadOptions.None);
             if (root.Name != "rulePack")
                 throw new FormatException("Rule pack schema is unsupported.");
             var schema = Attr(root, "schema");
             if (schema != Schema && schema != LegacySchema)
                 throw new FormatException("Rule pack schema is unsupported.");
             if (schema == Schema) ValidateV2CoreShape(root);
             var notBefore = Time(root, "notBeforeUtc");
             var expires = Time(root, "expiresAtUtc");
             if (notBefore >= expires) throw new FormatException("Rule pack validity window is invalid.");
             if (nowUtc < notBefore) throw new InvalidOperationException("RULE_PACK_NOT_ACTIVE");
             if (nowUtc >= expires) throw new InvalidOperationException("RULE_PACK_EXPIRED");
             var minimumClient = Attr(root, "minimumClientReleaseId");
             if (CompareVersion(clientReleaseId, minimumClient, schema == Schema) < 0)
                 throw new InvalidOperationException("RULE_PACK_CLIENT_TOO_OLD");
            var format = root.Element("format") ?? throw new FormatException("Format rules are missing.");
            var corrections = root.Element("corrections")?.Elements("correction")
                .Select(item => new TextCorrectionRule(Attr(item, "wrong"), Attr(item, "replacement"))).ToArray()
                ?? Array.Empty<TextCorrectionRule>();
             var telex = root.Element("telex")?.Elements("rule")
                 .Select(item => new TelexRule(Attr(item, "pattern"), Attr(item, "replacement"))).ToArray()
                 ?? Array.Empty<TelexRule>();
             var hidden = root.Element("hiddenCharacters")?.Elements("character")
                 .Select(ParseCharacter).ToArray()
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
             var academicTypography = schema == LegacySchema
                 ? AcademicTypographyAdvisoryProfile.LegacyV1()
                 : ParseAcademicTypographyBoundary(root);
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
                 lexicon, academicTypography);
         }

         private static AcademicTypographyAdvisoryProfile ParseAcademicTypographyBoundary(XElement root)
         {
             try
             {
                 var containers = root.Elements("advisoryProfiles").ToArray();
                 if (containers.Length == 0) return AcademicTypographyAdvisoryProfile.NotConfigured();
                 if (containers.Length != 1) throw new FormatException("Duplicate advisory profile containers.");
                 return ParseAcademicTypography(containers[0]);
             }
             catch (Exception exception) when (
                 exception is FormatException ||
                 exception is InvalidOperationException ||
                 exception is ArgumentException ||
                 exception is OverflowException)
             {
                 return AcademicTypographyAdvisoryProfile.Malformed();
             }
         }

         private static AcademicTypographyAdvisoryProfile ParseAcademicTypography(XElement container)
         {
             ValidateAttributes(container, Array.Empty<string>());
             var profiles = container.Elements().ToArray();
             if (profiles.Length != 1 || profiles[0].Name.LocalName != "profile")
                 throw new FormatException("AcademicTypography advisory profile is missing or ambiguous.");

             var profile = profiles[0];
             ValidateAttributes(profile, new[] { "code", "enabled", "detectorPolicyVersion" });
             if (Attr(profile, "code") != AcademicTypographyAdvisoryProfile.ProfileCode)
                 throw new FormatException("Advisory profile code is unsupported.");

             bool enabled;
             if (!bool.TryParse(Attr(profile, "enabled"), out enabled))
                 throw new FormatException("Advisory profile enabled flag is invalid.");
             var detectorPolicyVersion = Integer(profile, "detectorPolicyVersion");

             var enabledRules = ParseRuleCodes(RequiredSingleElement(profile, "enabledRules"), false);
             var autoFixRules = ParseRuleCodes(RequiredSingleElement(profile, "autoFixRules"), true);
             if (autoFixRules.Any(code => !enabledRules.Contains(code, StringComparer.Ordinal)))
                 throw new FormatException("Advisory auto-fix rule is not enabled.");

             var thresholdElement = RequiredSingleElement(profile, "thresholds");
             ValidateNoUnexpectedChildren(profile, new[] { "enabledRules", "autoFixRules", "thresholds" });
             ValidateAttributes(thresholdElement, new[]
             {
                 "headingConfidenceMinimum", "bodyConfidenceMinimum",
                 "captionMaxBlankParagraphs", "mathMinimumSignalCount"
             });
             if (thresholdElement.Elements().Any())
                 throw new FormatException("Advisory thresholds cannot contain child elements.");
             var thresholds = new AcademicTypographyAdvisoryThresholds(
                 BoundedNumber(thresholdElement, "headingConfidenceMinimum", 0d, 1d),
                 BoundedNumber(thresholdElement, "bodyConfidenceMinimum", 0d, 1d),
                 BoundedInteger(thresholdElement, "captionMaxBlankParagraphs", 0, 2),
                 BoundedInteger(thresholdElement, "mathMinimumSignalCount", 1, 10));

             if (detectorPolicyVersion != AcademicTypographyAdvisoryProfile.SupportedDetectorPolicyVersion)
                 return AcademicTypographyAdvisoryProfile.UnsupportedPolicyVersion(detectorPolicyVersion);

             return new AcademicTypographyAdvisoryProfile(
                 enabled ? AdvisoryProfileStatus.Enabled : AdvisoryProfileStatus.DisabledByPolicy,
                 enabled ? AdvisoryProfileDiagnosticCode.ProfileEnabled :
                     AdvisoryProfileDiagnosticCode.ProfileDisabledByPolicy,
                 detectorPolicyVersion,
                 enabledRules,
                 autoFixRules,
                 thresholds);
         }

         private static IReadOnlyList<string> ParseRuleCodes(XElement container, bool autoFix)
         {
             ValidateAttributes(container, Array.Empty<string>());
             if (container.Elements().Any(item => item.Name.LocalName != "rule"))
                 throw new FormatException("Advisory rule list contains an unsupported element.");
             var result = new List<string>();
             var seen = new HashSet<string>(StringComparer.Ordinal);
             foreach (var rule in container.Elements("rule"))
             {
                 ValidateAttributes(rule, new[] { "code" });
                 if (rule.Elements().Any()) throw new FormatException("Advisory rule cannot contain child elements.");
                 var code = Attr(rule, "code");
                 if (!AcademicTypographyRuleCodes.IsCanonical(code))
                     throw new FormatException("Unknown AcademicTypography rule code.");
                 if (autoFix && !AcademicTypographyRuleCodes.CanAutoFix(code))
                     throw new FormatException("AcademicTypography auto-fix rule is not permitted.");
                 if (!seen.Add(code)) throw new FormatException("Duplicate AcademicTypography rule code.");
                 result.Add(code);
             }
             return result.AsReadOnly();
         }

         private static void ValidateV2CoreShape(XElement root)
         {
             ValidateAttributes(root, new[]
             {
                 "schema", "packId", "version", "notBeforeUtc", "expiresAtUtc", "minimumClientReleaseId"
             });
             ValidateNoUnexpectedChildren(root, new[]
             {
                 "format", "corrections", "lexicon", "telex", "hiddenCharacters",
                 "capitalizations", "documentTypeAbbreviations", "advisoryProfiles"
             });
             RequiredSingleElement(root, "format");
             RejectDuplicateCoreContainer(root, "corrections");
             RejectDuplicateCoreContainer(root, "lexicon");
             RejectDuplicateCoreContainer(root, "telex");
             RejectDuplicateCoreContainer(root, "hiddenCharacters");
             RejectDuplicateCoreContainer(root, "capitalizations");
             RejectDuplicateCoreContainer(root, "documentTypeAbbreviations");

             NonBlankAttr(root, "packId");
             NumericVersionAttr(root, "version");
             NonBlankAttr(root, "minimumClientReleaseId");
             ValidateFormatElement(RequiredSingleElement(root, "format"));
             ValidateRepeatedElementContainer(root, "corrections", "correction", new[] { "wrong", "replacement" });
             ValidateRepeatedElementContainer(root, "lexicon", "word", new[] { "value" });
             ValidateRepeatedElementContainer(root, "telex", "rule", new[] { "pattern", "replacement" });
             ValidateRepeatedElementContainer(root, "hiddenCharacters", "character", new[] { "codePoint" });
             ValidateRepeatedElementContainer(root, "capitalizations", "entry", new[] { "category", "expected" });
             ValidateRepeatedElementContainer(root, "documentTypeAbbreviations", "entry", new[] { "typeName", "abbreviation" });
         }

         private static void ValidateFormatElement(XElement format)
         {
             ValidateAttributes(format, new[]
             {
                 "a4WidthMm", "a4HeightMm", "topMinMm", "topMaxMm", "bottomMinMm", "bottomMaxMm",
                 "leftMinMm", "leftMaxMm", "rightMinMm", "rightMaxMm", "bodyFontName",
                 "bodyFontMinPoints", "bodyFontMaxPoints", "bodyFirstLineIndentMinMm",
                 "bodyFirstLineIndentMaxMm", "bodySpaceAfterMinPoints", "bodyAlignment",
                 "mottoLineMinRatio", "mottoLineMaxRatio", "organLineMinRatio", "organLineMaxRatio",
                 "subjectLineMinRatio", "subjectLineMaxRatio", "partyTitleLineMinRatio",
                 "partyTitleLineMaxRatio"
             });
             if (format.Elements().Any()) throw new FormatException("Format rules cannot contain child elements.");
         }

         private static void ValidateRepeatedElementContainer(
             XElement root,
             string containerName,
             string itemName,
             IReadOnlyList<string> itemAttributes)
         {
             var container = root.Element(containerName);
             if (container == null) return;
             ValidateAttributes(container, Array.Empty<string>());
             if (container.Elements().Any(item => item.Name != itemName))
                 throw new FormatException("Unsupported rule pack element in " + containerName + ".");
             foreach (var item in container.Elements(itemName))
             {
                 ValidateAttributes(item, itemAttributes);
                 if (item.Elements().Any())
                     throw new FormatException("Core rule pack item cannot contain child elements: " + itemName + ".");
                 foreach (var attribute in itemAttributes) NonBlankAttr(item, attribute);
                 if (containerName == "telex") ValidateRegularExpression(Attr(item, "pattern"));
             }
         }

         private static void ValidateRegularExpression(string pattern)
         {
             try
             {
                 _ = new Regex(pattern, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
             }
             catch (ArgumentException exception)
             {
                 throw new FormatException("Signed Telex rule contains an invalid regular expression.", exception);
             }
         }

         private static void RejectDuplicateCoreContainer(XElement root, string name)
         {
             if (root.Elements(name).Skip(1).Any())
                 throw new FormatException("Duplicate core rule pack element: " + name + ".");
         }

         private static XElement RequiredSingleElement(XElement parent, string name)
         {
             var matches = parent.Elements(name).ToArray();
             if (matches.Length != 1) throw new FormatException("Rule pack element must occur exactly once: " + name + ".");
             return matches[0];
         }

         private static void ValidateNoUnexpectedChildren(XElement parent, IEnumerable<string> allowedNames)
         {
             var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
             var unexpected = parent.Elements().FirstOrDefault(item =>
                 !string.IsNullOrEmpty(item.Name.NamespaceName) || !allowed.Contains(item.Name.LocalName));
             if (unexpected != null)
                 throw new FormatException("Unsupported rule pack element: " + unexpected.Name.LocalName + ".");
         }

         private static void ValidateAttributes(XElement element, IEnumerable<string> allowedNames)
         {
             var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
             var unexpected = element.Attributes().FirstOrDefault(item =>
                 item.IsNamespaceDeclaration || !string.IsNullOrEmpty(item.Name.NamespaceName) ||
                 !allowed.Contains(item.Name.LocalName));
             if (unexpected != null)
                 throw new FormatException("Unsupported rule pack attribute: " + unexpected.Name.LocalName + ".");
         }

         private static string NonBlankAttr(XElement element, string name)
         {
             var value = Attr(element, name);
             if (string.IsNullOrWhiteSpace(value))
                 throw new FormatException("Rule pack attribute cannot be blank: " + name + ".");
             return value;
         }

         private static Version NumericVersionAttr(XElement element, string name)
         {
             var value = NonBlankAttr(element, name);
             Version version;
             if (!Version.TryParse(value, out version))
                 throw new FormatException("Rule pack version attribute is invalid: " + name + ".");
             return version;
         }

         private static int CompareVersion(string actual, string required, bool requireVersionSyntax)
         {
             Version left, right;
             if (Version.TryParse(actual, out left) && Version.TryParse(required, out right))
                 return left.CompareTo(right);
             if (requireVersionSyntax)
                 throw new FormatException("Rule pack client release compatibility value is invalid.");
             return string.CompareOrdinal(actual, required);
         }
         private static char ParseCharacter(XElement element)
         {
             var value = int.Parse(Attr(element, "codePoint"), NumberStyles.HexNumber,
                 CultureInfo.InvariantCulture);
             if (value < char.MinValue || value > char.MaxValue || char.IsSurrogate((char)value))
                 throw new FormatException("Hidden-character code point is outside the supported UTF-16 range.");
             return (char)value;
         }
        private static string Attr(XElement element, string name) =>
            (string)element.Attribute(name) ?? throw new FormatException("Missing rule pack attribute: " + name + ".");
         private static double Number(XElement element, string name) =>
             Finite(double.Parse(Attr(element, name), NumberStyles.Float, CultureInfo.InvariantCulture), name);
        private static double OptionalNumber(XElement element, string name, double fallback)
        {
            var value = (string)element.Attribute(name);
             return string.IsNullOrWhiteSpace(value) ? fallback :
                 Finite(double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture), name);
        }
         private static int OptionalInteger(XElement element, string name, int fallback)
        {
            var value = (string)element.Attribute(name);
            return string.IsNullOrWhiteSpace(value) ? fallback :
                int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
         }
         private static int Integer(XElement element, string name) =>
             int.Parse(Attr(element, name), NumberStyles.Integer, CultureInfo.InvariantCulture);
         private static double BoundedNumber(XElement element, string name, double minimum, double maximum)
         {
             var value = Number(element, name);
             if (value < minimum || value > maximum)
                 throw new FormatException("Rule pack number is outside its allowed range: " + name + ".");
             return value;
         }
         private static int BoundedInteger(XElement element, string name, int minimum, int maximum)
         {
             var value = Integer(element, name);
             if (value < minimum || value > maximum)
                 throw new FormatException("Rule pack integer is outside its allowed range: " + name + ".");
             return value;
         }
         private static double Finite(double value, string name)
         {
             if (double.IsNaN(value) || double.IsInfinity(value))
                 throw new FormatException("Rule pack number must be finite: " + name + ".");
             return value;
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
