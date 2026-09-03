using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace ChuanHoa.Api.Development;

public sealed record DevelopmentBootstrapRequest(string DeviceThumbprint, string ClientReleaseId);

public sealed record DevelopmentBootstrapResponse(
    DateTimeOffset ServerTimeUtc,
    string Lease,
    string RulePack,
    DateTimeOffset LeaseExpiresAtUtc,
    DateTimeOffset RulePackExpiresAtUtc);

public sealed class DevelopmentArtifactIssuer
{
    public const string KeyId = "CHUANHOA-LOCAL-DEVELOPMENT-1";
    public const string LeaseKind = "offlineLease";
    public const string RulePackKind = "rulePack";
    private const string SignedArtifactSchema = "chuanhoa.signed-artifact.v1";
    private const string LeaseSchema = "chuanhoa.offline-lease.v1";
    private const string RulePackSchema = "chuanhoa.local-rule-pack.v1";
    private readonly IWebHostEnvironment _environment;
    private readonly IConfiguration _configuration;
    private readonly TimeProvider _timeProvider;

    public DevelopmentArtifactIssuer(
        IWebHostEnvironment environment,
        IConfiguration configuration,
        TimeProvider timeProvider)
    {
        _environment = environment;
        _configuration = configuration;
        _timeProvider = timeProvider;
    }

    public DevelopmentBootstrapResponse Issue(DevelopmentBootstrapRequest request)
    {
        if (!IsEnabled)
        {
            throw new InvalidOperationException("DEVELOPMENT_BOOTSTRAP_DISABLED");
        }
        if (request is null || string.IsNullOrWhiteSpace(request.DeviceThumbprint) ||
            string.IsNullOrWhiteSpace(request.ClientReleaseId))
        {
            throw new ArgumentException("Device thumbprint and client release id are required.", nameof(request));
        }

        var now = _timeProvider.GetUtcNow();
        var leaseExpires = now.AddDays(7);
        var packExpires = now.AddDays(30);
        using var rsa = LoadPrivateKey();
        var lease = new XElement("offlineLease",
            new XAttribute("schema", LeaseSchema),
            new XAttribute("leaseId", Guid.NewGuid().ToString("D")),
            new XAttribute("subjectId", "development-user"),
            new XAttribute("deviceThumbprint", request.DeviceThumbprint),
            new XAttribute("clientReleaseId", request.ClientReleaseId),
            new XAttribute("issuedAtUtc", now.ToString("O", CultureInfo.InvariantCulture)),
            new XAttribute("notBeforeUtc", now.AddMinutes(-1).ToString("O", CultureInfo.InvariantCulture)),
            new XAttribute("expiresAtUtc", leaseExpires.ToString("O", CultureInfo.InvariantCulture)),
            new XElement("features",
                new XElement("feature", new XAttribute("code", "FORMAT_SCAN")),
                new XElement("feature", new XAttribute("code", "SPELLING_SCAN")),
                new XElement("feature", new XAttribute("code", "DOCUMENT_TOOLS")),
                new XElement("feature", new XAttribute("code", "AUTOFIX"))));
        var rulePack = BuildRulePack(now, packExpires, request.ClientReleaseId);
        return new DevelopmentBootstrapResponse(
            now,
            Sign(rsa, LeaseKind, lease),
            Sign(rsa, RulePackKind, rulePack),
            leaseExpires,
            packExpires);
    }

    public bool IsEnabled => _environment.IsDevelopment() &&
        _configuration.GetValue<bool>("ChuanHoa:EnableDevelopmentBootstrap");

    private XElement BuildRulePack(DateTimeOffset now, DateTimeOffset expires, string clientReleaseId)
    {
        var corrections = LoadCorrections();
        var lexicon = LoadLexicon();
        return new XElement("rulePack",
            new XAttribute("schema", RulePackSchema),
            new XAttribute("packId", "LOCAL-DEVELOPMENT-ND30-V2"),
            new XAttribute("version", "1.1.0"),
            new XAttribute("notBeforeUtc", now.AddMinutes(-1).ToString("O", CultureInfo.InvariantCulture)),
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
                CapitalizationEntries().Select(item => new XElement("entry",
                    new XAttribute("category", item.Category), new XAttribute("expected", item.Expected)))),
            new XElement("documentTypeAbbreviations",
                DocumentTypeEntries().Select(item => new XElement("entry",
                    new XAttribute("typeName", item.TypeName), new XAttribute("abbreviation", item.Abbreviation)))));
    }

    private static IEnumerable<(string Category, string Expected)> CapitalizationEntries()
    {
        yield return ("administrative", "Hà Nội");
        yield return ("administrative", "Thành phố Hồ Chí Minh");
        yield return ("administrative", "Đà Nẵng");
        yield return ("administrative", "Hải Phòng");
        yield return ("administrative", "Cần Thơ");
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

    private IReadOnlyDictionary<string, string> LoadCorrections()
    {
        var configured = _configuration["ChuanHoa:DevelopmentRuleDictionaryPath"];
        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = Path.GetFullPath(Path.Combine(
                _environment.ContentRootPath, "..", "..", "shared", "dictionaries", "typo_dictionary.json"));
        }
        if (string.IsNullOrWhiteSpace(configured) || !File.Exists(configured))
        {
            throw new InvalidOperationException("Development rule dictionary is not configured.");
        }
        using var stream = File.OpenRead(configured);
        var data = JsonSerializer.Deserialize<Dictionary<string, string>>(stream)
            ?? throw new InvalidOperationException("Development rule dictionary is invalid.");
        return data
            .Where(item => !string.IsNullOrWhiteSpace(item.Key) && !string.IsNullOrWhiteSpace(item.Value) &&
                !string.Equals(item.Key, item.Value, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);
    }

    private IReadOnlyList<string> LoadLexicon()
    {
        var configured = _configuration["ChuanHoa:DevelopmentRuleLexiconDirectory"];
        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = Path.GetFullPath(Path.Combine(
                _environment.ContentRootPath, "..", "..", "shared", "dictionaries", "hunspell-vi"));
        }
        if (string.IsNullOrWhiteSpace(configured) || !Directory.Exists(configured))
            throw new InvalidOperationException("Development Vietnamese lexicon is not configured.");

        var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in Directory.EnumerateFiles(configured, "*.dic", SearchOption.TopDirectoryOnly))
        {
            foreach (var raw in File.ReadLines(path, Encoding.UTF8))
            {
                var value = raw.Trim();
                if (value.Length == 0 || value.All(char.IsDigit)) continue;
                var flag = value.IndexOf('/');
                if (flag >= 0) value = value.Substring(0, flag);
                if (value.Length > 0) words.Add(value.Normalize(NormalizationForm.FormC));
            }
        }
        if (words.Count < 1000)
            throw new InvalidOperationException("Development Vietnamese lexicon is incomplete.");
        return words.OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private RSA LoadPrivateKey()
    {
        var path = _configuration["ChuanHoa:DevelopmentSigningKeyPath"];
        if (string.IsNullOrWhiteSpace(path))
        {
            path = Path.GetFullPath(Path.Combine(
                _environment.ContentRootPath, "..", "..", ".dev-secrets", "development-signing-key.xml"));
        }
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new InvalidOperationException("Development signing key is not configured.");
        }
        var root = XElement.Load(path, LoadOptions.None);
        byte[] Read(string name) => Convert.FromBase64String(root.Element(name)?.Value
            ?? throw new InvalidOperationException("Development signing key is incomplete: " + name + "."));
        var rsa = RSA.Create();
        rsa.ImportParameters(new RSAParameters
        {
            Modulus = Read("Modulus"), Exponent = Read("Exponent"), D = Read("D"), P = Read("P"),
            Q = Read("Q"), DP = Read("DP"), DQ = Read("DQ"), InverseQ = Read("InverseQ")
        });
        return rsa;
    }

    private static string Sign(RSA rsa, string kind, XElement payloadElement)
    {
        var payload = Encoding.UTF8.GetBytes(payloadElement.ToString(SaveOptions.DisableFormatting));
        var signature = rsa.SignData(payload, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return new XDocument(new XElement("signedArtifact",
            new XAttribute("schema", SignedArtifactSchema), new XAttribute("kind", kind),
            new XAttribute("keyId", KeyId), new XAttribute("algorithm", "RS256"),
            new XElement("payload", Convert.ToBase64String(payload)),
            new XElement("signature", Convert.ToBase64String(signature))))
            .ToString(SaveOptions.DisableFormatting);
    }
}
