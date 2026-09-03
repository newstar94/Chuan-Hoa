using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

[JsonConverter(typeof(JsonStringEnumConverter<FixOperationType>))]
public enum FixOperationType
{
    SetSectionPageSetup,
    InsertSectionBreak,
    RemoveVerifiedTrailingEmptyPage,
    UpsertStyle,
    ApplyComponentFormatting,
    SetParagraphKeepWithNext,
    SetPageNumbering,
    SetCharacterSpacing,
    SetTableHeaderRows,
    FormatTable,
    FormatImage,
    SetCellAlignment,
    ReplaceProtectedTextSpan,
    InsertQrImage
}

public sealed record FixOperation(
    Guid OperationId,
    FixOperationType Type,
    RiskTier RiskTier,
    string Scope,
    IReadOnlyDictionary<string, string> Preconditions,
    IReadOnlyDictionary<string, string> Parameters,
    IReadOnlyDictionary<string, string> ExpectedOutcome);

public sealed record FixPlanPayload(
    Guid FixPlanId,
    Guid SubjectId,
    Guid? OrganizationId,
    string DeviceKeyThumbprint,
    string CommandId,
    string DocumentFingerprint,
    long Revision,
    string ClientReleaseId,
    string ProtocolVersion,
    string RuleReleaseId,
    string EngineVersion,
    DateTimeOffset IssuedAtUtc,
    DateTimeOffset NotBeforeUtc,
    DateTimeOffset ExpiresAtUtc,
    string Nonce,
    string Jti,
    IReadOnlyList<FixOperation> Operations);
