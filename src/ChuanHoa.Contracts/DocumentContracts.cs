using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

[JsonConverter(typeof(JsonStringEnumConverter<RegimeCode>))]
public enum RegimeCode
{
    Unknown,
    Nd30,
    PartyHd05,
    Viettel
}

[JsonConverter(typeof(JsonStringEnumConverter<DocumentTypeCode>))]
public enum DocumentTypeCode
{
    Unknown,
    OfficialLetter,
    Decision,
    Report,
    Plan,
    Submission,
    Regulation,
    Guidance,
    Notice,
    Minutes
}

[JsonConverter(typeof(JsonStringEnumConverter<ComponentRole>))]
public enum ComponentRole
{
    Unknown,
    NationalTitle,
    Motto,
    PartyTitle,
    ParentOrganName,
    OrganName,
    CodeNumberNotation,
    PlaceDate,
    TypeName,
    Subject,
    OfficialLetterSubject,
    LegalBasis,
    BodyText,
    PartChapterTitle,
    SectionTitle,
    Article,
    Clause,
    Point,
    SignAuthority,
    SignPosition,
    SignFullName,
    RecipientsSalutation,
    RecipientsLabel,
    RecipientListItem,
    RecipientArchiveLine,
    AppendixTitle,
    AppendixBody,
    Confidentiality,
    Urgency,
    TableContent
}

public sealed record DocumentSnapshot(
    int SchemaVersion,
    string DocumentFingerprint,
    long Revision,
    RegimeCode Regime,
    DocumentTypeCode DocumentType,
    bool RegimeWasSelectedManually,
    bool DocumentTypeWasSelectedManually,
    DocumentPreflight Preflight,
    IReadOnlyList<SectionSnapshot> Sections,
    IReadOnlyList<ParagraphSnapshot> Paragraphs,
    IReadOnlyList<TableSnapshot> Tables,
    IReadOnlyList<ProtectedSpanSnapshot>? ProtectedSpans = null);

public sealed record DocumentPreflight(
    bool IsReadOnly,
    bool IsProtected,
    bool IsCompatibilityMode,
    bool TrackChangesEnabled,
    string FileFormat,
    bool HasActiveWindow);

public sealed record SectionSnapshot(
    int Index,
    double PageWidthPoints,
    double PageHeightPoints,
    double TopMarginPoints,
    double BottomMarginPoints,
    double LeftMarginPoints,
    double RightMarginPoints,
    bool IsLandscape);

public sealed record ParagraphSnapshot(
    int Index,
    string Text,
    ComponentRole Role,
    double RoleConfidence,
    string? FontName,
    double? FontSizePoints,
    bool? Bold,
    bool? Italic,
    int? Alignment,
    double? FirstLineIndentPoints,
    double? SpaceBeforePoints,
    double? SpaceAfterPoints,
    bool IsInTable,
    string StoryType = "MainTextStory",
    int SectionIndex = 1,
    int AbsoluteStart = 0,
    int? TableIndex = null,
    int? RowIndex = null,
    int? CellIndex = null);

public sealed record ProtectedSpanSnapshot(
    string StoryType,
    int SectionIndex,
    int AbsoluteStart,
    int Length,
    string Kind);

public sealed record TableSnapshot(
    int Index,
    int RowCount,
    int ColumnCount,
    bool HasMergedCells,
    bool IsNested,
    IReadOnlyList<int> HeaderRowIndexes);
