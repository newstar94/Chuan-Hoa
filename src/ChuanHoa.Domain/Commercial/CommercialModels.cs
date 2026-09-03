namespace ChuanHoa.Domain.Commercial;

public enum OfferStatus
{
    Draft,
    Scheduled,
    Published,
    Retired
}

public sealed record OfferVersion(
    Guid Id,
    string ProductCode,
    int Version,
    string Audience,
    string Channel,
    string Currency,
    long AmountMinor,
    DateTimeOffset EffectiveFromUtc,
    DateTimeOffset? EffectiveUntilUtc,
    OfferStatus Status,
    IReadOnlySet<string> Features,
    TimeSpan PurchasedTerm);

public sealed record Quote(
    Guid Id,
    Guid SubjectId,
    Guid OfferId,
    int OfferVersion,
    string ProductCode,
    string Currency,
    long AmountMinor,
    IReadOnlySet<string> Features,
    TimeSpan PurchasedTerm,
    DateTimeOffset CreatedAtUtc,
    DateTimeOffset ExpiresAtUtc,
    string IntegrityBinding,
    bool Consumed);
