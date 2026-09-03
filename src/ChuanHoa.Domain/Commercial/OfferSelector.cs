using ChuanHoa.Domain.Common;

namespace ChuanHoa.Domain.Commercial;

public sealed class OfferSelector
{
    public OfferVersion SelectCurrent(
        IEnumerable<OfferVersion> offers,
        string productCode,
        string audience,
        string channel,
        string currency,
        DateTimeOffset serverNowUtc)
    {
        var matches = offers
            .Where(offer => offer.ProductCode == productCode)
            .Where(offer => offer.Audience == audience)
            .Where(offer => offer.Channel == channel)
            .Where(offer => offer.Currency == currency)
            .Where(offer => offer.Status == OfferStatus.Published)
            .Where(offer => offer.EffectiveFromUtc <= serverNowUtc)
            .Where(offer => offer.EffectiveUntilUtc is null || serverNowUtc < offer.EffectiveUntilUtc)
            .OrderByDescending(offer => offer.Version)
            .ToArray();

        return matches.Length switch
        {
            1 => matches[0],
            0 => throw new DomainException("OFFER_NOT_AVAILABLE", "No published offer is effective for the resolved audience."),
            _ => throw new DomainException("OFFER_EFFECTIVE_RANGE_OVERLAP", "More than one published offer is effective for the resolved audience.")
        };
    }
}
