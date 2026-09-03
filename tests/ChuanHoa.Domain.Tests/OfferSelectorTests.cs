using ChuanHoa.Domain.Commercial;
using ChuanHoa.Domain.Common;

namespace ChuanHoa.Domain.Tests;

public sealed class OfferSelectorTests
{
    private readonly OfferSelector _selector = new();
    private readonly DateTimeOffset _now = DateTimeOffset.Parse("2026-10-01T00:00:00Z");

    [Fact]
    public void Selects_only_effective_published_offer_for_server_resolved_audience()
    {
        var offer = Offer(Guid.NewGuid(), 2, OfferStatus.Published, _now.AddDays(-1), _now.AddDays(1));
        var selected = _selector.SelectCurrent([offer], "CHUAN_HOA", "PERSONAL", "DIRECT", "VND", _now);
        Assert.Equal(offer.Id, selected.Id);
    }

    [Fact]
    public void Rejects_overlapping_published_offers()
    {
        var first = Offer(Guid.NewGuid(), 1, OfferStatus.Published, _now.AddDays(-1), _now.AddDays(2));
        var second = Offer(Guid.NewGuid(), 2, OfferStatus.Published, _now.AddDays(-2), _now.AddDays(1));
        var error = Assert.Throws<DomainException>(() =>
            _selector.SelectCurrent([first, second], "CHUAN_HOA", "PERSONAL", "DIRECT", "VND", _now));
        Assert.Equal("OFFER_EFFECTIVE_RANGE_OVERLAP", error.Code);
    }

    [Fact]
    public void Does_not_select_draft_offer()
    {
        var draft = Offer(Guid.NewGuid(), 1, OfferStatus.Draft, _now.AddDays(-1), _now.AddDays(1));
        var error = Assert.Throws<DomainException>(() =>
            _selector.SelectCurrent([draft], "CHUAN_HOA", "PERSONAL", "DIRECT", "VND", _now));
        Assert.Equal("OFFER_NOT_AVAILABLE", error.Code);
    }

    private static OfferVersion Offer(
        Guid id,
        int version,
        OfferStatus status,
        DateTimeOffset start,
        DateTimeOffset end)
    {
        return new OfferVersion(
            id,
            "CHUAN_HOA",
            version,
            "PERSONAL",
            "DIRECT",
            "VND",
            100_000,
            start,
            end,
            status,
            new HashSet<string> { "AUTOFIX" },
            TimeSpan.FromDays(365));
    }
}
