using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.Client.Core.Tests;

public sealed class ComponentLineAnchorPolicyTests
{
    [Theory]
    [InlineData(15, 1, 3, 1, 1, false)] // printed page 1 of a later section is not the cover
    [InlineData(15, 15, 3, 3, 520, true)]
    [InlineData(15, 0, 3, 3, 520, false)]
    [InlineData(0, 0, 3, 3, 520, false)]
    [InlineData(15, 15, 3, 2, 520, false)]
    [InlineData(15, 15, 3, 3, 499, false)]
    [InlineData(15, 15, 3, 3, 551, false)]
    public void Anchor_stays_on_known_physical_page_section_and_logical_document(
        int ownerPage, int page, int section, int candidateSection, int paragraph, bool expected) =>
        Assert.Equal(expected, ComponentLineAnchorPolicy.CanAnchor(ownerPage, page, section, candidateSection, paragraph, 500, 550));
}
