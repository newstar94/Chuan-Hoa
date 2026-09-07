namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>All pages here are physical pages, never restarted printed page numbers.</summary>
    public static class ComponentLineAnchorPolicy
    {
        public static bool CanAnchor(int ownerPage, int candidatePage, int ownerSection,
            int candidateSection, int candidateParagraph, int blockStart, int blockEnd) =>
            ownerPage > 0 && candidatePage == ownerPage && ownerSection == candidateSection &&
            candidateParagraph >= blockStart && candidateParagraph <= blockEnd;
    }
}
