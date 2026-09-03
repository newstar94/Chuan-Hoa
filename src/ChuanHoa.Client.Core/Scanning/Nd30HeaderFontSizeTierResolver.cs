using System;
using System.Collections.Generic;
using System.Linq;

namespace ChuanHoa.Client.Core.Scanning
{
    public sealed class Nd30HeaderFontSizeTier
    {
        public Nd30HeaderFontSizeTier(double nationalTitle, double nationalMotto, double placeAndIssuedDate)
        {
            NationalTitle = nationalTitle;
            NationalMotto = nationalMotto;
            PlaceAndIssuedDate = placeAndIssuedDate;
        }

        public double NationalTitle { get; }
        public double NationalMotto { get; }
        public double PlaceAndIssuedDate { get; }
    }

    public static class Nd30HeaderFontSizeTierResolver
    {
        private static readonly Nd30HeaderFontSizeTier Small = new Nd30HeaderFontSizeTier(12d, 13d, 13d);
        private static readonly Nd30HeaderFontSizeTier Large = new Nd30HeaderFontSizeTier(13d, 14d, 14d);

        public static Nd30HeaderFontSizeTier Resolve(LocalScanSnapshot snapshot,
            IDictionary<int, string> roles)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (roles == null) throw new ArgumentNullException(nameof(roles));

            var nationalTitle = FirstSize(snapshot, roles, "nationalTitle");
            if (Is(nationalTitle, 13d)) return Large;
            if (Is(nationalTitle, 12d)) return Small;

            var largeVotes = 0;
            var smallVotes = 0;
            Vote(FirstSize(snapshot, roles, "nationalMotto"), ref smallVotes, ref largeVotes);
            Vote(FirstSize(snapshot, roles, "placeAndIssuedDate"), ref smallVotes, ref largeVotes);
            return largeVotes > smallVotes ? Large : Small;
        }

        private static double? FirstSize(LocalScanSnapshot snapshot, IDictionary<int, string> roles,
            string role)
        {
            return snapshot.Paragraphs
                .Where(paragraph => paragraph.FontSizePoints.HasValue &&
                    roles.TryGetValue(paragraph.Index, out var detectedRole) &&
                    string.Equals(detectedRole, role, StringComparison.Ordinal))
                .OrderBy(paragraph => paragraph.Index)
                .Select(paragraph => paragraph.FontSizePoints)
                .FirstOrDefault();
        }

        private static void Vote(double? size, ref int smallVotes, ref int largeVotes)
        {
            if (Is(size, 14d)) largeVotes++;
            else if (Is(size, 13d)) smallVotes++;
        }

        private static bool Is(double? actual, double expected) =>
            actual.HasValue && Math.Abs(actual.Value - expected) <= .1d;
    }
}
