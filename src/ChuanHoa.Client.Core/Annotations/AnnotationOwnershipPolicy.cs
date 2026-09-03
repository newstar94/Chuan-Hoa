using System;
using System.Linq;

namespace ChuanHoa.Client.Core.Annotations
{
    public static class AnnotationOwnershipPolicy
    {
        public const int WordRedColor = 255;

        public static bool IsOwnedComment(string commentText, string lane)
        {
            if (commentText == null || string.IsNullOrWhiteSpace(lane))
            {
                return false;
            }

            return commentText.StartsWith(
                "[CHUẨN HÓA:" + lane.ToUpperInvariant() + ":",
                StringComparison.Ordinal);
        }

        public static string NormalizeLane(string lane)
        {
            if (string.IsNullOrWhiteSpace(lane))
            {
                throw new ArgumentException("An annotation lane is required.", nameof(lane));
            }

            var characters = lane.Where(char.IsLetterOrDigit).Take(12).ToArray();
            if (characters.Length == 0)
            {
                throw new ArgumentException("The annotation lane has no safe identifier characters.", nameof(lane));
            }

            return new string(characters).ToUpperInvariant();
        }

        public static int ColorAfterClearing(int currentColor, int originalColor)
        {
            return currentColor == WordRedColor ? originalColor : currentColor;
        }
    }
}
