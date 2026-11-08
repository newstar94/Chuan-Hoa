using System;
using System.Security.Cryptography;
using System.Text;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Identifies the words the spelling lane reads: the story and text of every
    /// paragraph, without positions, formatting or control marks. Comments and
    /// red highlights added by an annotation pass, and formatting changes, leave
    /// it unchanged; any edit to the text changes it.
    /// </summary>
    public static class SpellingTextFingerprint
    {
        public static string Compute(LocalScanSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var builder = new StringBuilder();
            foreach (var paragraph in snapshot.Paragraphs)
            {
                builder.Append(paragraph.StoryType).Append('\u001F');
                foreach (var character in paragraph.Text ?? string.Empty)
                {
                    // Comment references and other object marks are control
                    // characters; tabs, line and page breaks are part of the text.
                    if (character >= ' ' || character == '\t' || character == '\v' || character == '\f')
                        builder.Append(character);
                }
                builder.Append('\u001E');
            }
            using (var sha256 = SHA256.Create())
            {
                var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(builder.ToString()));
                return BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }
    }
}
