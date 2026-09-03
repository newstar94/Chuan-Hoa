using System;
using System.IO;

namespace ChuanHoa.Client.Core.Safety
{
    public static class SupportedWordDocumentFormatPolicy
    {
        // Word 97-2003 document (.doc).
        private const int WordBinaryDocument = 0;

        // Office Open XML document (.docx), transitional and strict.
        private const int WordOpenXmlDocument = 12;
        private const int WordStrictOpenXmlDocument = 24;

        public static bool IsSupported(string fullName, int wordSaveFormat)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return false;
            }

            var extension = Path.GetExtension(fullName);
            if (string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase))
            {
                return wordSaveFormat == WordBinaryDocument;
            }

            if (string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
            {
                return wordSaveFormat == WordOpenXmlDocument ||
                    wordSaveFormat == WordStrictOpenXmlDocument;
            }

            return false;
        }
    }
}
