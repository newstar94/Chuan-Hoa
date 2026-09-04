using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Lexicon
{
    public enum VietnameseTokenKind
    {
        Word,
        AcronymOrUpper,
        Number,
        Punctuation,
        Whitespace,
        UrlOrEmail,
        Other
    }

    public sealed class VietnameseTokenSpan
    {
        public VietnameseTokenSpan(string text, int startOffset, int length, int paragraphIndex, int sentenceIndex, VietnameseTokenKind kind)
        {
            Text = text ?? string.Empty;
            StartOffset = startOffset;
            Length = length;
            ParagraphIndex = paragraphIndex;
            SentenceIndex = sentenceIndex;
            Kind = kind;
        }

        public string Text { get; }
        public int StartOffset { get; }
        public int Length { get; }
        public int ParagraphIndex { get; }
        public int SentenceIndex { get; }
        public VietnameseTokenKind Kind { get; }

        public override string ToString() => $"{Kind}: '{Text}' [{StartOffset}..{StartOffset + Length}]";
    }

    /// <summary>
    /// Accurate, offset-preserving tokenizer for Vietnamese text inside Word documents.
    /// Preserves exact character offsets and distinguishes URLs, numbers, acronyms, and words.
    /// </summary>
    public static class VietnameseWordTokenizer
    {
        // Combined regex capturing URLs/Emails, Numbers, Words (with Vietnamese diacritics), Whitespace, and Punctuation
        private static readonly Regex TokenRegex = new Regex(
            @"(?<url>https?://[^\s/$.?#].[^\s]*|[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})|" +
            @"(?<space>[ \t\r\n]+)|" +
            @"(?<number>\d+(?:[.,]\d+)*)|" +
            @"(?<word>[a-zA-Z\u00C0-\u1EF9]+(?:[-'][a-zA-Z\u00C0-\u1EF9]+)*)|" +
            @"(?<punct>[,.;:!?()[\]{}/\\""“”—–…\-])|" +
            @"(?<other>[^\s])",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex SentenceBoundaryRegex = new Regex(@"[.!?]+(?:\s+|$)", RegexOptions.Compiled);

        public static IReadOnlyList<VietnameseTokenSpan> TokenizeParagraph(string text, int paragraphIndex = 0)
        {
            if (string.IsNullOrEmpty(text))
                return Array.Empty<VietnameseTokenSpan>();

            var tokens = new List<VietnameseTokenSpan>();
            var sentenceIndex = 0;

            var matches = TokenRegex.Matches(text);
            for (var i = 0; i < matches.Count; i++)
            {
                var match = matches[i];
                if (!match.Success || match.Length == 0) continue;

                var kind = VietnameseTokenKind.Other;
                if (match.Groups["url"].Success)
                    kind = VietnameseTokenKind.UrlOrEmail;
                else if (match.Groups["space"].Success)
                    kind = VietnameseTokenKind.Whitespace;
                else if (match.Groups["number"].Success)
                    kind = VietnameseTokenKind.Number;
                else if (match.Groups["punct"].Success)
                {
                    kind = VietnameseTokenKind.Punctuation;
                    if (match.Value == "." || match.Value == "!" || match.Value == "?")
                        sentenceIndex++;
                }
                else if (match.Groups["word"].Success)
                {
                    var word = match.Value;
                    if (word.Length > 1 && IsAllUpper(word))
                        kind = VietnameseTokenKind.AcronymOrUpper;
                    else
                        kind = VietnameseTokenKind.Word;
                }

                tokens.Add(new VietnameseTokenSpan(
                    match.Value,
                    match.Index,
                    match.Length,
                    paragraphIndex,
                    sentenceIndex,
                    kind));
            }

            return tokens;
        }

        private static bool IsAllUpper(string value)
        {
            for (var i = 0; i < value.Length; i++)
            {
                if (char.IsLetter(value[i]) && !char.IsUpper(value[i]))
                    return false;
            }
            return true;
        }
    }
}
