using System;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Text
{
    /// <summary>How a line break between two lines should be resolved.</summary>
    public enum LineBreakDecision
    {
        /// <summary>A real paragraph or list boundary: leave the break.</summary>
        Keep,

        /// <summary>A layout wrap inside a sentence: replace the break by one space.</summary>
        JoinWithSpace,

        /// <summary>A wrap inside a token such as "4869/QĐ-|UBND": remove the break.</summary>
        JoinWithoutSpace
    }

    /// <summary>
    /// Recognises line breaks that only reflect the line width of the source
    /// (typically text copied from a PDF) as opposed to paragraph, list and
    /// heading boundaries. Conservative by design: anything that may start a
    /// new item, or follows the end of a sentence, keeps its break.
    /// </summary>
    public static class BrokenLineJoiner
    {
        private static readonly TimeSpan Timeout = TimeSpan.FromMilliseconds(200);

        // "- nội dung", "+ nội dung", "• nội dung"
        private static readonly Regex BulletStart = new Regex(
            @"^[-–—+•*·▪◦]\s", RegexOptions.CultureInvariant, Timeout);

        // "1. ", "1.2. ", "2) ", "II. ", "a) ", "đ) ". Item numbers have at most
        // three digits, so a wrapped "năm|2026. Giao..." is not taken for one.
        private static readonly Regex NumberedStart = new Regex(
            @"^(?:\d{1,3}(?:\.\d{1,3})*[.)]|[IVXLCDM]+[.)]|[a-zđ]\))(?:\s|$)",
            RegexOptions.CultureInvariant, Timeout);

        // "Điều 1.", "Điều 5a.", "Chương II", "Mục 1:", "Phần thứ nhất", "Phụ lục I"
        private static readonly Regex LegalHeadingStart = new Regex(
            @"^(?:Điều\s+\d+[a-zđ]?\.|(?:Chương|Mục|Phần|Phụ\s+lục)\s+(?:[IVXLCDM]+|\d+|thứ\s+\S+)(?:\s*[.:]|\s*$))",
            RegexOptions.CultureInvariant, Timeout);

        // Each legal basis, consideration or salutation is its own paragraph even
        // when the previous one lacks its closing ";" or ":".
        private static readonly Regex FormulaStart = new Regex(
            @"^(?:Căn cứ|Xét|Theo đề nghị|Kính gửi|Nơi nhận)\b",
            RegexOptions.CultureInvariant, Timeout);

        private const string TerminalPunctuation = ".;:!?…";
        private const string ClosingMarks = "\"'”’»)]";
        private const string LeadingPunctuation = ",.;:!?)]”’»";

        /// <summary>
        /// Decides whether the break between <paramref name="previousLine"/> and
        /// <paramref name="nextLine"/> should be removed. Both lines are given
        /// without the break character itself.
        /// </summary>
        public static LineBreakDecision Decide(string? previousLine, string? nextLine)
        {
            var previous = TrimLineEnd(previousLine ?? string.Empty);
            var next = TrimLineStart(nextLine ?? string.Empty);
            if (previous.Length == 0 || next.Length == 0) return LineBreakDecision.Keep;
            if (EndsSentence(previous)) return LineBreakDecision.Keep;
            if (StartsNewItem(next)) return LineBreakDecision.Keep;
            if (IsUppercaseHeading(previous) || IsUppercaseHeading(next)) return LineBreakDecision.Keep;

            if (LeadingPunctuation.IndexOf(next[0]) >= 0) return LineBreakDecision.JoinWithoutSpace;
            var last = previous[previous.Length - 1];
            if ((last == '-' || last == '/') && previous.Length > 1 && !IsBlank(previous[previous.Length - 2]))
                return LineBreakDecision.JoinWithoutSpace;
            // A line only wraps once it reaches the margin. A short line followed
            // by a capitalised one is a heading ("1. Mục tiêu") or a deliberate
            // break; a lowercase continuation is still joined.
            if (previous.Length < MinimumWrappedLineLength && char.IsUpper(next[0]))
                return LineBreakDecision.Keep;
            return LineBreakDecision.JoinWithSpace;
        }

        /// <summary>
        /// Lines shorter than this did not reach the right margin of an ordinary
        /// A4 page, so their break is not a wrap unless the next line plainly
        /// continues the sentence.
        /// </summary>
        public const int MinimumWrappedLineLength = 40;

        /// <summary>Number of blank characters (space, tab, no-break space) at the end of a line.</summary>
        public static int TrailingBlankCount(string line)
        {
            if (line == null) throw new ArgumentNullException(nameof(line));
            var count = 0;
            while (count < line.Length && IsBlank(line[line.Length - 1 - count])) count++;
            return count;
        }

        /// <summary>Number of blank characters (space, tab, no-break space) at the start of a line.</summary>
        public static int LeadingBlankCount(string line)
        {
            if (line == null) throw new ArgumentNullException(nameof(line));
            var count = 0;
            while (count < line.Length && IsBlank(line[count])) count++;
            return count;
        }

        private static bool EndsSentence(string line)
        {
            var index = line.Length - 1;
            while (index >= 0 && ClosingMarks.IndexOf(line[index]) >= 0) index--;
            return index >= 0 && TerminalPunctuation.IndexOf(line[index]) >= 0;
        }

        private static bool StartsNewItem(string line) =>
            BulletStart.IsMatch(line) || NumberedStart.IsMatch(line) || LegalHeadingStart.IsMatch(line) ||
            FormulaStart.IsMatch(line);

        /// <summary>A line whose letters are all capitals ("QUYẾT ĐỊNH", "CHƯƠNG I").</summary>
        private static bool IsUppercaseHeading(string line)
        {
            var letters = 0;
            foreach (var character in line)
            {
                if (!char.IsLetter(character)) continue;
                if (char.IsLower(character)) return false;
                letters++;
            }
            return letters >= 2;
        }

        private static string TrimLineEnd(string line) => line.Substring(0, line.Length - TrailingBlankCount(line));

        private static string TrimLineStart(string line) => line.Substring(LeadingBlankCount(line));

        private static bool IsBlank(char character) => character == ' ' || character == '\t' || character == ' ';
    }
}
