using System;
using System.Collections.Generic;
using System.Text;

namespace ChuanHoa.Client.Core.Text
{
    /// <summary>
    /// One contiguous change to a source string: the <see cref="Length"/>
    /// characters at <see cref="Offset"/> (equal to <see cref="ExpectedText"/>)
    /// are replaced by <see cref="Replacement"/>. Length 0 is an insertion.
    /// </summary>
    public sealed class TextSpanEdit
    {
        public TextSpanEdit(int offset, int length, string replacement, string expectedText)
        {
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
            if (length < 0) throw new ArgumentOutOfRangeException(nameof(length));
            Offset = offset;
            Length = length;
            Replacement = replacement ?? throw new ArgumentNullException(nameof(replacement));
            ExpectedText = expectedText ?? throw new ArgumentNullException(nameof(expectedText));
        }

        public int Offset { get; }
        public int Length { get; }
        public string Replacement { get; }
        public string ExpectedText { get; }
    }

    /// <summary>
    /// Turns "source → target" into the fewest character edits (Myers diff).
    /// Writing only these spans back to Word keeps the formatting, fields and
    /// anchors of every character that did not change; assigning the whole
    /// paragraph text instead gives all of it the formatting of its first
    /// character.
    /// </summary>
    public static class MinimalTextEditPlanner
    {
        /// <summary>Edit scripts longer than this are not planned (memory grows with its square).</summary>
        public const int DefaultMaximumEditDistance = 1024;

        /// <summary>
        /// Returns the edits ordered by offset, non-overlapping and separated by
        /// at least one unchanged character, or null when the edit distance
        /// exceeds <paramref name="maximumEditDistance"/>.
        /// </summary>
        public static IReadOnlyList<TextSpanEdit>? Plan(string source, string target,
            int maximumEditDistance = DefaultMaximumEditDistance)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (maximumEditDistance < 0) throw new ArgumentOutOfRangeException(nameof(maximumEditDistance));
            if (string.Equals(source, target, StringComparison.Ordinal)) return Array.Empty<TextSpanEdit>();

            // Only the differing middle is diffed.
            var shorter = Math.Min(source.Length, target.Length);
            var prefix = 0;
            while (prefix < shorter && source[prefix] == target[prefix]) prefix++;
            var suffix = 0;
            while (suffix < shorter - prefix &&
                   source[source.Length - 1 - suffix] == target[target.Length - 1 - suffix])
                suffix++;
            var a = source.Substring(prefix, source.Length - prefix - suffix);
            var b = target.Substring(prefix, target.Length - prefix - suffix);

            var operations = Diff(a, b, maximumEditDistance);
            if (operations == null) return null;
            var edits = Merge(operations, b, prefix, source);
            // The script is verified rather than trusted: a caller writes these
            // spans into a user's document.
            return string.Equals(ApplyTo(source, edits), target, StringComparison.Ordinal) ? edits : null;
        }

        /// <summary>Applies edits produced by <see cref="Plan"/> to <paramref name="source"/>.</summary>
        public static string ApplyTo(string source, IReadOnlyList<TextSpanEdit> edits)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (edits == null) throw new ArgumentNullException(nameof(edits));
            var builder = new StringBuilder(source);
            for (var index = edits.Count - 1; index >= 0; index--)
            {
                var edit = edits[index];
                builder.Remove(edit.Offset, edit.Length);
                builder.Insert(edit.Offset, edit.Replacement);
            }
            return builder.ToString();
        }

        private enum OperationKind : byte { Equal, Delete, Insert }

        private struct Operation
        {
            public Operation(OperationKind kind, int index)
            {
                Kind = kind;
                Index = index;
            }

            public OperationKind Kind { get; }

            /// <summary>Index in the source for Equal/Delete, in the target for Insert.</summary>
            public int Index { get; }
        }

        private static List<Operation>? Diff(string a, string b, int maximumEditDistance)
        {
            var n = a.Length;
            var m = b.Length;
            var max = n + m;
            if (max == 0) return new List<Operation>();
            var bound = Math.Min(max, maximumEditDistance);
            var offset = max + 1;
            var v = new int[2 * max + 3];
            // trace[d] holds V for diagonals -(d+1)..d+1 as it was before round d.
            var trace = new List<int[]>();
            for (var d = 0; d <= bound; d++)
            {
                var snapshot = new int[2 * d + 3];
                Array.Copy(v, offset - d - 1, snapshot, 0, snapshot.Length);
                trace.Add(snapshot);
                for (var k = -d; k <= d; k += 2)
                {
                    var down = k == -d || (k != d && v[offset + k - 1] < v[offset + k + 1]);
                    var x = down ? v[offset + k + 1] : v[offset + k - 1] + 1;
                    var y = x - k;
                    while (x < n && y < m && a[x] == b[y])
                    {
                        x++;
                        y++;
                    }
                    v[offset + k] = x;
                    if (x >= n && y >= m) return Backtrack(trace, a, b, d);
                }
            }
            return null;
        }

        private static List<Operation>? Backtrack(List<int[]> trace, string a, string b, int distance)
        {
            var reversed = new List<Operation>(a.Length + distance);
            var x = a.Length;
            var y = b.Length;
            for (var d = distance; d > 0; d--)
            {
                var before = trace[d];
                var k = x - y;
                if (k < -d || k > d) return null;
                var down = k == -d || (k != d && before[k - 1 + d + 1] < before[k + 1 + d + 1]);
                var previousK = down ? k + 1 : k - 1;
                var previousX = before[previousK + d + 1];
                var previousY = previousX - previousK;
                var stepX = down ? previousX : previousX + 1;
                var stepY = stepX - k;
                while (x > stepX && y > stepY)
                {
                    x--;
                    y--;
                    reversed.Add(new Operation(OperationKind.Equal, x));
                }
                if (down)
                {
                    if (previousY < 0 || previousY >= b.Length) return null;
                    reversed.Add(new Operation(OperationKind.Insert, previousY));
                }
                else
                {
                    if (previousX < 0 || previousX >= a.Length) return null;
                    reversed.Add(new Operation(OperationKind.Delete, previousX));
                }
                x = previousX;
                y = previousY;
            }
            while (x > 0 && y > 0)
            {
                x--;
                y--;
                reversed.Add(new Operation(OperationKind.Equal, x));
            }
            if (x != 0 || y != 0) return null;
            reversed.Reverse();
            return reversed;
        }

        private static List<TextSpanEdit> Merge(List<Operation> operations, string b,
            int prefix, string source)
        {
            var edits = new List<TextSpanEdit>();
            var position = 0;
            var start = -1;
            var deleted = 0;
            var inserted = new StringBuilder();
            foreach (var operation in operations)
            {
                switch (operation.Kind)
                {
                    case OperationKind.Equal:
                        Flush();
                        position++;
                        break;
                    case OperationKind.Delete:
                        if (start < 0) start = position;
                        deleted++;
                        position++;
                        break;
                    default:
                        if (start < 0) start = position;
                        inserted.Append(b[operation.Index]);
                        break;
                }
            }
            Flush();
            return edits;

            void Flush()
            {
                if (start < 0) return;
                var absolute = prefix + start;
                edits.Add(new TextSpanEdit(absolute, deleted, inserted.ToString(),
                    source.Substring(absolute, deleted)));
                start = -1;
                deleted = 0;
                inserted.Clear();
            }
        }
    }
}
