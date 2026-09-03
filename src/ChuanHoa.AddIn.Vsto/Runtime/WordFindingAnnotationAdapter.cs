using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using ChuanHoa.Client.Core.Annotations;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// Applies a pre-validated annotation plan. It only removes comments, bookmarks and
    /// document variables that carry the add-in's exact lane marker.
    /// </summary>
    public sealed class WordFindingAnnotationAdapter
    {
        private const string VariablePrefix = "CHANN_";
        private const string BookmarkPrefix = "CHAF_";
        private const string CommentVariablePrefix = "CHCOM_";
        private const string CommentBookmarkPrefix = "CHAC_";
        private const int AddInRed = (int)Word.WdColor.wdColorRed;
        private const int MixedFormatting = (int)Word.WdConstants.wdUndefined;
        private readonly Word.Application _application;
        private readonly Word.Document _document;
        private readonly bool _supportsCustomUndoRecord;

        public WordFindingAnnotationAdapter(Word.Application application, Word.Document document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _supportsCustomUndoRecord = ReadWordMajorVersion(application) >= 15;
        }

        public void Apply(AnnotationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            if (plan.Unresolved.Count > 0)
            {
                throw new InvalidOperationException(
                    "The annotation plan contains unresolved anchors and must not be applied.");
            }
            if (_document.ReadOnly ||
                _document.ProtectionType != Word.WdProtectionType.wdNoProtection ||
                _document.TrackRevisions)
            {
                throw new InvalidOperationException(
                    "Annotations are disabled for read-only, protected or Track Changes documents.");
            }

            ValidateRanges(plan);
            var undoStarted = false;
            try
            {
                if (_supportsCustomUndoRecord)
                {
                    _application.UndoRecord.StartCustomRecord("Chuẩn hóa: đánh dấu lỗi");
                    undoStarted = true;
                }

                ClearLane(plan.Lane);
                for (var index = 0; index < plan.VisualRanges.Count; index++)
                {
                    ApplyRedMarker(plan, plan.VisualRanges[index], index);
                }
                for (var index = 0; index < plan.Comments.Count; index++)
                {
                    AddComment(plan, plan.Comments[index], index);
                }
            }
            catch
            {
                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                    undoStarted = false;
                    object count = 1;
                    _document.Undo(ref count);
                }
                else
                {
                    ClearLane(plan.Lane);
                }
                throw;
            }
            finally
            {
                if (undoStarted)
                {
                    _application.UndoRecord.EndCustomRecord();
                }
            }
        }

        public void ClearLane(string lane)
        {
            var normalizedLane = AnnotationOwnershipPolicy.NormalizeLane(lane);
            var registeredCommentSignatures = ReadRegisteredCommentSignaturesExcluding(normalizedLane);
            ClearRegisteredComments(normalizedLane);

            // Remove comments produced by versions that persisted a hidden marker in
            // the comment body. New comments use an external registry so Word's Modern
            // Comments pane no longer displays the misleading "Missing content" card.
            for (var index = _document.Comments.Count; index >= 1; index--)
            {
                Word.Comment? comment = null;
                Word.Range? commentRange = null;
                try
                {
                    comment = _document.Comments[index];
                    commentRange = comment.Range;
                    commentRange.TextRetrievalMode.IncludeHiddenText = true;
                    var text = commentRange.Text ?? string.Empty;
                    if (AnnotationOwnershipPolicy.IsOwnedComment(text, lane))
                    {
                        comment.Delete();
                    }
                }
                finally
                {
                    Release(commentRange);
                    Release(comment);
                }
            }

            // A previous text mutation or a Word Modern Comments reconciliation can
            // remove/move the external bookmark before the registered comment is
            // cleared. Such an orphan has no usable finding identity and makes
            // "Sửa lỗi đang chọn" fail forever. Remove only the add-in's exact author,
            // initials and two-line presentation; user comments are left untouched.
            for (var index = _document.Comments.Count; index >= 1; index--)
            {
                Word.Comment? comment = null;
                Word.Range? commentRange = null;
                try
                {
                    comment = _document.Comments[index];
                    if (!string.Equals(comment.Author, "Chuẩn hóa", StringComparison.Ordinal) ||
                        !string.Equals(comment.Initial, "CH", StringComparison.Ordinal))
                        continue;
                    commentRange = comment.Range;
                    var text = commentRange.Text ?? string.Empty;
                    if (text.StartsWith("Hiện tại:", StringComparison.Ordinal) &&
                        text.IndexOf("Yêu cầu đúng:", StringComparison.Ordinal) >= 0 &&
                        !registeredCommentSignatures.Contains(CommentSignature(text)))
                        comment.Delete();
                }
                finally
                {
                    Release(commentRange);
                    Release(comment);
                }
            }

            var variableNamePrefix = VariablePrefix + normalizedLane + "_";
            for (var index = _document.Variables.Count; index >= 1; index--)
            {
                Word.Variable? variable = null;
                try
                {
                    variable = _document.Variables[index];
                    if (!variable.Name.StartsWith(variableNamePrefix, StringComparison.Ordinal))
                    {
                        continue;
                    }
                    RestoreVisualMarker(variable.Value ?? string.Empty);
                    variable.Delete();
                }
                finally
                {
                    Release(variable);
                }
            }
        }

        private HashSet<string> ReadRegisteredCommentSignaturesExcluding(string normalizedLane)
        {
            var signatures = new HashSet<string>(StringComparer.Ordinal);
            var excludedPrefix = CommentVariablePrefix + normalizedLane + "_";
            for (var index = 1; index <= _document.Variables.Count; index++)
            {
                Word.Variable? variable = null;
                try
                {
                    variable = _document.Variables[index];
                    if (!variable.Name.StartsWith(CommentVariablePrefix, StringComparison.Ordinal)) continue;
                    if (variable.Name.StartsWith(excludedPrefix, StringComparison.Ordinal)) continue;
                    var values = (variable.Value ?? string.Empty).Split(';');
                    if (values.Length > 1 && !string.IsNullOrWhiteSpace(values[1])) signatures.Add(values[1]);
                }
                finally { Release(variable); }
            }
            return signatures;
        }

        public bool TryGetSelectedFinding(out string lane, out string findingId)
        {
            lane = string.Empty;
            findingId = string.Empty;
            Word.Range? selectedRange = null;
            try
            {
                selectedRange = ResolveSelectedCommentScopeOrDocumentRange();
                var bestLength = int.MaxValue;
                for (var variableIndex = 1; variableIndex <= _document.Variables.Count; variableIndex++)
                {
                    Word.Variable? variable = null;
                    Word.Bookmark? bookmark = null;
                    Word.Range? scope = null;
                    try
                    {
                        variable = _document.Variables[variableIndex];
                        if (!variable.Name.StartsWith(CommentVariablePrefix, StringComparison.Ordinal)) continue;
                        var separator = variable.Name.IndexOf('_', CommentVariablePrefix.Length);
                        if (separator <= CommentVariablePrefix.Length) continue;
                        var values = (variable.Value ?? string.Empty).Split(';');
                        if (values.Length < 3) continue;
                        var matchesSelection = false;
                        var candidateLength = int.MaxValue;
                        int storedStory;
                        int storedStart;
                        int storedEnd;
                        if (values.Length >= 6 &&
                            int.TryParse(values[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out storedStory) &&
                            int.TryParse(values[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out storedStart) &&
                            int.TryParse(values[5], NumberStyles.Integer, CultureInfo.InvariantCulture, out storedEnd))
                        {
                            matchesSelection = storedStory == (int)selectedRange.StoryType &&
                                NumericRangesTouch(storedStart, storedEnd, selectedRange.Start, selectedRange.End);
                            candidateLength = Math.Max(0, storedEnd - storedStart);
                        }
                        else if (_document.Bookmarks.Exists(values[0]))
                        {
                            bookmark = _document.Bookmarks[values[0]];
                            scope = bookmark.Range.Duplicate;
                            matchesSelection = scope.StoryType == selectedRange.StoryType &&
                                RangesTouch(scope, selectedRange);
                            candidateLength = Math.Max(0, scope.End - scope.Start);
                        }
                        if (!matchesSelection) continue;
                        var candidateId = DecodeFindingId(values[2]);
                        if (string.IsNullOrWhiteSpace(candidateId)) continue;
                        if (candidateLength >= bestLength) continue;
                        bestLength = candidateLength;
                        lane = variable.Name.Substring(CommentVariablePrefix.Length,
                            separator - CommentVariablePrefix.Length).ToLowerInvariant();
                        findingId = candidateId;
                    }
                    finally
                    {
                        Release(scope);
                        Release(bookmark);
                        Release(variable);
                    }
                }
                return !string.IsNullOrWhiteSpace(lane) && !string.IsNullOrWhiteSpace(findingId);
            }
            finally { Release(selectedRange); }
        }

        public bool TryGetSelectedDocumentRange(out string storyType, out int start, out int end)
        {
            storyType = string.Empty;
            start = 0;
            end = 0;
            Word.Range? selectedRange = null;
            try
            {
                selectedRange = ResolveSelectedCommentScopeOrDocumentRange();
                storyType = selectedRange.StoryType.ToString();
                start = selectedRange.Start;
                end = selectedRange.End;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            finally { Release(selectedRange); }
        }

        public bool TryFocusDocumentSelection()
        {
            Word.Range? documentRange = null;
            try
            {
                documentRange = ResolveSelectedCommentScopeOrDocumentRange();
                if (documentRange.StoryType == Word.WdStoryType.wdCommentsStory) return false;
                // Commands only need to return keyboard focus to the document. Keeping
                // the whole comment scope selected is dangerous: after Modern Comments
                // reconciles a deleted thread, Word can treat the stale selection as a
                // pending replacement range and remove its text during a later Ribbon
                // command. Collapse to a caret before selecting so focus changes can
                // never overwrite the finding's document content.
                documentRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                documentRange.Select();
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            finally { Release(documentRange); }
        }

        public void ClearOwnedAnnotationsAt(string lane, string storyType, int start, int end)
        {
            var normalizedLane = AnnotationOwnershipPolicy.NormalizeLane(lane);
            var expectedStory = ParseStoryType(storyType);
            var commentPrefix = CommentVariablePrefix + normalizedLane + "_";
            for (var variableIndex = _document.Variables.Count; variableIndex >= 1; variableIndex--)
            {
                Word.Variable? variable = null;
                Word.Bookmark? bookmark = null;
                Word.Range? scope = null;
                try
                {
                    variable = _document.Variables[variableIndex];
                    if (!variable.Name.StartsWith(commentPrefix, StringComparison.Ordinal)) continue;
                    var values = (variable.Value ?? string.Empty).Split(';');
                    if (values.Length < 2) continue;
                    int storedStory;
                    int storedStart;
                    int storedEnd;
                    var matches = values.Length >= 6 &&
                        int.TryParse(values[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out storedStory) &&
                        int.TryParse(values[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out storedStart) &&
                        int.TryParse(values[5], NumberStyles.Integer, CultureInfo.InvariantCulture, out storedEnd) &&
                        storedStory == (int)expectedStory && NumericRangesTouch(storedStart, storedEnd, start, end);
                    if (!matches) continue;
                    if (_document.Bookmarks.Exists(values[0]))
                    {
                        bookmark = _document.Bookmarks[values[0]];
                        scope = bookmark.Range.Duplicate;
                        DeleteRegisteredComment(scope, values[1]);
                        bookmark.Delete();
                    }
                    variable.Delete();
                }
                finally
                {
                    Release(scope);
                    Release(bookmark);
                    Release(variable);
                }
            }

            var visualPrefix = VariablePrefix + normalizedLane + "_";
            for (var variableIndex = _document.Variables.Count; variableIndex >= 1; variableIndex--)
            {
                Word.Variable? variable = null;
                Word.Bookmark? bookmark = null;
                Word.Range? scope = null;
                try
                {
                    variable = _document.Variables[variableIndex];
                    if (!variable.Name.StartsWith(visualPrefix, StringComparison.Ordinal)) continue;
                    var values = (variable.Value ?? string.Empty).Split(';');
                    if (values.Length == 0 || !_document.Bookmarks.Exists(values[0])) continue;
                    bookmark = _document.Bookmarks[values[0]];
                    scope = bookmark.Range.Duplicate;
                    if (scope.StoryType != expectedStory ||
                        !NumericRangesTouch(scope.Start, scope.End, start, end)) continue;
                    RestoreVisualMarker(variable.Value ?? string.Empty);
                    variable.Delete();
                }
                finally
                {
                    Release(scope);
                    Release(bookmark);
                    Release(variable);
                }
            }
        }

        private Word.Range ResolveSelectedCommentScopeOrDocumentRange()
        {
            var selected = _application.Selection.Range.Duplicate;
            if (selected.StoryType != Word.WdStoryType.wdCommentsStory) return selected;
            for (var commentIndex = 1; commentIndex <= _document.Comments.Count; commentIndex++)
            {
                Word.Comment? comment = null;
                Word.Range? body = null;
                try
                {
                    comment = _document.Comments[commentIndex];
                    body = comment.Range;
                    if (body.StoryType != selected.StoryType || !RangesTouch(body, selected)) continue;
                    var scope = comment.Scope.Duplicate;
                    Release(selected);
                    return scope;
                }
                finally
                {
                    Release(body);
                    Release(comment);
                }
            }
            return selected;
        }

        private static bool RangesTouch(Word.Range left, Word.Range right)
        {
            if (left.Start == left.End) return right.Start <= left.Start && left.Start <= right.End;
            if (right.Start == right.End) return left.Start <= right.Start && right.Start <= left.End;
            return left.Start < right.End && right.Start < left.End;
        }

        private static bool NumericRangesTouch(int leftStart, int leftEnd, int rightStart, int rightEnd)
        {
            if (leftStart == leftEnd) return rightStart <= leftStart && leftStart <= rightEnd;
            if (rightStart == rightEnd) return leftStart <= rightStart && rightStart <= leftEnd;
            return leftStart < rightEnd && rightStart < leftEnd;
        }

        private void ClearRegisteredComments(string normalizedLane)
        {
            var prefix = CommentVariablePrefix + normalizedLane + "_";
            for (var variableIndex = _document.Variables.Count; variableIndex >= 1; variableIndex--)
            {
                Word.Variable? variable = null;
                Word.Bookmark? bookmark = null;
                Word.Range? bookmarkRange = null;
                try
                {
                    variable = _document.Variables[variableIndex];
                    if (!variable.Name.StartsWith(prefix, StringComparison.Ordinal)) continue;

                    var values = (variable.Value ?? string.Empty).Split(';');
                    var bookmarkName = values.Length > 0 ? values[0] : string.Empty;
                    var signature = values.Length > 1 ? values[1] : string.Empty;
                    if (!string.IsNullOrWhiteSpace(bookmarkName) && _document.Bookmarks.Exists(bookmarkName))
                    {
                        bookmark = _document.Bookmarks[bookmarkName];
                        bookmarkRange = bookmark.Range.Duplicate;
                        DeleteRegisteredComment(bookmarkRange, signature);
                        bookmark.Delete();
                    }
                    variable.Delete();
                }
                finally
                {
                    Release(bookmarkRange);
                    Release(bookmark);
                    Release(variable);
                }
            }
        }

        private void DeleteRegisteredComment(Word.Range bookmarkRange, string signature)
        {
            for (var commentIndex = _document.Comments.Count; commentIndex >= 1; commentIndex--)
            {
                Word.Comment? comment = null;
                Word.Range? scope = null;
                Word.Range? body = null;
                try
                {
                    comment = _document.Comments[commentIndex];
                    scope = comment.Scope;
                    if (scope.StoryType != bookmarkRange.StoryType ||
                        scope.Start != bookmarkRange.Start || scope.End != bookmarkRange.End ||
                        !string.Equals(comment.Author, "Chuẩn hóa", StringComparison.Ordinal) ||
                        !string.Equals(comment.Initial, "CH", StringComparison.Ordinal))
                        continue;

                    body = comment.Range;
                    if (!string.Equals(CommentSignature(body.Text ?? string.Empty), signature,
                            StringComparison.Ordinal))
                        continue;
                    comment.Delete();
                    return;
                }
                finally
                {
                    Release(body);
                    Release(scope);
                    Release(comment);
                }
            }
        }

        private void ValidateRanges(AnnotationPlan plan)
        {
            foreach (var comment in plan.Comments)
            {
                try
                {
                    using (var range = new ComRange(ResolveRange(
                        comment.StoryType,
                        comment.SectionIndex,
                        comment.Start,
                        comment.Length)))
                    {
                    }
                }
                catch (COMException exception)
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture,
                        "Invalid Word comment range for finding {0}: story={1}, section={2}, start={3}, length={4}.",
                        comment.FindingId, comment.StoryType, comment.SectionIndex, comment.Start, comment.Length), exception);
                }
            }
            foreach (var visual in plan.VisualRanges)
            {
                try
                {
                    using (var range = new ComRange(ResolveRange(
                        visual.StoryType,
                        visual.SectionIndex,
                        visual.Start,
                        visual.Length)))
                    {
                        if (range.Value.Editors.Count > 0)
                        {
                            throw new InvalidOperationException("A red marker intersects an editor-protected range.");
                        }
                        if (range.Value.Fields.Count > 0 ||
                            range.Value.ContentControls.Count > 0 ||
                            range.Value.InlineShapes.Count > 0)
                        {
                            throw new InvalidOperationException(
                                "A red marker intersects a field, content control or inline object.");
                        }
                    }
                }
                catch (COMException exception)
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture,
                        "Invalid Word visual range: story={0}, section={1}, start={2}, length={3}.",
                        visual.StoryType, visual.SectionIndex, visual.Start, visual.Length), exception);
                }
            }
        }

        private void ApplyRedMarker(
            AnnotationPlan plan,
            AnnotationVisualInstruction instruction,
            int index)
        {
            using (var range = new ComRange(ResolveRange(
                instruction.StoryType,
                instruction.SectionIndex,
                instruction.Start,
                instruction.Length)))
            {
                var bookmarkName = CreateBookmarkName(plan, index);
                var state = CaptureColorState(range.Value);
                if (state.Length > 30000)
                {
                    throw new InvalidOperationException("The original color state is too large to persist safely.");
                }

                _document.Bookmarks.Add(bookmarkName, range.Value);
                var variableName = VariablePrefix + AnnotationOwnershipPolicy.NormalizeLane(plan.Lane) + "_" + index.ToString("D4", CultureInfo.InvariantCulture);
                _document.Variables.Add(variableName, bookmarkName + ";" + state);
                WriteFontColor(range.Value, AddInRed);
            }
        }

        private void AddComment(AnnotationPlan plan, AnnotationCommentInstruction instruction, int index)
        {
            using (var range = new ComRange(ResolveRange(
                instruction.StoryType,
                instruction.SectionIndex,
                instruction.Start,
                instruction.Length)))
            {
                Word.Comment? comment = null;
                Word.Variable? variable = null;
                try
                {
                    var bookmarkName = CreateCommentBookmarkName(plan, index);
                    _document.Bookmarks.Add(bookmarkName, range.Value);
                    var variableName = CommentVariablePrefix + AnnotationOwnershipPolicy.NormalizeLane(plan.Lane) + "_" +
                        index.ToString("D4", CultureInfo.InvariantCulture);
                    variable = _document.Variables.Add(variableName,
                        bookmarkName + ";" + CommentSignature(instruction.CommentText) + ";" +
                        Convert.ToBase64String(Encoding.UTF8.GetBytes(instruction.FindingId)) + ";" +
                        ((int)range.Value.StoryType).ToString(CultureInfo.InvariantCulture) + ";" +
                        range.Value.Start.ToString(CultureInfo.InvariantCulture) + ";" +
                        range.Value.End.ToString(CultureInfo.InvariantCulture));
                    comment = _document.Comments.Add(range.Value, instruction.CommentText);
                    comment.Author = "Chuẩn hóa";
                    comment.Initial = "CH";
                }
                finally
                {
                    Release(variable);
                    Release(comment);
                }
            }
        }

        private string CaptureColorState(Word.Range range)
        {
            var rangeLength = range.End - range.Start;
            if (rangeLength <= 0)
            {
                return string.Empty;
            }

            // Word returns wdUndefined when a range contains mixed font colours. Most
            // findings are uniformly coloured, so persist that range in one COM call
            // instead of crossing the Word boundary once for every character.
            var uniformColor = ReadFontColor(range);
            if (uniformColor != MixedFormatting)
            {
                return EncodeSegment(0, rangeLength, uniformColor);
            }

            var segments = new List<ColorRun>();
            CaptureColorRuns(range, 0, rangeLength, segments);
            return string.Join(",", segments.Select(item =>
                EncodeSegment(item.Start, item.Length, item.Color)));
        }

        private void RestoreVisualMarker(string serialized)
        {
            var separator = serialized.IndexOf(';');
            if (separator <= 0)
            {
                return;
            }
            var bookmarkName = serialized.Substring(0, separator);
            if (!_document.Bookmarks.Exists(bookmarkName))
            {
                return;
            }

            Word.Bookmark? bookmark = null;
            Word.Range? bookmarkRange = null;
            try
            {
                bookmark = _document.Bookmarks[bookmarkName];
                bookmarkRange = bookmark.Range.Duplicate;
                var state = serialized.Substring(separator + 1);
                foreach (var encodedSegment in state.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    RestoreColorSegment(bookmarkRange, encodedSegment);
                }
                bookmark.Delete();
            }
            finally
            {
                Release(bookmarkRange);
                Release(bookmark);
            }
        }

        private static void RestoreColorSegment(Word.Range bookmarkRange, string encoded)
        {
            var values = encoded.Split(':');
            int offset;
            int length;
            int originalColor;
            if (values.Length != 3 ||
                !int.TryParse(values[0], NumberStyles.None, CultureInfo.InvariantCulture, out offset) ||
                !int.TryParse(values[1], NumberStyles.None, CultureInfo.InvariantCulture, out length) ||
                !int.TryParse(values[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out originalColor) ||
                offset < 0 || length <= 0 || offset + length > bookmarkRange.End - bookmarkRange.Start)
            {
                return;
            }

            Word.Range? segment = null;
            try
            {
                segment = bookmarkRange.Duplicate;
                segment.SetRange(bookmarkRange.Start + offset, bookmarkRange.Start + offset + length);

                // Preserve a user's later colour change. A uniformly red segment is
                // still owned by the add-in and can be restored with one Word call; a
                // uniformly non-red segment belongs to the user and must be left alone.
                var segmentColor = ReadFontColor(segment);
                if (segmentColor == AddInRed)
                {
                    WriteFontColor(segment, originalColor);
                    return;
                }
                if (segmentColor != MixedFormatting)
                {
                    return;
                }

                RestoreOwnedRedRuns(segment, originalColor);
            }
            finally
            {
                Release(segment);
            }
        }

        private static void CaptureColorRuns(
            Word.Range source,
            int offset,
            int length,
            IList<ColorRun> result)
        {
            Word.Range? probe = null;
            try
            {
                probe = source.Duplicate;
                probe.SetRange(source.Start + offset, source.Start + offset + length);
                var color = ReadFontColor(probe);
                if (color != MixedFormatting || length == 1)
                {
                    AppendColorRun(result, offset, length, color);
                    return;
                }
            }
            finally
            {
                Release(probe);
            }

            // Binary subdivision turns a normal mixed range with a few colour runs
            // into O(runs * log(length)) Word calls instead of O(length) calls.
            var leftLength = length / 2;
            CaptureColorRuns(source, offset, leftLength, result);
            CaptureColorRuns(source, offset + leftLength, length - leftLength, result);
        }

        private static void AppendColorRun(IList<ColorRun> result, int start, int length, int color)
        {
            if (result.Count > 0)
            {
                var previous = result[result.Count - 1];
                if (previous.Color == color && previous.Start + previous.Length == start)
                {
                    previous.Length += length;
                    return;
                }
            }
            result.Add(new ColorRun(start, length, color));
        }

        private static void RestoreOwnedRedRuns(Word.Range source, int originalColor)
        {
            var length = source.End - source.Start;
            if (length <= 0)
            {
                return;
            }

            var color = ReadFontColor(source);
            if (color == AddInRed)
            {
                WriteFontColor(source, originalColor);
                return;
            }
            if (color != MixedFormatting || length == 1)
            {
                return;
            }

            var midpoint = source.Start + length / 2;
            Word.Range? left = null;
            Word.Range? right = null;
            try
            {
                left = source.Duplicate;
                left.SetRange(source.Start, midpoint);
                right = source.Duplicate;
                right.SetRange(midpoint, source.End);
                RestoreOwnedRedRuns(left, originalColor);
                RestoreOwnedRedRuns(right, originalColor);
            }
            finally
            {
                Release(right);
                Release(left);
            }
        }

        private static int ReadFontColor(Word.Range range)
        {
            Word.Font? font = null;
            try
            {
                font = range.Font;
                return (int)font.Color;
            }
            finally
            {
                Release(font);
            }
        }

        private static void WriteFontColor(Word.Range range, int color)
        {
            Word.Font? font = null;
            try
            {
                font = range.Font;
                font.Color = (Word.WdColor)color;
            }
            finally
            {
                Release(font);
            }
        }

        private Word.Range ResolveRange(string storyType, int sectionIndex, int start, int length)
        {
            Word.Range? storyRange = null;
            Word.Section? section = null;
            Word.HeaderFooter? headerFooter = null;
            try
            {
                var story = ParseStoryType(storyType);
                if (story == Word.WdStoryType.wdMainTextStory)
                {
                    return _document.Range(start, checked(start + length));
                }

                section = _document.Sections[sectionIndex];
                if (IsHeaderStory(story))
                {
                    headerFooter = section.Headers[HeaderFooterIndex(story)];
                }
                else if (IsFooterStory(story))
                {
                    headerFooter = section.Footers[HeaderFooterIndex(story)];
                }
                else
                {
                    throw new NotSupportedException("Unsupported Word story: " + storyType + ".");
                }
                if (!headerFooter.Exists)
                {
                    throw new InvalidOperationException("The requested header or footer does not exist.");
                }
                storyRange = headerFooter.Range.Duplicate;
                if (start < storyRange.Start || start + length > storyRange.End)
                {
                    throw new InvalidOperationException("The annotation range is outside its Word story.");
                }
                storyRange.SetRange(start, checked(start + length));
                var result = storyRange;
                storyRange = null;
                return result;
            }
            finally
            {
                Release(storyRange);
                Release(headerFooter);
                Release(section);
            }
        }

        private static Word.WdStoryType ParseStoryType(string storyType)
        {
            Word.WdStoryType result;
            return Enum.TryParse(storyType.StartsWith("wd", StringComparison.Ordinal) ? storyType : "wd" + storyType, out result)
                ? result
                : throw new NotSupportedException("Unknown Word story: " + storyType + ".");
        }

        private static bool IsHeaderStory(Word.WdStoryType story)
        {
            return story == Word.WdStoryType.wdPrimaryHeaderStory ||
                story == Word.WdStoryType.wdFirstPageHeaderStory ||
                story == Word.WdStoryType.wdEvenPagesHeaderStory;
        }

        private static bool IsFooterStory(Word.WdStoryType story)
        {
            return story == Word.WdStoryType.wdPrimaryFooterStory ||
                story == Word.WdStoryType.wdFirstPageFooterStory ||
                story == Word.WdStoryType.wdEvenPagesFooterStory;
        }

        private static Word.WdHeaderFooterIndex HeaderFooterIndex(Word.WdStoryType story)
        {
            if (story == Word.WdStoryType.wdFirstPageHeaderStory || story == Word.WdStoryType.wdFirstPageFooterStory)
            {
                return Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage;
            }
            if (story == Word.WdStoryType.wdEvenPagesHeaderStory || story == Word.WdStoryType.wdEvenPagesFooterStory)
            {
                return Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages;
            }
            return Word.WdHeaderFooterIndex.wdHeaderFooterPrimary;
        }

        private static string EncodeSegment(int start, int length, int color)
        {
            return start.ToString(CultureInfo.InvariantCulture) + ":" +
                length.ToString(CultureInfo.InvariantCulture) + ":" +
                color.ToString(CultureInfo.InvariantCulture);
        }

        private static string CreateBookmarkName(AnnotationPlan plan, int index)
        {
            using (var sha1 = SHA1.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(plan.Lane + "\u001f" + plan.ScanId + "\u001f" + index);
                var hash = sha1.ComputeHash(bytes);
                var token = string.Concat(hash.Take(8).Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
                return BookmarkPrefix + token;
            }
        }

        private static string CreateCommentBookmarkName(AnnotationPlan plan, int index)
        {
            using (var sha1 = SHA1.Create())
            {
                var bytes = Encoding.UTF8.GetBytes("comment\u001f" + plan.Lane + "\u001f" + plan.ScanId + "\u001f" + index);
                var hash = sha1.ComputeHash(bytes);
                var token = string.Concat(hash.Take(8).Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
                return CommentBookmarkPrefix + token;
            }
        }

        private static string CommentSignature(string value)
        {
            var normalized = (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            using (var sha256 = SHA256.Create())
            {
                var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(normalized));
                return string.Concat(hash.Take(12).Select(item => item.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        private static string DecodeFindingId(string value)
        {
            try { return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty)); }
            catch (FormatException) { return string.Empty; }
        }

        private static int ReadWordMajorVersion(Word.Application application)
        {
            var value = application.Version ?? string.Empty;
            var separator = value.IndexOf('.');
            var major = separator < 0 ? value : value.Substring(0, separator);
            int result;
            return int.TryParse(major, NumberStyles.None, CultureInfo.InvariantCulture, out result) ? result : 0;
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        private sealed class ComRange : IDisposable
        {
            public ComRange(Word.Range value)
            {
                Value = value;
            }

            public Word.Range Value { get; }

            public void Dispose()
            {
                Release(Value);
            }
        }

        private sealed class ColorRun
        {
            public ColorRun(int start, int length, int color)
            {
                Start = start;
                Length = length;
                Color = color;
            }

            public int Start { get; }
            public int Length { get; set; }
            public int Color { get; }
        }
    }
}
