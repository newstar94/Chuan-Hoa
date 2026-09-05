#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Annotations;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AnnotationSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main()
        {
            Word.Application application = null;
            Word.Document document = null;
            Word.Range userRange = null;
            Word.Range findingRange = null;
            Word.Comment userComment = null;
            try
            {
                application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                document = application.Documents.Add();
                document.Content.Text = "Người dùng sai định dạng\r";

                userRange = document.Range(0, 10);
                userRange.Font.Color = Word.WdColor.wdColorRed;
                userComment = document.Comments.Add(userRange, "Comment của người dùng");

                findingRange = document.Range(11, 14);
                var originalFindingColor = (int)findingRange.Font.Color;
                var adapter = new WordFindingAnnotationAdapter(application, document);
                var plan = CreatePlan();

                adapter.Apply(plan);
                Assert(document.Comments.Count == 2, "Expected one user comment and one add-in comment.");
                AssertAddInCommentPresentation(document.Comments[2], "[CHUẨN HÓA:FORMAT:finding-01]");
                Assert((int)findingRange.Font.Color == AnnotationOwnershipPolicy.WordRedColor,
                    "The exact finding span was not red.");
                Assert((int)userRange.Font.Color == AnnotationOwnershipPolicy.WordRedColor,
                    "The user's existing red text changed during apply.");
                findingRange.Select();
                string selectedLane;
                string selectedFindingId;
                Assert(adapter.TryGetSelectedFinding(out selectedLane, out selectedFindingId),
                    "The add-in could not resolve a selected owned finding.");
                Assert(selectedLane == "format" && selectedFindingId == "finding-01",
                    "The selected finding registry returned the wrong identity.");

                adapter.Apply(plan);
                Assert(document.Comments.Count == 2, "Rerun was not idempotent.");

                Word.Range changedByUser = null;
                try
                {
                    changedByUser = document.Range(11, 12);
                    changedByUser.Font.Color = Word.WdColor.wdColorGreen;
                }
                finally
                {
                    Release(changedByUser);
                }

                adapter.ClearLane("format");
                Assert(document.Comments.Count == 1, "Clear removed a user comment or retained an add-in comment.");
                Assert(document.Comments[1].Range.Text == "Comment của người dùng",
                    "The surviving user comment changed.");
                Assert((int)userRange.Font.Color == AnnotationOwnershipPolicy.WordRedColor,
                    "Clear removed the user's existing red text.");

                Word.Range preservedUserChange = null;
                Word.Range restoredByAddIn = null;
                try
                {
                    preservedUserChange = document.Range(11, 12);
                    restoredByAddIn = document.Range(12, 14);
                    Assert((int)preservedUserChange.Font.Color == (int)Word.WdColor.wdColorGreen,
                        "Clear overwrote a color changed by the user after annotation.");
                    Assert((int)restoredByAddIn.Font.Color == originalFindingColor,
                        "Clear did not restore the original finding color.");
                }
                finally
                {
                    Release(restoredByAddIn);
                    Release(preservedUserChange);
                }

                Assert(document.Variables.Count == 0, "Owned annotation variables were not cleaned.");

                RunInjectedFailureRollback(application);
                RunRerunStress(application);
                Console.WriteLine("ANNOTATION_WORD_SMOKE_PASS INJECTED_FAILURE_ROLLBACK");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("ANNOTATION_WORD_SMOKE_FAIL: " + exception);
                return 1;
            }
            finally
            {
                Release(userComment);
                Release(findingRange);
                Release(userRange);
                if (document != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    document.Close(ref save);
                }
                if (application != null)
                {
                    application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                Release(document);
                Release(application);
            }
        }

        private static AnnotationPlan CreatePlan()
        {
            var comments = new[]
            {
                new AnnotationCommentInstruction(
                    "finding-01",
                    "MainTextStory",
                    1,
                    11,
                    3,
                    "[CHUẨN HÓA:FORMAT:finding-01]",
                    "Hiện tại: Sai phông chữ.\nYêu cầu đúng: Dùng phông Times New Roman.")
            };
            var visualRanges = new[]
            {
                new AnnotationVisualInstruction("MainTextStory", 1, 11, 3)
            };
            return new AnnotationPlan(
                "format",
                "word-smoke-01",
                comments,
                visualRanges,
                new List<UnresolvedAnnotation>());
        }

        private static void AssertAddInCommentPresentation(Word.Comment comment, string marker)
        {
            Word.Range commentRange = null;
            Word.Range ownershipRange = null;
            try
            {
                commentRange = comment.Range;
                var visibleText = commentRange.Text ?? string.Empty;
                Assert(!visibleText.Contains(marker), "The ownership marker is visible to the user.");
                Assert(visibleText.StartsWith("Hiện tại:", StringComparison.Ordinal),
                    "The visible comment does not start with the current-state line.");

                ownershipRange = commentRange.Duplicate;
                ownershipRange.TextRetrievalMode.IncludeHiddenText = true;
                Assert(!(ownershipRange.Text ?? string.Empty).Contains(marker),
                    "The comment still contains the legacy hidden ownership marker.");
                Assert(visibleText.Contains("Hiện tại: Sai phông chữ."), "The current-state line is missing.");
                Assert(visibleText.Contains("Yêu cầu đúng: Dùng phông Times New Roman."),
                    "The correction line is missing.");
                Assert(!visibleText.Contains("Mã quy tắc:") && !visibleText.Contains("Mức độ:") && !visibleText.Contains("Căn cứ:"),
                    "A removed user-facing field is still present.");
            }
            finally
            {
                Release(ownershipRange);
                Release(commentRange);
            }
        }

        private static void RunRerunStress(Word.Application application)
        {
            Word.Document stressDocument = null;
            try
            {
                const int textLength = 12000;
                stressDocument = application.Documents.Add();
                stressDocument.Content.Text = new string('A', textLength) + "\r";
                var adapter = new WordFindingAnnotationAdapter(application, stressDocument);
                var comments = new[]
                {
                    new AnnotationCommentInstruction(
                        "stress-01",
                        "MainTextStory",
                        1,
                        0,
                        textLength,
                        "[CHUẨN HÓA:FORMAT:stress-01]",
                        "[CHUẨN HÓA:FORMAT:stress-01]\nMã quy tắc: STRESS")
                };
                var visuals = new[]
                {
                    new AnnotationVisualInstruction("MainTextStory", 1, 0, textLength)
                };
                var plan = new AnnotationPlan(
                    "format",
                    "word-stress-01",
                    comments,
                    visuals,
                    new List<UnresolvedAnnotation>());

                var first = Stopwatch.StartNew();
                adapter.Apply(plan);
                first.Stop();

                var second = Stopwatch.StartNew();
                adapter.Apply(plan);
                second.Stop();

                Console.WriteLine(
                    "ANNOTATION_RERUN_TIMING length=" + textLength +
                    " first_ms=" + first.ElapsedMilliseconds +
                    " second_ms=" + second.ElapsedMilliseconds);
                Assert(stressDocument.Comments.Count == 1, "Stress rerun duplicated add-in comments.");
            }
            finally
            {
                if (stressDocument != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    stressDocument.Close(ref save);
                }
                Release(stressDocument);
            }
        }

        private static void RunInjectedFailureRollback(Word.Application application)
        {
            Word.Document rollbackDocument = null;
            Word.Range userRange = null;
            Word.Range findingRange = null;
            Word.Comment userComment = null;
            try
            {
                rollbackDocument = application.Documents.Add();
                rollbackDocument.Content.Text = "Người dùng sai định dạng\r";
                userRange = rollbackDocument.Range(0, 10);
                userRange.Font.Color = Word.WdColor.wdColorGreen;
                userComment = rollbackDocument.Comments.Add(userRange, "Comment cần được bảo toàn");
                findingRange = rollbackDocument.Range(11, 14);

                var plan = CreatePlan();
                var stableAdapter = new WordFindingAnnotationAdapter(application, rollbackDocument);
                stableAdapter.Apply(plan);
                Assert(rollbackDocument.Comments.Count == 2,
                    "Injected rollback setup did not create the stable annotation state.");
                Assert(rollbackDocument.Variables.Count == 2,
                    "Injected rollback setup did not register both owned annotation records.");

                var injected = false;
                var failingAdapter = new WordFindingAnnotationAdapter(
                    application,
                    rollbackDocument,
                    (stage, index) =>
                    {
                        if (stage == "after-visual" && index == 0)
                        {
                            injected = true;
                            throw new InvalidOperationException("ANNOTATION_FAULT_INJECTION");
                        }
                    });
                try
                {
                    failingAdapter.Apply(plan);
                    throw new InvalidOperationException("The injected annotation failure did not fire.");
                }
                catch (InvalidOperationException exception)
                {
                    Assert(injected && exception.Message == "ANNOTATION_FAULT_INJECTION",
                        "The annotation failure was not the expected injected fault.");
                }

                Assert(rollbackDocument.Comments.Count == 2,
                    "Rollback did not restore the previous add-in comment and user comment.");
                Assert(rollbackDocument.Comments[1].Range.Text == "Comment cần được bảo toàn",
                    "Rollback changed the user's comment.");
                Assert((int)userRange.Font.Color == (int)Word.WdColor.wdColorGreen,
                    "Rollback changed the user's font colour.");
                Assert((int)findingRange.Font.Color == AnnotationOwnershipPolicy.WordRedColor,
                    "Rollback did not restore the previous owned red marker.");
                Assert(rollbackDocument.Variables.Count == 2,
                    "Rollback left a partial or missing annotation registry.");
                findingRange.Select();
                string selectedLane;
                string selectedFindingId;
                Assert(stableAdapter.TryGetSelectedFinding(out selectedLane, out selectedFindingId),
                    "Rollback did not restore the selected-finding identity.");
                Assert(selectedLane == "format" && selectedFindingId == "finding-01",
                    "Rollback restored the wrong finding identity.");
                Console.WriteLine("ANNOTATION_INJECTED_FAILURE_ROLLBACK_PASS");
            }
            finally
            {
                Release(userComment);
                Release(findingRange);
                Release(userRange);
                if (rollbackDocument != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    rollbackDocument.Close(ref save);
                }
                Release(rollbackDocument);
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }
    }
}
