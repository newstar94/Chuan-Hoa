using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Scanning;

namespace ChuanHoa.DevelopmentAccessSmoke
{
    internal static class Program
    {
        private static int Main()
        {
            try
            {
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    access.WarmUp();
                    var refreshDeadline = Stopwatch.StartNew();
                    while (!access.HasCachedFeature(LocalAccessManager.FormatFeature) &&
                        refreshDeadline.Elapsed < TimeSpan.FromSeconds(10))
                    {
                        Thread.Sleep(25);
                    }
                    var format = access.GetRulePack(LocalAccessManager.FormatFeature);
                    var spelling = access.GetRulePack(LocalAccessManager.SpellingFeature);
                    var autoFix = access.GetRulePack(LocalAccessManager.AutoFixFeature);
                    if (!string.Equals(format.PackId, spelling.PackId, StringComparison.Ordinal))
                        throw new InvalidOperationException("Feature lanes resolved to different rule packs.");
                    if (!string.Equals(format.PackId, autoFix.PackId, StringComparison.Ordinal))
                        throw new InvalidOperationException("AUTOFIX resolved to a different rule pack.");
                    if (spelling.Lexicon.Count < 6000)
                        throw new InvalidOperationException("The signed Vietnamese lexicon is incomplete.");
                    var lexicon = new VietnameseLexiconSpellChecker(spelling.Lexicon);
                    if (!lexicon.IsKnown("dự") || lexicon.IsKnown("ự"))
                        throw new InvalidOperationException("Vietnamese lexicon lookup is incorrect.");
                    if (!spelling.Corrections.Any(item =>
                            string.Equals(item.Wrong, "quyết định xố", StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(item.Replacement, "quyết định số", StringComparison.OrdinalIgnoreCase)) ||
                        !spelling.Corrections.Any(item =>
                            string.Equals(item.Wrong, "ự án", StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(item.Replacement, "dự án", StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidOperationException("Contextual Vietnamese spelling corrections are missing.");
                    var spellingResult = new LocalDocumentScanner().ScanSpelling(
                        new LocalScanSnapshot("sha256:development-spelling", 1,
                            new[] { new LocalSectionSnapshot(1, 595, 842, 57, 57, 85, 43, false) },
                            new[] { new LocalParagraphSnapshot(1,
                                "Quyết định xố, ự án và nhậnn.", "wdMainTextStory", 1, 0,
                                "Times New Roman") },
                            Array.Empty<ChuanHoa.Client.Core.Annotations.AnnotationProtectedSpan>()),
                        spelling);
                    if (spellingResult.Findings.Count(item => item.RuleCode == "LOCAL-TYPO-DICT") != 2 ||
                        !spellingResult.Findings.Any(item => item.RuleCode == "LOCAL-TYPO-LEXICON" &&
                            string.Equals(item.Anchor.ExpectedText, "nhậnn", StringComparison.Ordinal)))
                        throw new InvalidOperationException("Vietnamese spelling scan did not detect the regression examples.");
                    var failClosedStopwatch = Stopwatch.StartNew();
                    try
                    {
                        access.GetRulePack("UNAVAILABLE_SMOKE_FEATURE");
                        throw new InvalidOperationException("An unavailable feature did not fail closed.");
                    }
                    catch (InvalidOperationException exception)
                    {
                        if (string.Equals(exception.Message,
                            "An unavailable feature did not fail closed.", StringComparison.Ordinal))
                            throw;
                    }
                    failClosedStopwatch.Stop();
                    if (failClosedStopwatch.ElapsedMilliseconds > 250)
                        throw new InvalidOperationException(
                            "A command waited for license network I/O: " +
                            failClosedStopwatch.ElapsedMilliseconds + " ms.");
                    var stopwatch = Stopwatch.StartNew();
                    for (var index = 0; index < 5000; index++)
                    {
                        if (!access.HasCachedFeature(LocalAccessManager.DocumentToolsFeature))
                            throw new InvalidOperationException("DOCUMENT_TOOLS disappeared from the validated cache.");
                    }
                    stopwatch.Stop();
                    if (stopwatch.ElapsedMilliseconds > 1000)
                        throw new InvalidOperationException(
                            "Ribbon access checks are too slow: " + stopwatch.ElapsedMilliseconds + " ms.");
                    Console.WriteLine("DEVELOPMENT_ACCESS_SMOKE_PASS " + format.PackId + " " + format.Version +
                        " LEXICON=" + spelling.Lexicon.Count +
                        " CACHE_5000_MS=" + stopwatch.ElapsedMilliseconds +
                        " FAIL_CLOSED_MS=" + failClosedStopwatch.ElapsedMilliseconds);
                }
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("DEVELOPMENT_ACCESS_SMOKE_FAIL " + exception);
                return 1;
            }
        }
    }
}
