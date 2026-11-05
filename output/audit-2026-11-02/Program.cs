using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ChuanHoa.AddIn.Vsto.Runtime;
using ChuanHoa.Client.Core.Annotations;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Scanning;
using ChuanHoa.Client.Core.Text;
using Word = Microsoft.Office.Interop.Word;

internal static class Program
{
    const string Title = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
    const string Motto = "Độc lập - Tự do - Hạnh phúc";
    const string Rule = "ND30-PL1-M2-K1-TN-LINE";
    static readonly BindingFlags PrivateStatic = BindingFlags.NonPublic | BindingFlags.Static;

    [STAThread]
    static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        if (args.Length > 1) Console.SetOut(new System.IO.StreamWriter(args[1], false, Encoding.UTF8) { AutoFlush = true });
        Console.WriteLine("PROBE_PROCESS x64=" + Environment.Is64BitProcess);
        if (args.Length == 0 || args[0] == "core") { Core(); return 0; }
        if (args[0] == "profile") { Profile(); return 0; }
        if (args[0] == "edge") { EdgeCases(); return 0; }
        Word.Application? app = null;
        try
        {
            app = new Word.Application();
            app.Visible = false; app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            Console.WriteLine("WORD_INSTANCE version=" + app.Version);
            if (args[0] == "word")
            {
                LateBorder(app);
                OldLine(app);
                SafetyGuard(app);
            }
            Console.WriteLine("WORD_PROBE_COMPLETE");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
        finally
        {
            if (app != null)
            {
                try { app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch (Exception ex) { Console.WriteLine("QUIT_FAILED " + ex.Message); }
                Release(app);
            }
        }
    }

    static LocalParagraphSnapshot P(int index, string text, string story = "wdMainTextStory",
        int? row = null, int? column = null) =>
        new LocalParagraphSnapshot(index, text, story, 1, index * 100, "Times New Roman",
            tableIndex: row.HasValue ? 1 : (int?)null, rowIndex: row, cellIndex: column,
            fontSizePoints: 13, bold: true, alignment: 1, isInTable: row.HasValue);

    static LocalScanSnapshot Snapshot(params LocalParagraphSnapshot[] paragraphs) =>
        new LocalScanSnapshot("audit", 1, Array.Empty<LocalSectionSnapshot>(), paragraphs,
            Array.Empty<AnnotationProtectedSpan>());

    static void Core()
    {
        var scanner = new LocalDocumentScanner();
        var rules = BundledRulePack.Load("1.0.0.192");
        var variants = new Dictionary<string, string>
        {
            ["canonical"] = Title + "\v" + Motto,
            ["NFD"] = (Title + "\v" + Motto).Normalize(NormalizationForm.FormD),
            ["nonbreaking-hyphen"] = Title + "\v" + Motto.Replace('-', '\u2011'),
            ["missing-separators"] = Title + "\vĐộc lập Tự do Hạnh phúc",
            ["zero-width"] = Title.Replace("CỘNG", "CỘNG\u200b") + "\v" + Motto
        };
        foreach (var pair in variants)
        {
            var snap = Snapshot(P(1, pair.Value), P(2, "QUYẾT ĐỊNH"));
            var roles = new DocumentRoleDetector().Detect(snap);
            Console.WriteLine("COMBINED case=" + pair.Key + " recognized=" +
                ComponentTextNormalizer.IsCombinedNationalHeader(pair.Value) + " breaks=" +
                CombinedNationalHeader.GetBreakOffsets(pair.Value).Length + " role=" +
                (roles.TryGetValue(1, out var role) ? role : "none") + " splitFinding=" +
                scanner.ScanFormat(snap, rules).Findings.Any(f => f.RuleCode == Rule));
        }
        foreach (var story in new[] { "wdMainTextStory", "wdPrimaryHeaderStory", "wdTextFrameStory" })
        {
            var roles = new DocumentRoleDetector().Detect(Snapshot(P(1, Title, story), P(2, Motto, story)));
            Console.WriteLine("STORY case=" + story + " title=" + roles.ContainsKey(1) + " motto=" + roles.ContainsKey(2));
        }
        var wrongCell = new DocumentRoleDetector().Detect(Snapshot(P(1, Title, row: 1, column: 1),
            P(2, Motto, row: 1, column: 2), P(3, "QUYẾT ĐỊNH")));
        Console.WriteLine("TABLE_ADJACENT_WRONG_COLUMN mottoRole=" + (wrongCell.TryGetValue(2, out var wrongRole) ? wrongRole : "none"));
        var gap = new List<LocalParagraphSnapshot> { P(1, "ỦY BAN NHÂN DÂN", row: 1, column: 1), P(2, Title, row: 1, column: 2) };
        for (var i = 3; i < 10; i++) gap.Add(P(i, "Cơ quan cấp trên " + i, row: 2, column: 1));
        gap.Add(P(10, Motto, row: 2, column: 2)); gap.Add(P(11, "QUYẾT ĐỊNH"));
        var gapRoles = new DocumentRoleDetector().Detect(Snapshot(gap.ToArray()));
        Console.WriteLine("TABLE_VALID_ROW_BELOW_GAP7 titleRole=" + gapRoles[2] + " mottoRole=" +
            (gapRoles.TryGetValue(10, out var gapRole) ? gapRole : "none"));
        foreach (var n in new[] { 500, 2000, 8000 })
        {
            var snap = Snapshot(Enumerable.Range(1, n).Select(i => P(i,
                "Nội dung kiểm tra tài liệu và hướng dẫn thực hiện công việc theo kế hoạch.")).ToArray());
            var watch = Stopwatch.StartNew(); var format = scanner.ScanFormat(snap, rules);
            var formatMs = watch.ElapsedMilliseconds; watch.Restart(); var spelling = scanner.ScanSpelling(snap, rules);
            Console.WriteLine("CORE_BENCH paragraphs=" + n + " format_ms=" + formatMs + " spelling_ms=" +
                watch.ElapsedMilliseconds + " findings=" + (format.Findings.Count + spelling.Findings.Count));
        }
    }

    static Word.Document Document(Word.Application app, string text)
    {
        var doc = app.Documents.Add(); var range = doc.Content;
        range.Text = text; range.Font.Name = "Times New Roman"; range.Font.Size = 13;
        range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        Release(range); doc.Activate(); return doc;
    }

    static void Profile()
    {
        var rules = BundledRulePack.Load("1.0.0.192");
        var paragraphs = Enumerable.Range(1, 1000).Select(i => P(i,
            "Nội dung kiểm tra tài liệu và hướng dẫn thực hiện công việc theo kế hoạch.")).ToArray();
        var snapshot = Snapshot(paragraphs); var roles = new DocumentRoleDetector().Detect(snapshot);
        var findings = new List<AnnotationFinding>(); var timer = Stopwatch.StartNew();
        typeof(CanonicalRuleScanner).GetMethod("CheckConfiguredCapitalizations", PrivateStatic)!
            .Invoke(null, new object?[] { findings, paragraphs, rules, roles, System.Threading.CancellationToken.None, null });
        Console.WriteLine("SPELLING_PHASE capitalization_rules=" + rules.Capitalizations.Count + " paragraphs=" +
            paragraphs.Length + " elapsed_ms=" + timer.ElapsedMilliseconds + " findings=" + findings.Count);
        timer.Restart(); var result = new LocalDocumentScanner().ScanSpelling(snapshot, rules);
        Console.WriteLine("SPELLING_FULL paragraphs=" + paragraphs.Length + " elapsed_ms=" + timer.ElapsedMilliseconds +
            " findings=" + result.Findings.Count + " first=" + result.Findings.FirstOrDefault()?.CurrentIssue);
    }

    static void EdgeCases()
    {
        var rules = BundledRulePack.Load("1.0.0.192"); var scanner = new LocalDocumentScanner();
        var paragraph = new LocalParagraphSnapshot(2, Motto, "wdMainTextStory", 1, 100, "Times New Roman",
            fontSizePoints: 13, bold: true, alignment: 1, pageNumber: 1,
            pageLeftPoints: 200, pageTopPoints: 100, textWidthPoints: 180);
        LocalLineShapeSnapshot Line(string name, double top, int dash) =>
            new LocalLineShapeSnapshot(1, name, 9, "wdMainTextStory", 1, 100, 2, 1,
                200, top, 180, 0, 200, top, 1, 1, true, dash, .75, 0, 1, 1);
        var old = Line("OldDashedSeparator", 150, 4); var added = Line("CHUANHOA2_MOTTO_P2", 119, 1);
        foreach (var lines in new[] { new[] { old }, new[] { old, added } })
        {
            var snapshot = new LocalScanSnapshot("audit-lines", 1, Array.Empty<LocalSectionSnapshot>(),
                new[] { P(1, Title), paragraph, P(3, "QUYẾT ĐỊNH") }, Array.Empty<AnnotationProtectedSpan>(),
                lineShapes: lines);
            var findings = scanner.ScanFormat(snapshot, rules).Findings.Where(f => f.RuleCode == Rule).ToArray();
            Console.WriteLine("LINE_RESIDUE shapes=" + lines.Length + " lineFindings=" + findings.Length +
                " issues=" + string.Join("|", findings.Select(f => f.CurrentIssue)));
        }
        var many = Snapshot(Enumerable.Range(1, 1000).Select(i => P(i,
            "Nội dung kiểm tra tài liệu và hướng dẫn thực hiện công việc theo kế hoạch.")).ToArray());
        using (var cancellation = new System.Threading.CancellationTokenSource())
        {
            var timer = Stopwatch.StartNew(); long requestMs = -1;
            cancellation.Token.Register(() => System.Threading.Interlocked.Exchange(ref requestMs, timer.ElapsedMilliseconds));
            cancellation.CancelAfter(100);
            try { scanner.ScanSpelling(many, rules, cancellation.Token); Console.WriteLine("CANCEL completedWithoutCancel"); }
            catch (OperationCanceledException)
            { Console.WriteLine("CANCEL request_ms=" + requestMs + " returned_ms=" + timer.ElapsedMilliseconds); }
        }
    }

    static WordDocumentSnapshot Capture(Word.Application app, Word.Document doc)
    {
        var watch = Stopwatch.StartNew();
        var builder = new WordDocumentSnapshotBuilder((stage, ms) =>
        { if (ms >= 100) Console.WriteLine("STAGE " + stage + " ms=" + ms); });
        var result = builder.Build(doc, new DocumentContext(1), new WordDocumentCapabilityProvider(app).Evaluate(doc));
        Console.WriteLine("CAPTURE paragraphs=" + result.Paragraphs.Count + " total_ms=" + watch.ElapsedMilliseconds);
        return result;
    }

    static void LateBorder(Word.Application app)
    {
        var body = Title + "\r" + Motto + "\rQUYẾT ĐỊNH\rVề việc kiểm tra\r" +
            string.Concat(Enumerable.Repeat("Nội dung kiểm tra tài liệu.\r", 90)) +
            Title + "\r" + Motto + "\rQUYẾT ĐỊNH\rVề việc kiểm tra lần hai\r";
        var doc = Document(app, body);
        try
        {
            var offset = body.LastIndexOf(Motto, StringComparison.Ordinal);
            var range = doc.Range(offset, offset + Motto.Length);
            range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            var before = (int)range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle;
            var snap = WordLocalScanRuntime.ToLocalSnapshot(Capture(app, doc));
            var motto = snap.Paragraphs.Last(p => p.Text == Motto);
            var roles = new DocumentRoleDetector().Detect(snap);
            var result = Normalize(doc, snap, motto);
            Console.WriteLine("LATE_BORDER paragraph=" + motto.Index + " detectedRole=" + roles[motto.Index] +
                " actualBefore=" + before + " snapshotHasBorder=" + motto.HasBottomBorder + " normalized=" + result +
                " actualAfter=" + (int)range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle +
                " shapesAfter=" + doc.Shapes.Count);
            Release(range);
        }
        finally { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); }
    }

    static bool Normalize(Word.Document doc, LocalScanSnapshot snap, LocalParagraphSnapshot paragraph) =>
        (bool)typeof(WordOneClickRuntime).GetMethods(PrivateStatic)
            .Single(m => m.Name == "NormalizeRequiredLine" && m.GetParameters().Length == 6)
            .Invoke(null, new object[] { doc, snap, paragraph, 1d, Rule, new Dictionary<string, string>() });

    static void OldLine(Word.Application app)
    {
        var doc = Document(app, Title + "\r" + Motto + "\rQUYẾT ĐỊNH\rVề việc kiểm tra\r");
        try
        {
            var first = WordLocalScanRuntime.ToLocalSnapshot(Capture(app, doc));
            var motto = first.Paragraphs.Single(p => p.Text == Motto);
            var anchor = doc.Range(motto.AbsoluteStart, motto.AbsoluteEnd); object a = anchor;
            var left = (float)motto.PageLeftPoints!.Value;
            var y = (float)(motto.PageTopPoints!.Value + 50);
            var shape = doc.Shapes.AddLine(left, y, left + (float)motto.TextWidthPoints!.Value, y, ref a);
            shape.Name = "AuditOldSeparator";
            shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            shape.Left = left; shape.Top = y; shape.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            shape.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            Release(shape); Release(anchor);
            var snap = WordLocalScanRuntime.ToLocalSnapshot(Capture(app, doc));
            var rules = BundledRulePack.Load("1.0.0.192"); var scanner = new LocalDocumentScanner();
            Console.WriteLine("OLD_LINE beforeFinding=" + string.Join("|", scanner.ScanFormat(snap, rules).Findings
                .Where(f => f.RuleCode == Rule).Select(f => f.CurrentIssue)));
            Normalize(doc, snap, snap.Paragraphs.Single(p => p.Text == Motto));
            var after = WordLocalScanRuntime.ToLocalSnapshot(Capture(app, doc));
            Console.WriteLine("OLD_LINE afterShapes=" + after.LineShapes.Count + " oldStillExists=" +
                after.LineShapes.Any(l => l.Name == "AuditOldSeparator") + " remainingLineFindings=" +
                scanner.ScanFormat(after, rules).Findings.Count(f => f.RuleCode == Rule));
        }
        finally { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); }
    }

    static void SafetyGuard(Word.Application app)
    {
        var doc = Document(app, Title + "\r" + Motto + "\rQUYẾT ĐỊNH\rVề việc kiểm tra\r" +
            string.Concat(Enumerable.Repeat("Nội dung tài liệu dài để kiểm tra giới hạn an toàn khi lấy tọa độ trang và đo đường kẻ.\r", 2600)));
        try
        {
            var allowed = typeof(WordDocumentSnapshotBuilder).GetMethod("IsPageLayoutCaptureSafe", PrivateStatic)!
                .Invoke(null, new object[] { doc });
            var result = Capture(app, doc); var motto = result.Paragraphs.Single(p => p.Text == Motto);
            Console.WriteLine("SAFETY_GUARD chars=" + doc.Content.End + " allowLayout=" + allowed +
                " mottoPage=" + motto.PageNumber + " hasPageTop=" + motto.PageTopPoints.HasValue +
                " hasTextWidth=" + motto.TextWidthPoints.HasValue);
        }
        finally { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); }
    }

    static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
