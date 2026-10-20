using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;

namespace ChuanHoa.Client.Core.Annotations
{
    /// <summary>
    /// Builds the local, self-contained HTML report that lists every finding of a
    /// scan, including those not turned into Word comments because of the D2
    /// comment budget. The report is written only to the user's machine.
    /// </summary>
    public static class AnnotationReportBuilder
    {
        public static string BuildHtml(string documentName, string lane, DateTime generatedAt,
            IReadOnlyList<AnnotationFinding> findings, AnnotationPlan plan)
        {
            if (findings == null) throw new ArgumentNullException(nameof(findings));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            var commented = new HashSet<string>(plan.Comments.Select(item => item.FindingId),
                StringComparer.Ordinal);
            var omitted = new HashSet<string>(plan.Omitted.Select(item => item.FindingId),
                StringComparer.Ordinal);
            var laneTitle = string.Equals(lane, "spelling", StringComparison.OrdinalIgnoreCase)
                ? "Kiểm tra chính tả"
                : "Kiểm tra thể thức";

            var builder = new StringBuilder();
            builder.AppendLine("<!doctype html>");
            builder.AppendLine("<html lang=\"vi\"><head><meta charset=\"utf-8\">");
            builder.Append("<title>").Append(Encode(laneTitle)).AppendLine(" – Chuẩn hóa</title>");
            builder.AppendLine("<style>body{font-family:'Times New Roman',serif;margin:24px;color:#111;background:#fff}" +
                "table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:4px 6px;vertical-align:top;font-size:13px}" +
                "th{background:#eee;text-align:left}tr.omitted td{background:#fff8e1}.muted{color:#555}</style>");
            builder.AppendLine("</head><body>");
            builder.Append("<h1>").Append(Encode(laneTitle)).AppendLine("</h1>");
            builder.Append("<p class=\"muted\">Tài liệu: ").Append(Encode(documentName ?? string.Empty))
                .Append(" · Lập lúc: ")
                .Append(Encode(generatedAt.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture)))
                .AppendLine("</p>");
            builder.Append("<p>Tổng số điểm: <b>").Append(findings.Count.ToString(CultureInfo.InvariantCulture))
                .Append("</b>. Đã gắn ").Append(plan.Comments.Count.ToString(CultureInfo.InvariantCulture))
                .Append(" comment cho ").Append(plan.AnnotatedFindingCount.ToString(CultureInfo.InvariantCulture))
                .Append(" điểm; ").Append(plan.Omitted.Count.ToString(CultureInfo.InvariantCulture))
                .Append(" điểm chỉ có trong báo cáo này (tô nền vàng); ")
                .Append(plan.Unresolved.Count.ToString(CultureInfo.InvariantCulture))
                .AppendLine(" điểm không neo được vào văn bản.</p>");

            builder.AppendLine("<h2>Thống kê theo mã quy tắc</h2>");
            builder.AppendLine("<table><tr><th>Mã quy tắc</th><th>Số điểm</th></tr>");
            foreach (var rule in findings.GroupBy(item => item.RuleCode, StringComparer.Ordinal)
                .OrderByDescending(group => group.Count()).ThenBy(group => group.Key, StringComparer.Ordinal))
            {
                builder.Append("<tr><td>").Append(Encode(rule.Key)).Append("</td><td>")
                    .Append(rule.Count().ToString(CultureInfo.InvariantCulture)).AppendLine("</td></tr>");
            }
            builder.AppendLine("</table>");

            builder.AppendLine("<h2>Danh sách đầy đủ</h2>");
            builder.AppendLine("<table><tr><th>#</th><th>Đoạn</th><th>Mã</th><th>Hiện tại</th><th>Yêu cầu đúng</th><th>Đánh dấu</th></tr>");
            var number = 0;
            foreach (var finding in findings
                .OrderBy(item => item.Anchor.ParagraphIndex ?? int.MaxValue)
                .ThenBy(item => item.Anchor.StartOffset ?? 0))
            {
                number++;
                var state = commented.Contains(finding.FindingId) ? "comment"
                    : omitted.Contains(finding.FindingId) ? "chỉ trong báo cáo"
                    : "gộp/không neo";
                builder.Append(omitted.Contains(finding.FindingId) ? "<tr class=\"omitted\">" : "<tr>")
                    .Append("<td>").Append(number.ToString(CultureInfo.InvariantCulture)).Append("</td>")
                    .Append("<td>").Append(finding.Anchor.ParagraphIndex.HasValue
                        ? finding.Anchor.ParagraphIndex.Value.ToString(CultureInfo.InvariantCulture)
                        : "–").Append("</td>")
                    .Append("<td>").Append(Encode(finding.RuleCode)).Append("</td>")
                    .Append("<td>").Append(Encode(finding.CurrentIssue)).Append("</td>")
                    .Append("<td>").Append(Encode(finding.Expected)).Append("</td>")
                    .Append("<td>").Append(Encode(state)).AppendLine("</td></tr>");
            }
            builder.AppendLine("</table></body></html>");
            return builder.ToString();
        }

        private static string Encode(string value) => WebUtility.HtmlEncode(value ?? string.Empty);
    }
}
