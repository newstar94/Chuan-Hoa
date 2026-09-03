#nullable disable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.WordRuleAudit
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length != 2) return 2;
            var input = Path.GetFullPath(args[0]);
            var output = Path.GetFullPath(args[1]);
            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                object fileName = input;
                object readOnly = true;
                object addToRecent = false;
                object visible = false;
                document = application.Documents.Open(ref fileName, ReadOnly: ref readOnly, AddToRecentFiles: ref addToRecent, Visible: ref visible);
                var report = Audit(document, input);
                Directory.CreateDirectory(Path.GetDirectoryName(output));
                var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue, RecursionLimit = 100 };
                File.WriteAllText(output, serializer.Serialize(report), new UTF8Encoding(false));
                Console.WriteLine("WORD_RULE_AUDIT_PASS " + Path.GetFileName(input));
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("WORD_RULE_AUDIT_FAIL " + exception);
                return 1;
            }
            finally
            {
                if (document != null)
                {
                    object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                    document.Close(ref save);
                }
                if (application != null) application.Quit();
                Release(document);
                Release(application);
            }
        }

        private static Dictionary<string, object> Audit(Word.Document document, string path)
        {
            var roles = new List<object>();
            var bottomBorders = new List<object>();
            var styleCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            var paragraphs = document.Paragraphs;
            try
            {
                for (var index = 1; index <= paragraphs.Count; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    Word.Font font = null;
                    Word.ParagraphFormat format = null;
                    Word.Borders borders = null;
                    Word.Border bottom = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        range = paragraph.Range.Duplicate;
                        var text = Clean(range.Text);
                        if (text.Length == 0) continue;
                        font = range.Font;
                        format = range.ParagraphFormat;
                        var role = DetectRole(text);
                        var styleName = ReadStyleName(range);
                        var styleKey = styleName + "|" + ReadName(font) + "|" + Number(font.Size) + "|" + font.Bold + "|" + font.Italic + "|" + (int)format.Alignment;
                        styleCounts[styleKey] = styleCounts.ContainsKey(styleKey) ? styleCounts[styleKey] + 1 : 1;
                        borders = format.Borders;
                        bottom = borders[Word.WdBorderType.wdBorderBottom];
                        var hasBottom = bottom.LineStyle != Word.WdLineStyle.wdLineStyleNone;
                        if (hasBottom)
                        {
                            bottomBorders.Add(ParagraphRecord(index, text, role, range, font, format, bottom));
                        }
                        if (role != "other" && roles.Count(item => ((Dictionary<string, object>)item)["role"].ToString() == role) < 30)
                        {
                            roles.Add(ParagraphRecord(index, text, role, range, font, format, bottom));
                        }
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        Release(bottom);
                        Release(borders);
                        Release(format);
                        Release(font);
                        Release(range);
                        Release(paragraph);
                    }
                }
            }
            finally
            {
                Release(paragraphs);
            }
            return new Dictionary<string, object>
            {
                ["sourcePath"] = path,
                ["sourceLength"] = new FileInfo(path).Length,
                ["sourceLastWriteUtc"] = File.GetLastWriteTimeUtc(path).ToString("O", CultureInfo.InvariantCulture),
                ["pageCount"] = document.ComputeStatistics(Word.WdStatistic.wdStatisticPages),
                ["paragraphCount"] = document.Paragraphs.Count,
                ["sectionCount"] = document.Sections.Count,
                ["tableCount"] = document.Tables.Count,
                ["shapeCount"] = document.Shapes.Count,
                ["inlineShapeCount"] = document.InlineShapes.Count,
                ["sections"] = AuditSections(document),
                ["roleSamples"] = roles,
                ["bottomBorderParagraphs"] = bottomBorders,
                ["shapes"] = AuditShapes(document),
                ["styleClusters"] = styleCounts.OrderByDescending(item => item.Value).Take(100)
                    .Select(item => new Dictionary<string, object> { ["key"] = item.Key, ["count"] = item.Value }).ToArray()
            };
        }

        private static Dictionary<string, object> ParagraphRecord(int index, string text, string role, Word.Range range,
            Word.Font font, Word.ParagraphFormat format, Word.Border bottom)
        {
            return new Dictionary<string, object>
            {
                ["index"] = index,
                ["role"] = role,
                ["text"] = text.Length > 500 ? text.Substring(0, 500) : text,
                ["style"] = ReadStyleName(range),
                ["font"] = ReadName(font),
                ["fontSize"] = Number(font.Size),
                ["bold"] = font.Bold,
                ["italic"] = font.Italic,
                ["underline"] = (int)font.Underline,
                ["fontColor"] = (int)font.Color,
                ["alignment"] = (int)format.Alignment,
                ["firstLineIndent"] = Number(format.FirstLineIndent),
                ["leftIndent"] = Number(format.LeftIndent),
                ["rightIndent"] = Number(format.RightIndent),
                ["spaceBefore"] = Number(format.SpaceBefore),
                ["spaceAfter"] = Number(format.SpaceAfter),
                ["lineSpacing"] = Number(format.LineSpacing),
                ["lineSpacingRule"] = (int)format.LineSpacingRule,
                ["section"] = SafeInformation(range, Word.WdInformation.wdActiveEndSectionNumber),
                ["page"] = SafeInformation(range, Word.WdInformation.wdActiveEndAdjustedPageNumber),
                ["inTable"] = SafeInformation(range, Word.WdInformation.wdWithInTable) != 0,
                ["bottomBorderStyle"] = (int)bottom.LineStyle,
                ["bottomBorderWidth"] = (int)bottom.LineWidth,
                ["bottomBorderColor"] = (int)bottom.Color,
                ["underlineSegments"] = UnderlineSegments(range)
            };
        }

        private static object[] UnderlineSegments(Word.Range range)
        {
            var result = new List<object>();
            Word.Words words = null;
            try
            {
                words = range.Words;
                for (var index = 1; index <= words.Count && result.Count < 20; index++)
                {
                    Word.Range word = null;
                    Word.Font font = null;
                    try
                    {
                        word = words[index];
                        font = word.Font;
                        if (font.Underline != Word.WdUnderline.wdUnderlineNone)
                        {
                            result.Add(new Dictionary<string, object>
                            {
                                ["text"] = Clean(word.Text), ["underline"] = (int)font.Underline
                            });
                        }
                    }
                    finally
                    {
                        Release(font);
                        Release(word);
                    }
                }
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(words);
            }
            return result.ToArray();
        }

        private static object[] AuditSections(Word.Document document)
        {
            var result = new List<object>();
            for (var index = 1; index <= document.Sections.Count; index++)
            {
                Word.Section section = null;
                Word.PageSetup setup = null;
                try
                {
                    section = document.Sections[index];
                    setup = section.PageSetup;
                    result.Add(new Dictionary<string, object>
                    {
                        ["index"] = index,
                        ["pageWidth"] = Number(setup.PageWidth), ["pageHeight"] = Number(setup.PageHeight),
                        ["topMargin"] = Number(setup.TopMargin), ["bottomMargin"] = Number(setup.BottomMargin),
                        ["leftMargin"] = Number(setup.LeftMargin), ["rightMargin"] = Number(setup.RightMargin),
                        ["orientation"] = (int)setup.Orientation,
                        ["differentFirstPage"] = setup.DifferentFirstPageHeaderFooter,
                        ["oddAndEvenPages"] = setup.OddAndEvenPagesHeaderFooter,
                        ["pageNumbering"] = PageNumbering(section)
                    });
                }
                finally
                {
                    Release(setup);
                    Release(section);
                }
            }
            return result.ToArray();
        }

        private static object[] PageNumbering(Word.Section section)
        {
            var result = new List<object>();
            foreach (var pair in new[] { Tuple.Create("header", section.Headers), Tuple.Create("footer", section.Footers) })
            {
                var collection = pair.Item2;
                try
                {
                    foreach (Word.WdHeaderFooterIndex kind in new[] { Word.WdHeaderFooterIndex.wdHeaderFooterPrimary, Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage, Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages })
                    {
                        Word.HeaderFooter item = null;
                        Word.PageNumbers numbers = null;
                        try
                        {
                            item = collection[kind];
                            if (!item.Exists) continue;
                            numbers = item.PageNumbers;
                            if (numbers.Count > 0)
                            {
                                result.Add(new Dictionary<string, object>
                                {
                                    ["location"] = pair.Item1, ["kind"] = (int)kind, ["count"] = numbers.Count,
                                    ["restart"] = numbers.RestartNumberingAtSection, ["startingNumber"] = numbers.StartingNumber
                                });
                            }
                        }
                        finally
                        {
                            Release(numbers);
                            Release(item);
                        }
                    }
                }
                finally
                {
                    Release(collection);
                }
            }
            return result.ToArray();
        }

        private static object[] AuditShapes(Word.Document document)
        {
            var result = new List<object>();
            for (var index = 1; index <= document.Shapes.Count && result.Count < 500; index++)
            {
                Word.Shape shape = null;
                Word.Range anchor = null;
                try
                {
                    shape = document.Shapes[index];
                    anchor = shape.Anchor;
                    result.Add(new Dictionary<string, object>
                    {
                        ["index"] = index, ["name"] = shape.Name, ["type"] = (int)shape.Type,
                        ["width"] = Number(shape.Width), ["height"] = Number(shape.Height),
                        ["left"] = Number(shape.Left), ["top"] = Number(shape.Top),
                        ["anchorStart"] = anchor.Start, ["anchorText"] = Clean(anchor.Text),
                        ["anchorPage"] = SafeInformation(anchor, Word.WdInformation.wdActiveEndAdjustedPageNumber)
                    });
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(anchor);
                    Release(shape);
                }
            }
            return result.ToArray();
        }

        private static string DetectRole(string text)
        {
            var normalized = Regex.Replace(text, @"\s+", " ").Trim();
            if (normalized.IndexOf("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", StringComparison.OrdinalIgnoreCase) >= 0) return "nationalTitle";
            if (normalized.IndexOf("Độc lập", StringComparison.OrdinalIgnoreCase) >= 0 && normalized.IndexOf("Hạnh phúc", StringComparison.OrdinalIgnoreCase) >= 0) return "nationalMotto";
            if (normalized.IndexOf("ĐẢNG CỘNG SẢN VIỆT NAM", StringComparison.OrdinalIgnoreCase) >= 0) return "partyTitle";
            if (Regex.IsMatch(normalized, @"^Số\s*:?", RegexOptions.IgnoreCase)) return "codeNumber";
            if (Regex.IsMatch(normalized, @"^[^,]{2,80},?\s*ngày\s+\d{1,2}\s+tháng", RegexOptions.IgnoreCase)) return "placeAndIssuedDate";
            if (Regex.IsMatch(normalized, @"^(NGHỊ QUYẾT|QUYẾT ĐỊNH|CHỈ THỊ|THÔNG TƯ|THÔNG BÁO|KẾ HOẠCH|BÁO CÁO|TỜ TRÌNH|HƯỚNG DẪN)$", RegexOptions.IgnoreCase)) return "typeName";
            if (Regex.IsMatch(normalized, @"^(V/v|Về việc)\b", RegexOptions.IgnoreCase)) return "subject";
            if (Regex.IsMatch(normalized, @"^(Căn cứ|Xét đề nghị)\b", RegexOptions.IgnoreCase)) return "legalBasis";
            if (Regex.IsMatch(normalized, @"^(TM\.|KT\.|TL\.|TUQ\.|Q\.|CHỦ TỊCH|GIÁM ĐỐC|BỘ TRƯỞNG|TỔNG GIÁM ĐỐC)\b", RegexOptions.IgnoreCase)) return "signerAuthority";
            if (Regex.IsMatch(normalized, @"^Kính\s+(gửi|trình)", RegexOptions.IgnoreCase)) return "recipientSalutation";
            if (Regex.IsMatch(normalized, @"^Nơi\s+nhận", RegexOptions.IgnoreCase)) return "recipientLabel";
            if (Regex.IsMatch(normalized, @"^Phụ\s+lục(?:\s+[IVXLCDM\d]+)?\b", RegexOptions.IgnoreCase)) return "appendixLabel";
            if (Regex.IsMatch(normalized, @"^Điều\s+\d+", RegexOptions.IgnoreCase)) return "article";
            return "other";
        }

        private static string ReadStyleName(Word.Range range)
        {
            object style = null;
            try
            {
                style = range.get_Style();
                var wordStyle = style as Word.Style;
                return wordStyle == null ? Convert.ToString(style, CultureInfo.InvariantCulture) : wordStyle.NameLocal;
            }
            catch (COMException)
            {
                return string.Empty;
            }
            finally
            {
                Release(style);
            }
        }

        private static int SafeInformation(Word.Range range, Word.WdInformation information)
        {
            try { return Convert.ToInt32(range.get_Information(information), CultureInfo.InvariantCulture); }
            catch (COMException) { return 0; }
        }

        private static string ReadName(Word.Font font)
        {
            try { return font.Name ?? string.Empty; }
            catch (COMException) { return string.Empty; }
        }

        private static object Number(float value)
        {
            return value > 999998f || value < -999998f ? null : (object)Math.Round(value, 3);
        }

        private static string Clean(string value)
        {
            return (value ?? string.Empty).TrimEnd('\r', '\a').Replace("\u000b", " ").Trim();
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
