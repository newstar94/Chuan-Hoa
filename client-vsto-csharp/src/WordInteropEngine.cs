using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VietDocStandardizer
{
    public static class WordInteropEngine
    {
        public static string CurrentRegime { get; set; } = "ND30";

        private static Word.Application WordApp => ThisAddIn.Instance?.Application;

        /// <summary>
        /// Thực hiện 1-Click Auto-Fix toàn diện tài liệu
        /// </summary>
        public static void Perform1ClickAutoFix()
        {
            try
            {
                Word.Document doc = WordApp?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("Không có tài liệu Word nào đang mở!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                WordApp.ScreenUpdating = false;

                // 1. Căn lề trang A4
                ApplyPageSetup(doc);

                // 2. Định dạng toàn bộ phông chữ Times New Roman, màu đen
                doc.Content.Font.Name = "Times New Roman";
                doc.Content.Font.Color = Word.WdColor.wdColorBlack;
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                // 3. Chuẩn hóa từng đoạn văn bản
                int count = doc.Paragraphs.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Paragraph p = doc.Paragraphs[i];
                    string text = p.Range.Text.Trim();
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    string upper = text.ToUpper();

                    // Quốc hiệu
                    if (upper.Contains("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") || upper.Contains("CONG HOA XA HOI"))
                    {
                        p.Range.Font.Size = 12;
                        p.Range.Font.Bold = 1;
                        p.Range.Font.Italic = 0;
                        p.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        p.SpaceBefore = 0;
                        p.SpaceAfter = 0;
                    }
                    // Tiêu ngữ
                    else if (upper.Contains("ĐỘC LẬP - TỰ DO - HẠNH PHÚC") || upper.Contains("ĐỘC LẬP-TỰ DO-HẠNH PHÚC"))
                    {
                        p.Range.Font.Size = 13;
                        p.Range.Font.Bold = 1;
                        p.Range.Font.Italic = 0;
                        p.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        p.SpaceBefore = 0;
                        p.SpaceAfter = 0;
                    }
                    // Tiêu đề Đảng
                    else if (upper.Contains("ĐẢNG CỘNG SẢN VIỆT NAM"))
                    {
                        p.Range.Font.Size = 15;
                        p.Range.Font.Bold = 1;
                        p.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    // Tên loại văn bản (QUYẾT ĐỊNH, CHỈ THỊ...)
                    else if (Regex.IsMatch(upper, @"^(QUYẾT ĐỊNH|CHỈ THỊ|THÔNG TƯ|NGHỊ QUYẾT|BÁO CÁO|KẾ HOẠCH|QUY ĐỊNH|TỜ TRÌNH)$"))
                    {
                        p.Range.Font.Size = (CurrentRegime == "DANG_HD05") ? 15 : 14;
                        p.Range.Font.Bold = 1;
                        p.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        p.SpaceBefore = 12;
                        p.SpaceAfter = 0;
                    }
                    // Nội dung văn bản chính
                    else if (i > 5)
                    {
                        p.Range.Font.Size = 14;
                        p.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        p.FirstLineIndent = WordApp.CentimetersToPoints(1.0f);
                        p.SpaceAfter = 6;
                        
                        if (CurrentRegime == "DANG_HD05")
                        {
                            p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                            p.LineSpacing = 20; // 18 - 22 pt
                        }
                        else
                        {
                            p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                            p.LineSpacing = 14f; // 1.15 lines
                        }
                    }
                }

                // 4. Định dạng tất cả Bảng biểu
                RepeatTableHeaders();
                PreventTableRowSplitting();

                // 5. Xóa trang trắng cuối
                RemoveTrailingEmptyPages();

                MessageBox.Show("✅ Đã hoàn tất 1-Click Auto-Fix chuẩn hóa toàn bộ văn bản!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thực hiện Auto-Fix: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (WordApp != null) WordApp.ScreenUpdating = true;
            }
        }

        private static void ApplyPageSetup(Word.Document doc)
        {
            foreach (Word.Section section in doc.Sections)
            {
                Word.PageSetup ps = section.PageSetup;
                ps.PaperSize = Word.WdPaperSize.wdPaperA4;
                ps.Orientation = Word.WdOrientation.wdOrientPortrait;
                // Margins: Top 20mm, Bottom 20mm, Left 30mm, Right 15mm
                ps.TopMargin = WordApp.CentimetersToPoints(2.0f);
                ps.BottomMargin = WordApp.CentimetersToPoints(2.0f);
                ps.LeftMargin = WordApp.CentimetersToPoints(3.0f);
                ps.RightMargin = WordApp.CentimetersToPoints(1.5f);
            }
        }

        public static void RepeatTableHeaders()
        {
            Word.Document doc = WordApp?.ActiveDocument;
            if (doc == null) return;

            foreach (Word.Table table in doc.Tables)
            {
                if (table.Rows.Count > 0)
                {
                    table.Rows[1].HeadingFormat = -1; // True
                }
            }
        }

        public static void PreventTableRowSplitting()
        {
            Word.Document doc = WordApp?.ActiveDocument;
            if (doc == null) return;

            foreach (Word.Table table in doc.Tables)
            {
                foreach (Word.Row row in table.Rows)
                {
                    row.AllowBreakAcrossPages = 0; // False
                }
            }
        }

        public static void RemoveTrailingEmptyPages()
        {
            Word.Document doc = WordApp?.ActiveDocument;
            if (doc == null) return;

            while (doc.Paragraphs.Count > 0)
            {
                Word.Paragraph lastP = doc.Paragraphs[doc.Paragraphs.Count];
                if (string.IsNullOrWhiteSpace(lastP.Range.Text) || lastP.Range.Text == "\r")
                {
                    lastP.Range.Delete();
                }
                else
                {
                    break;
                }
            }
        }

        public static void NormalizeToneMarks()
        {
            Word.Document doc = WordApp?.ActiveDocument;
            if (doc == null) return;

            ReplaceWordText(doc, "hoà", "hòa");
            ReplaceWordText(doc, "hoá", "hóa");
            ReplaceWordText(doc, "thuý", "thúy");
            ReplaceWordText(doc, "thuỷ", "thủy");
            MessageBox.Show("Đã chuẩn hóa vị trí đặt dấu thanh tiếng Việt!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void NormalizeIySpellings()
        {
            Word.Document doc = WordApp?.ActiveDocument;
            if (doc == null) return;

            ReplaceWordText(doc, "kĩ thuật", "kỹ thuật");
            ReplaceWordText(doc, "xử lí", "xử lý");
            ReplaceWordText(doc, "quản lí", "quản lý");
            ReplaceWordText(doc, "qui định", "quy định");
            MessageBox.Show("Đã chuẩn hóa chính tả i/y theo chuẩn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void CleanSpacesAndPunctuation()
        {
            Word.Document doc = WordApp?.ActiveDocument;
            if (doc == null) return;

            ReplaceWordText(doc, "  ", " ");
            ReplaceWordText(doc, " ,", ",");
            ReplaceWordText(doc, " .", ".");
            ReplaceWordText(doc, " :", ":");
            ReplaceWordText(doc, " ;", ";");
            MessageBox.Show("Đã dọn sạch khoảng trắng thừa và chuẩn hóa dấu câu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void ReplaceWordText(Word.Document doc, string find, string replace)
        {
            Word.Find finder = doc.Content.Find;
            finder.ClearFormatting();
            finder.Replacement.ClearFormatting();
            finder.Text = find;
            finder.Replacement.Text = replace;
            finder.Forward = true;
            finder.Wrap = Word.WdFindWrap.wdFindContinue;
            finder.MatchCase = false;
            finder.MatchWholeWord = false;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }
    }
}
