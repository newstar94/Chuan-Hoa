using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>
    /// A horizontal rule drawn with Insert &gt; Shapes &gt; Line. Word 2007 reports
    /// msoLine (9); Word 2013+ can report the same drawing as an AutoShape (1),
    /// usually a straight connector. Both are separators for NĐ30 purposes.
    /// </summary>
    internal static class WordLineShapeKind
    {
        public const int MsoLine = 9;
        private const int MsoAutoShape = 1;

        public static bool IsLineLike(Word.Shape shape)
        {
            if (shape == null) return false;
            int type;
            try { type = (int)shape.Type; }
            catch (COMException) { return false; }
            if (type == MsoLine) return true;
            if (type != MsoAutoShape) return false;
            try
            {
                // A text box or any shape carrying text is never a separator.
                Word.TextFrame? frame = null;
                try
                {
                    frame = shape.TextFrame;
                    if (frame.HasText != 0) return false;
                }
                catch (COMException) { }
                finally { Release(frame); }
                var height = Math.Abs(shape.Height);
                var width = Math.Abs(shape.Width);
                return height <= 1.5f && width >= 10f;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static void Release(object? value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
