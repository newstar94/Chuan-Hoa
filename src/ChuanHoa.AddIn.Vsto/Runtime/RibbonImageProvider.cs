using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    internal sealed class RibbonImageProvider : IDisposable
    {
        private const string PersonalDictionaryResource =
            "ChuanHoa.AddIn.Vsto.Resources.Icons.personal-dictionary-32.png";
        private readonly Dictionary<string, Bitmap> _bitmaps =
            new Dictionary<string, Bitmap>(StringComparer.Ordinal);
        private readonly Dictionary<string, object> _pictures =
            new Dictionary<string, object>(StringComparer.Ordinal);
        private bool _disposed;

        public object GetImage(string controlId)
        {
            var scale = controlId == "btnScaleGiam" || controlId == "btnScale100" || controlId == "btnScaleTang" ||
                controlId == "btnCoChu" || controlId == "btnGianChuNormal" || controlId == "btnGianChuRa";
            if (_disposed || (!scale && !string.Equals(controlId, "btnTuDienCaNhan", StringComparison.Ordinal)))
                return null!;
            object picture;
            if (_pictures.TryGetValue(controlId, out picture)) return picture;
            try
            {
                var bitmap = scale ? DrawScaleIcon(controlId) : LoadBitmap(PersonalDictionaryResource);
                picture = PictureDispConverter.ToPictureDisp(bitmap);
                _bitmaps.Add(controlId, bitmap);
                _pictures.Add(controlId, picture);
                return picture;
            }
            catch (Exception exception)
            {
                Trace.TraceError("ChuanHoa Ribbon image load failed for {0}: {1}", controlId, exception);
                return null!;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            foreach (var bitmap in _bitmaps.Values) bitmap.Dispose();
            _pictures.Clear();
            _bitmaps.Clear();
            _disposed = true;
        }

        private static Bitmap DrawScaleIcon(string id)
        {
            if (id == "btnGianChuNormal") return DrawNormalIcon();
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            using (var font = new Font("Segoe UI", id == "btnScale100" ? 10f : 25f, FontStyle.Regular, GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(SystemColors.ControlText))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                graphics.Clear(SystemColors.Control);
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                format.FormatFlags = StringFormatFlags.NoWrap;
                if (id == "btnGianChuNormal")
                    graphics.FillEllipse(brush, 10, 10, 12, 12);
                else if (id == "btnCoChu" || id == "btnGianChuRa")
                    graphics.FillPolygon(brush, id == "btnCoChu"
                        ? new[] { new Point(10, 16), new Point(22, 9), new Point(22, 23) }
                        : new[] { new Point(22, 16), new Point(10, 9), new Point(10, 23) });
                else if (id == "btnScale100")
                    graphics.DrawString("100%", font, brush, new RectangleF(0, 0, 32, 32), format);
                else
                {
                    graphics.DrawString("A", font, brush, new RectangleF(-2, 0, 22, 32), format);
                    using (var signFont = new Font("Segoe UI", 19f, FontStyle.Regular, GraphicsUnit.Pixel))
                        graphics.DrawString(id == "btnScaleGiam" ? "−" : "+", signFont, brush,
                            new RectangleF(17, 0, 16, 32), format);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawNormalIcon()
        {
            // A square source remains circular in Office's square icon slot.
            // Preserve a 12px circle inside a native 16px small-button image.
            var bitmap = new Bitmap(16, 16);
            using (var graphics = Graphics.FromImage(bitmap))
            using (var brush = new SolidBrush(SystemColors.ControlText))
            {
                graphics.Clear(Color.Transparent);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.FillEllipse(brush, 2, 2, 12, 12);
            }
            return bitmap;
        }

        private static Bitmap LoadBitmap(string resourceName)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new FileNotFoundException("Embedded Ribbon image is missing.", resourceName);
                using (var source = new Bitmap(stream)) return new Bitmap(source);
            }
        }

        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter() : base(string.Empty) { }

            public static object ToPictureDisp(Image image) => GetIPictureDispFromPicture(image);
        }
    }
}
