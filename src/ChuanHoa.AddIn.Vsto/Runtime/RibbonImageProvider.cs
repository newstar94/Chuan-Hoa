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
            var scale = controlId == "btnScaleGiam" || controlId == "btnScale100" || controlId == "btnScaleTang";
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
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            using (var font = new Font("Segoe UI", id == "btnScale100" ? 10f : 16f, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(Color.FromArgb(25, 85, 145)))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                graphics.Clear(Color.White);
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                graphics.DrawString(id == "btnScale100" ? "100" : id == "btnScaleGiam" ? "A−" : "A+",
                    font, brush, new RectangleF(0, 0, 32, 32), format);
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
