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
            if (_disposed || !string.Equals(controlId, "btnTuDienCaNhan", StringComparison.Ordinal))
                return null!;
            object picture;
            if (_pictures.TryGetValue(controlId, out picture)) return picture;
            try
            {
                var bitmap = LoadBitmap(PersonalDictionaryResource);
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
