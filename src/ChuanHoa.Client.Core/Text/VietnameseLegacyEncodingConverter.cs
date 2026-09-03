using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Text
{
    public sealed class LegacyEncodingConversionResult
    {
        public LegacyEncodingConversionResult(string text, bool converted, int unmappedCharacters)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            Converted = converted;
            UnmappedCharacters = unmappedCharacters;
        }

        public string Text { get; }
        public bool Converted { get; }
        public int UnmappedCharacters { get; }
    }

    public static class VietnameseLegacyEncodingConverter
    {
        private static readonly Regex TcvnUpperFont = new Regex(
            @"^\.Vn.*H$", RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex TcvnLowerFont = new Regex(
            @"^\.Vn(?!.*H$).*$", RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex VniFont = new Regex(
            @"^VNI-", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly IReadOnlyDictionary<char, string> TcvnLower = BuildDirectMap(
            "A1=Ă;A2=Â;A3=Ê;A4=Ô;A5=Ơ;A6=Ư;A7=Đ;A8=ă;A9=â;AA=ê;AB=ô;AC=ơ;AD=ư;AE=đ;AF=Ằ;" +
            "B5=à;B6=ả;B7=ã;B8=á;B9=ạ;BA=Ẳ;BB=ằ;BC=ẳ;BD=ẵ;BE=ắ;BF=Ẵ;C0=Ắ;C1=Ầ;C2=Ẩ;C3=Ẫ;C4=Ấ;" +
            "C5=Ề;C6=ặ;C7=ầ;C8=ẩ;C9=ẫ;CA=ấ;CB=ậ;CC=è;CD=Ể;CE=ẻ;CF=ẽ;D0=é;D1=ẹ;D2=ề;D3=ể;D4=ễ;D5=ế;" +
            "D6=ệ;D7=ì;D8=ỉ;D9=Ễ;DA=Ế;DB=Ồ;DC=ĩ;DD=í;DE=ị;DF=ò;E0=Ổ;E1=ỏ;E2=õ;E3=ó;E4=ọ;E5=ồ;E6=ổ;" +
            "E7=ỗ;E8=ố;E9=ộ;EA=ờ;EB=ở;EC=ỡ;ED=ớ;EE=ợ;EF=ù;F0=Ỗ;F1=ủ;F2=ũ;F3=ú;F4=ụ;F5=ừ;F6=ử;F7=ữ;" +
            "F8=ứ;F9=ự;FA=ỳ;FB=ỷ;FC=ỹ;FD=ý;FE=ỵ;FF=Ố;" +
            "80=À;81=Ả;82=Ã;83=Á;84=Ạ;85=Ặ;86=Ậ;87=È;88=Ẻ;89=Ẽ;8A=É;8B=Ẹ;8C=Ệ;8D=Ì;8E=Ỉ;8F=Ĩ;" +
            "90=Í;91=Ị;92=Ò;93=Ỏ;94=Õ;95=Ó;96=Ọ;97=Ộ;98=Ờ;99=Ở;9A=Ỡ;9B=Ớ;9C=Ợ;9D=Ù;9E=Ủ;9F=Ũ");

        private static readonly IReadOnlyDictionary<char, string> TcvnUpper = BuildDirectMap(
            "A1=Ă;A2=Â;A3=Ê;A4=Ô;A5=Ơ;A6=Ư;A7=Đ;A8=Ă;A9=Â;AA=Ê;AB=Ô;AC=Ơ;AD=Ư;AE=Đ;AF=Ằ;" +
            "B5=À;B6=Ả;B7=Ã;B8=Á;B9=Ạ;BA=Ẳ;BB=Ằ;BC=Ẳ;BD=Ẵ;BE=Ắ;BF=Ẵ;C0=Ắ;C1=Ầ;C2=Ẩ;C3=Ẫ;C4=Ấ;" +
            "C5=Ề;C6=Ặ;C7=Ầ;C8=Ẩ;C9=Ẫ;CA=Ấ;CB=Ậ;CC=È;CD=Ể;CE=Ẻ;CF=Ẽ;D0=É;D1=Ẹ;D2=Ề;D3=Ể;D4=Ễ;D5=Ế;" +
            "D6=Ệ;D7=Ì;D8=Ỉ;D9=Ễ;DA=Ế;DB=Ồ;DC=Ĩ;DD=Í;DE=Ị;DF=Ò;E0=Ổ;E1=Ỏ;E2=Õ;E3=Ó;E4=Ọ;E5=Ồ;E6=Ổ;" +
            "E7=Ỗ;E8=Ố;E9=Ộ;EA=Ờ;EB=Ở;EC=Ỡ;ED=Ớ;EE=Ợ;EF=Ù;F0=Ỗ;F1=Ủ;F2=Ũ;F3=Ú;F4=Ụ;F5=Ừ;F6=Ử;F7=Ữ;" +
            "F8=Ứ;F9=Ự;FA=Ỳ;FB=Ỷ;FC=Ỹ;FD=Ý;FE=Ỵ;FF=Ố;" +
            "80=À;81=Ả;82=Ã;83=Á;84=Ạ;85=Ặ;86=Ậ;87=È;88=Ẻ;89=Ẽ;8A=É;8B=Ẹ;8C=Ệ;8D=Ì;8E=Ỉ;8F=Ĩ;" +
            "90=Í;91=Ị;92=Ò;93=Ỏ;94=Õ;95=Ó;96=Ọ;97=Ộ;98=Ờ;99=Ở;9A=Ỡ;9B=Ớ;9C=Ợ;9D=Ù;9E=Ủ;9F=Ũ");

        private static readonly IReadOnlyDictionary<char, string> VniDirect = BuildDirectMap(
            "C6=Ỉ;CC=Ì;CD=Í;CE=Ỵ;D1=Đ;D2=Ị;D3=Ĩ;D4=Ơ;D6=Ư;E6=ỉ;EC=ì;ED=í;EE=ỵ;F1=đ;F2=ị;F3=ĩ;F4=ơ;F6=ư");

        private static readonly IReadOnlyDictionary<char, string> VniCombining = BuildCombiningMap();

        public static bool IsLegacyFont(string? fontName)
        {
            var value = (fontName ?? string.Empty).Trim();
            return TcvnUpperFont.IsMatch(value) || TcvnLowerFont.IsMatch(value) || VniFont.IsMatch(value);
        }

        public static LegacyEncodingConversionResult Convert(string? fontName, string text)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            var value = (fontName ?? string.Empty).Trim();
            if (TcvnUpperFont.IsMatch(value)) return Apply(text, TcvnUpper, null);
            if (TcvnLowerFont.IsMatch(value)) return Apply(text, TcvnLower, null);
            if (VniFont.IsMatch(value)) return Apply(text, VniDirect, VniCombining);
            return new LegacyEncodingConversionResult(text, false, 0);
        }

        private static LegacyEncodingConversionResult Apply(
            string text,
            IReadOnlyDictionary<char, string> direct,
            IReadOnlyDictionary<char, string>? combining)
        {
            var pieces = new List<string>(text.Length);
            var unmapped = 0;
            foreach (var character in text)
            {
                string mapped;
                if (direct.TryGetValue(character, out mapped!))
                {
                    pieces.Add(mapped);
                }
                else if (combining != null && combining.TryGetValue(character, out mapped!))
                {
                    if (pieces.Count == 0)
                    {
                        pieces.Add(character.ToString());
                        unmapped++;
                    }
                    else
                    {
                        pieces[pieces.Count - 1] = (pieces[pieces.Count - 1] + mapped)
                            .Normalize(NormalizationForm.FormC);
                    }
                }
                else
                {
                    pieces.Add(character.ToString());
                    if (character >= 0x80 && character <= 0xFF) unmapped++;
                }
            }

            var output = string.Concat(pieces).Normalize(NormalizationForm.FormC);
            return new LegacyEncodingConversionResult(output, true, unmapped);
        }

        private static IReadOnlyDictionary<char, string> BuildDirectMap(string specification)
        {
            var result = new Dictionary<char, string>();
            foreach (var pair in specification.Split(';'))
            {
                var separator = pair.IndexOf('=');
                var code = int.Parse(pair.Substring(0, separator), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                result.Add((char)code, pair.Substring(separator + 1));
            }
            return result;
        }

        private static IReadOnlyDictionary<char, string> BuildCombiningMap()
        {
            return new Dictionary<char, string>
            {
                [(char)0xC0] = "\u0302\u0300", [(char)0xC1] = "\u0302\u0301", [(char)0xC2] = "\u0302",
                [(char)0xC3] = "\u0302\u0303", [(char)0xC4] = "\u0323\u0302", [(char)0xC5] = "\u0302\u0309",
                [(char)0xC8] = "\u0306\u0300", [(char)0xC9] = "\u0306\u0301", [(char)0xCA] = "\u0306",
                [(char)0xCB] = "\u0323\u0306", [(char)0xCF] = "\u0323", [(char)0xD5] = "\u0303",
                [(char)0xD8] = "\u0300", [(char)0xD9] = "\u0301", [(char)0xDA] = "\u0306\u0309",
                [(char)0xDB] = "\u0309", [(char)0xDC] = "\u0306\u0303", [(char)0xE0] = "\u0302\u0300",
                [(char)0xE1] = "\u0302\u0301", [(char)0xE2] = "\u0302", [(char)0xE3] = "\u0302\u0303",
                [(char)0xE4] = "\u0323\u0302", [(char)0xE5] = "\u0302\u0309", [(char)0xE8] = "\u0306\u0300",
                [(char)0xE9] = "\u0306\u0301", [(char)0xEA] = "\u0306", [(char)0xEB] = "\u0323\u0306",
                [(char)0xEF] = "\u0323", [(char)0xF5] = "\u0303", [(char)0xF8] = "\u0300",
                [(char)0xF9] = "\u0301", [(char)0xFA] = "\u0306\u0309", [(char)0xFB] = "\u0309",
                [(char)0xFC] = "\u0306\u0303"
            };
        }
    }
}
