using System;
using System.Text;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>Regime codes carried by <see cref="LocalScanSnapshot.RegimeCode"/>.</summary>
    public static class LocalRegimeCodes
    {
        public const string Nd30 = "ND30";
        public const string Viettel = "VIETTEL";
        public const string PartyHd05 = "PARTY_HD05";
        /// <summary>
        /// Văn bản quy phạm pháp luật: thể thức theo Phụ lục I Nghị định 78/2025/NĐ-CP,
        /// được thay thế bởi Phụ lục I Nghị định 187/2025/NĐ-CP (Điều 3 khoản 3 điểm a).
        /// </summary>
        public const string Vbqppl = "VBQPPL";
    }

    /// <summary>Which Phụ lục III (NĐ 187/2025/NĐ-CP) form a VBQPPL document follows.</summary>
    public enum VbqpplDocumentKind
    {
        /// <summary>Not identified as one of the forms implemented here.</summary>
        Other = 0,
        /// <summary>Nghị quyết của Hội đồng nhân dân các cấp — Mẫu số 17, 18.</summary>
        HdndResolution = 1,
        /// <summary>Quyết định của Ủy ban nhân dân các cấp — Mẫu số 19, 20.</summary>
        UbndDecision = 2
    }

    /// <summary>
    /// Facts quoted from Phụ lục I and Phụ lục III of Nghị định 187/2025/NĐ-CP
    /// (shared/docs/extracted_text/PL1_ND187.txt, PL3_ND187.txt). Only values that can be
    /// quoted from that text are defined here.
    /// </summary>
    public static class VbqpplRegimePolicy
    {
        // PL I, Mục 1, Phần I, khoản 3: "cách mép trái 30 - 35 mi-li-mét (mm)" and
        // "mép phải 15 - 20 mi-li-mét (mm)". The sơ đồ in Phần IV agrees on both.
        // Top and bottom margins are deliberately absent: khoản 3 says 15 - 20 mm while
        // the sơ đồ in Phần IV says 20-25 mm, so neither value is enforced.
        public const double LeftMarginMinMm = 30d;
        public const double LeftMarginMaxMm = 35d;
        public const double RightMarginMinMm = 15d;
        public const double RightMarginMaxMm = 20d;

        // PL I, Mục 1, Phần II, khoản 4 điểm b: "Tên loại văn bản ... cỡ chữ 14".
        public const double TypeNameFontPoints = 14d;
        // PL I, Mục 1, Phần II, khoản 7 điểm b: chức vụ người ký "cỡ chữ 13".
        public const double SignerAuthorityFontPoints = 13d;
        // PL I, Mục 1, Phần III, khoản 1: căn cứ "cỡ chữ 14, kiểu chữ nghiêng".
        public const double LegalBasisFontPoints = 14d;

        /// <summary>
        /// "Số: 05/2025/QĐ-UBND": PL I, Mục 1, Phần II, khoản 3 điểm d and đ — số thứ tự,
        /// năm ban hành (ghi đầy đủ các số), ký hiệu; Mẫu số 17-20 "Số:.../20...(2)/...".
        /// </summary>
        private static readonly Regex YearNotation = new Regex(
            @"^\s*Số\s*(?<colon>:?)\s*(?<number>\d+)\s*/\s*(?<year>\d{4})\s*/\s*(?<notation>[^\r\n\v]*?)\s*$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));

        private static readonly Regex HdndNotation = new Regex(@"^NQ\s*-\s*HĐND$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));
        private static readonly Regex UbndNotation = new Regex(@"^QĐ\s*-\s*UBND$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));

        private static readonly Regex HdndOrgan = new Regex(@"^HỘI\s+ĐỒNG\s+NHÂN\s+DÂN\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));
        private static readonly Regex UbndOrgan = new Regex(@"^(?:ỦY|UỶ)\s+BAN\s+NHÂN\s+DÂN\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));

        // PL I, Mục 1, Phần III, khoản 1 điểm đ; Mẫu 17-20: "Hội đồng nhân dân ban hành
        // Nghị quyết ..." / "Ủy ban nhân dân ban hành Quyết định ...".
        private static readonly Regex EnactingFormula = new Regex(
            @"^\s*(?:Hội\s+đồng\s+nhân\s+dân|(?:Ủy|Uỷ)\s+ban\s+nhân\s+dân)\b[^\r\n]*?\bban\s+hành\s+(?:Nghị\s+quyết|Quyết\s+định)\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));

        // The ND30-style operative line that Mẫu 17-20 do not contain.
        private static readonly Regex OperativeHeading = new Regex(@"^\s*(?:QUYẾT\s+NGHỊ|QUYẾT\s+ĐỊNH)\s*:\s*$",
            RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200));

        // PL I, Mục 1, Phần II, khoản 6 điểm c; Mẫu 17, 18: "Nghị quyết này đã được Hội đồng
        // nhân dân...(1) khóa..., Kỳ họp thứ... thông qua ngày...tháng...năm...".
        private static readonly Regex PassageLine = new Regex(
            @"^\s*Nghị\s+quyết\s+này\s+đã\s+được\s+Hội\s+đồng\s+nhân\s+dân\b[^\r\n]*\bthông\s+qua\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));
        private static readonly Regex PassageTerm = new Regex(@"\bkh(?:óa|oá)\s+(?<term>[^\s,;.]+)",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));
        private static readonly Regex PassageSession = new Regex(@"\bkỳ\s+họp\s+thứ\s+(?<session>[^\s,;.]+)",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));
        private static readonly Regex RomanNumeral = new Regex(@"^[IVXLCDM]+$",
            RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200));

        public static bool IsVbqppl(LocalScanSnapshot snapshot) =>
            snapshot != null && IsVbqppl(snapshot.RegimeCode);

        public static bool IsVbqppl(string? regimeCode) =>
            string.Equals(regimeCode, LocalRegimeCodes.Vbqppl, StringComparison.OrdinalIgnoreCase);

        public static string Normalize(string? value) =>
            Regex.Replace((value ?? string.Empty).Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();

        /// <summary>
        /// Parses "Số: NN/YYYY/KÝ-HIỆU". Returns false when the text has no four-digit year
        /// between two slashes (the ND30 administrative form "Số: NN/KÝ-HIỆU").
        /// </summary>
        public static bool TryParseYearNotation(string? codeNumberText, out Match match)
        {
            match = YearNotation.Match(Normalize(codeNumberText));
            return match.Success;
        }

        /// <summary>
        /// Kind from a header "Số, ký hiệu" that carries the year: only NQ-HĐND (Mẫu 17, 18)
        /// and QĐ-UBND (Mẫu 19, 20) are recognized. ND30 administrative resolutions and
        /// decisions of the same bodies are written without the year ("Số: 12/NQ-HĐND")
        /// and therefore return <see cref="VbqpplDocumentKind.Other"/>.
        /// </summary>
        public static VbqpplDocumentKind KindFromCodeNumber(string? codeNumberText)
        {
            Match match;
            if (!TryParseYearNotation(codeNumberText, out match)) return VbqpplDocumentKind.Other;
            var notation = match.Groups["notation"].Value.Trim();
            if (HdndNotation.IsMatch(notation)) return VbqpplDocumentKind.HdndResolution;
            if (UbndNotation.IsMatch(notation)) return VbqpplDocumentKind.UbndDecision;
            return VbqpplDocumentKind.Other;
        }

        /// <summary>Kind from the type name and issuing-organ heading when the notation is unclear.</summary>
        public static VbqpplDocumentKind KindFromHeader(string? typeNameText, string? organNameText)
        {
            var type = Normalize(typeNameText).ToUpperInvariant();
            var organ = Normalize(organNameText);
            if (type == "NGHỊ QUYẾT" && HdndOrgan.IsMatch(organ)) return VbqpplDocumentKind.HdndResolution;
            if (type == "QUYẾT ĐỊNH" && UbndOrgan.IsMatch(organ)) return VbqpplDocumentKind.UbndDecision;
            return VbqpplDocumentKind.Other;
        }

        public static bool IsEnactingFormula(string? text) => EnactingFormula.IsMatch(Normalize(text));

        public static bool IsOperativeHeading(string? text) => OperativeHeading.IsMatch(Normalize(text));

        public static bool IsPassageLine(string? text) => PassageLine.IsMatch(Normalize(text));

        /// <summary>
        /// PL I, Mục 1, Phần II, khoản 6 điểm c: "Số chỉ khóa ... Hội đồng nhân dân dùng số
        /// La Mã; số chỉ kỳ họp ... dùng số Ả Rập, trừ kỳ họp thứ nhất." Returns the problems
        /// found in a passage line (empty when none).
        /// </summary>
        public static string[] PassageNumberingProblems(string? passageText)
        {
            var text = Normalize(passageText);
            var problems = new System.Collections.Generic.List<string>();
            var term = PassageTerm.Match(text);
            if (term.Success && !RomanNumeral.IsMatch(term.Groups["term"].Value))
                problems.Add("số khóa chưa dùng số La Mã");
            var session = PassageSession.Match(text);
            if (session.Success)
            {
                var value = session.Groups["session"].Value;
                var isFirst = string.Equals(value, "nhất", StringComparison.OrdinalIgnoreCase);
                var isArabic = value.Length > 0 && Regex.IsMatch(value, @"^\d+$");
                if (!isFirst && !isArabic) problems.Add("số kỳ họp chưa dùng số Ả Rập");
            }
            return problems.ToArray();
        }
    }
}
