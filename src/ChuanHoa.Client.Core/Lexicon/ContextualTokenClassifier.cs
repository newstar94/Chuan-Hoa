using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Lexicon
{
    /// <summary>Why a token that is not a Vietnamese word is still correct where it stands.</summary>
    public enum ContextualTokenKind
    {
        None,
        MeasurementUnit,
        Multiplication,
        StructuralLabel,
        ListMarker,
        Abbreviation,
        CodeOrFormula
    }

    /// <summary>
    /// Recognises tokens that a dictionary check must not report because their
    /// context, not the lexicon, makes them correct: unit symbols after a number,
    /// in a compound unit or in a unit column ("25 kg", "CFU/g", "m2"); the "x" of
    /// a product ("3,12 x 10", "dài x rộng"); the letter of a structural reference
    /// ("điểm b, c khoản 1", "phụ lục d"); inline list markers ("a)", "(ii)"); slash
    /// and dotted abbreviations ("V/v", "đ/c", "v.v.", "tr. 15", "Sđt:"); and
    /// letters glued to digits or beside operators in codes and formulas ("A4",
    /// "P = m x g"). A token outside such a context is still checked, so a stray
    /// "b" or "x" in running text is reported as before.
    /// </summary>
    public static class ContextualTokenClassifier
    {
        private static readonly CultureInfo Vietnamese = CultureInfo.GetCultureInfo("vi-VN");

        // Unit symbols that are never Vietnamese syllables: correct anywhere.
        private static readonly HashSet<string> UnambiguousUnits = new(StringComparer.OrdinalIgnoreCase)
        {
            "mg", "kg", "mcg", "µg", "nm", "µm", "mm", "cm", "dm", "km", "md", "ml", "dl",
            "kw", "kwh", "mw", "mwh", "gw", "gwh", "kva", "kvar", "kv", "hz", "khz", "mhz", "ghz",
            "kb", "mb", "gb", "kbps", "mbps", "gbps", "ppm", "ppb", "cfu", "mol", "mmol",
            "kcal", "kj", "kpa", "mpa", "hpa", "rpm", "dpi", "vnđ", "vnd", "usd", "eur"
        };

        // Unit symbols that are also letters or syllables: correct only where a
        // unit belongs (after a number, in a compound unit, a unit column or label).
        private static readonly HashSet<string> ContextUnits = new(StringComparer.OrdinalIgnoreCase)
        {
            "g", "m", "l", "t", "h", "s", "w", "v", "n", "j", "k", "đ", "ha", "lb", "oz", "ft",
            "cal", "pa", "bar", "lux", "min"
        };

        private static readonly HashSet<string> NumberWords = new(StringComparer.OrdinalIgnoreCase)
        {
            "một", "hai", "ba", "bốn", "tư", "năm", "lăm", "sáu", "bảy", "bẩy", "tám", "chín",
            "mười", "mươi", "trăm", "nghìn", "ngàn", "triệu", "tỷ", "tỉ", "chục"
        };

        // Nouns that name a lettered part: "điểm b", "phụ lục c", "phương án a".
        private static readonly HashSet<string> LabelNouns = new(StringComparer.OrdinalIgnoreCase)
        {
            "điểm", "tiết", "mục", "phần", "mẫu", "biểu", "bảng", "hình", "loại", "nhóm", "hạng",
            "khu", "lô", "dãy", "khối", "kiểu", "dạng", "cột", "ô", "câu", "ý", "vùng", "tòa", "toà",
            "phụ lục", "phương án", "trường hợp", "sơ đồ"
        };

        // Words joining lettered parts: "điểm a, b và c", "điểm a đến đ".
        private static readonly HashSet<string> LabelConnectors = new(StringComparer.OrdinalIgnoreCase)
        {
            ",", "và", "hoặc", "đến", "tới", "-", "–", "—", "&"
        };

        // Lower-case administrative abbreviations written with "." or ":".
        private static readonly HashSet<string> KnownAbbreviations = new(StringComparer.OrdinalIgnoreCase)
        {
            "sđt", "đt", "stk", "tk", "mst", "đc", "đk", "vd", "tr", "tx", "tt", "qđ", "hđ", "cv", "tb",
            "ths", "ts", "pgs", "gs", "bs", "ks", "cn", "nxb"
        };

        private static readonly Regex LowerRomanNumeral = new(
            @"^(?=[ivx])x{0,3}(?:ix|iv|v?i{0,3})$", RegexOptions.CultureInvariant);

        private const string SuperscriptDigits = "⁰¹²³⁴⁵⁶⁷⁸⁹";
        private const string Operators = "=+*×÷^<>≤≥±≈≠";

        public static ContextualTokenKind Classify(IReadOnlyList<VietnameseTokenSpan> tokens, int index,
            bool inUnitColumn = false)
        {
            if (tokens == null || index < 0 || index >= tokens.Count) return ContextualTokenKind.None;
            var value = tokens[index].Text.Normalize(NormalizationForm.FormC);
            var lower = value.ToLower(Vietnamese);
            var letters = CountLetters(value);
            if (letters == 0) return ContextualTokenKind.None;

            if (UnambiguousUnits.Contains(lower) || string.Equals(value, "pH", StringComparison.Ordinal))
                return ContextualTokenKind.MeasurementUnit;
            if (inUnitColumn && letters <= 4) return ContextualTokenKind.MeasurementUnit;
            if (ContextUnits.Contains(lower) && IsUnitPosition(tokens, index))
                return ContextualTokenKind.MeasurementUnit;
            if (IsMultiplication(tokens, index, lower)) return ContextualTokenKind.Multiplication;
            if (letters <= 4 && IsGluedToDigits(tokens, index)) return ContextualTokenKind.CodeOrFormula;
            if (letters <= 2 && IsInFormula(tokens, index)) return ContextualTokenKind.CodeOrFormula;
            if (letters == 1 && IsStructuralLabel(tokens, index, 0)) return ContextualTokenKind.StructuralLabel;
            if (IsListMarker(tokens, index, lower, letters)) return ContextualTokenKind.ListMarker;
            if (IsAbbreviation(tokens, index, lower, letters)) return ContextualTokenKind.Abbreviation;
            return ContextualTokenKind.None;
        }

        /// <summary>
        /// True when an unknown short token stands where only a unit or code can:
        /// in a unit column, behind a unit label ("ĐVT: md") or glued to a number
        /// ("120md"). The document then uses it as its own unit, so the same token
        /// is accepted in its other paragraphs as well.
        /// </summary>
        public static bool IsDocumentUnitEvidence(IReadOnlyList<VietnameseTokenSpan> tokens, int index,
            bool inUnitColumn)
        {
            if (tokens == null || index < 0 || index >= tokens.Count) return false;
            var token = tokens[index];
            if (token.Kind != VietnameseTokenKind.Word) return false;
            var letters = CountLetters(token.Text);
            if (letters < 2 || letters > 4) return false;
            return inUnitColumn || FollowsUnitLabel(tokens, index) ||
                (index > 0 && IsNumeric(tokens[index - 1]));
        }

        private static bool IsUnitPosition(IReadOnlyList<VietnameseTokenSpan> tokens, int index)
        {
            var previous = Significant(tokens, index, -1);
            if (previous >= 0 && (IsNumeric(tokens[previous]) ||
                NumberWords.Contains(tokens[previous].Text.Normalize(NormalizationForm.FormC))))
                return true;
            // A compound unit: "CFU/g", "đ/kg", "km/h", "m/s".
            if (Is(tokens, index - 1, "/") && IsUnitOperand(tokens, index - 2)) return true;
            if (Is(tokens, index + 1, "/") && IsUnitOperand(tokens, index + 2)) return true;
            // An exponent: "m2", "m³".
            if (index + 1 < tokens.Count && IsNumeric(tokens[index + 1])) return true;
            // A unit heading: "Diện tích (m)", "Thành tiền (đ)".
            if (Is(tokens, index - 1, "(") && Is(tokens, index + 1, ")")) return true;
            return FollowsUnitLabel(tokens, index);
        }

        private static bool IsUnitOperand(IReadOnlyList<VietnameseTokenSpan> tokens, int index)
        {
            if (index < 0 || index >= tokens.Count) return false;
            var token = tokens[index];
            return token.Kind == VietnameseTokenKind.Number || token.Kind == VietnameseTokenKind.AcronymOrUpper ||
                UnambiguousUnits.Contains(token.Text) || ContextUnits.Contains(token.Text);
        }

        private static bool FollowsUnitLabel(IReadOnlyList<VietnameseTokenSpan> tokens, int index)
        {
            var colon = Significant(tokens, index, -1);
            if (colon < 0 || tokens[colon].Text != ":") return false;
            var label = Significant(tokens, colon, -1);
            if (label < 0) return false;
            var word = tokens[label].Text.Normalize(NormalizationForm.FormC).ToLower(Vietnamese);
            if (word == "đvt" || word == "dvt") return true;
            if (word == "tính") label = Significant(tokens, label, -1);
            else if (word != "vị") return false;
            if (label < 0) return false;
            var unitWord = tokens[label].Text.Normalize(NormalizationForm.FormC).ToLower(Vietnamese);
            if (unitWord != "vị") return false;
            var unitHead = Significant(tokens, label, -1);
            return unitHead >= 0 &&
                tokens[unitHead].Text.Normalize(NormalizationForm.FormC).ToLower(Vietnamese) == "đơn";
        }

        private static bool IsMultiplication(IReadOnlyList<VietnameseTokenSpan> tokens, int index, string lower)
        {
            if (lower != "x") return false;
            var previous = Significant(tokens, index, -1);
            var next = Significant(tokens, index, 1);
            if (previous < 0 || next < 0) return false;
            if (IsFactor(tokens[previous], true) && IsFactor(tokens[next], false)) return true;
            // "Kích thước (dài x rộng x cao): 20 x 30 x 40 cm": once the paragraph
            // multiplies numbers, an "x" between the dimension names is the same sign.
            if (tokens[previous].Kind != VietnameseTokenKind.Word || tokens[next].Kind != VietnameseTokenKind.Word)
                return false;
            for (var other = 0; other < tokens.Count; other++)
            {
                if (other == index || !string.Equals(tokens[other].Text, "x", StringComparison.OrdinalIgnoreCase))
                    continue;
                var before = Significant(tokens, other, -1);
                var after = Significant(tokens, other, 1);
                if (before >= 0 && after >= 0 && IsFactor(tokens[before], true) && IsFactor(tokens[after], false))
                    return true;
            }
            return false;
        }

        private static bool IsFactor(VietnameseTokenSpan token, bool left) =>
            IsNumeric(token) || token.Text == (left ? ")" : "(");

        private static bool IsGluedToDigits(IReadOnlyList<VietnameseTokenSpan> tokens, int index) =>
            (index > 0 && IsNumeric(tokens[index - 1])) ||
            (index + 1 < tokens.Count && IsNumeric(tokens[index + 1]));

        /// <summary>
        /// True when the token belongs to an expression with an operator: walking
        /// out over variables, numbers and brackets ("P = m x g") reaches "=", "+", ...
        /// </summary>
        private static bool IsInFormula(IReadOnlyList<VietnameseTokenSpan> tokens, int index)
        {
            foreach (var direction in new[] { -1, 1 })
            {
                var position = index;
                while ((position = Significant(tokens, position, direction)) >= 0)
                {
                    var token = tokens[position];
                    var text = token.Text;
                    // A leading "+" or "*" is a bullet ("+ Hàm lượng ..."), not an operator.
                    if (text.Length == 1 && Operators.IndexOf(text[0]) >= 0 &&
                        Significant(tokens, position, -1) >= 0)
                        return true;
                    var term = IsNumeric(token) || text == "(" || text == ")" ||
                        ((token.Kind == VietnameseTokenKind.Word || token.Kind == VietnameseTokenKind.AcronymOrUpper) &&
                         CountLetters(text) <= 2);
                    if (!term) break;
                }
            }
            return false;
        }

        private static bool IsStructuralLabel(IReadOnlyList<VietnameseTokenSpan> tokens, int index, int depth)
        {
            if (depth > 8) return false;
            var previous = Significant(tokens, index, -1);
            if (previous < 0) return false;
            var word = tokens[previous].Text.Normalize(NormalizationForm.FormC);
            if (LabelNouns.Contains(word)) return true;
            var beforeWord = Significant(tokens, previous, -1);
            if (beforeWord >= 0 && LabelNouns.Contains(
                    tokens[beforeWord].Text.Normalize(NormalizationForm.FormC) + " " + word))
                return true;
            if (!LabelConnectors.Contains(word) || beforeWord < 0) return false;
            return CountLetters(tokens[beforeWord].Text) == 1 && IsStructuralLabel(tokens, beforeWord, depth + 1);
        }

        private static bool IsListMarker(IReadOnlyList<VietnameseTokenSpan> tokens, int index, string lower,
            int letters)
        {
            if (letters != 1 && !LowerRomanNumeral.IsMatch(lower)) return false;
            if (Is(tokens, index - 1, "(") && Is(tokens, index + 1, ")")) return true;
            var previous = Significant(tokens, index, -1);
            // At the start of a paragraph: "d.Bước 4", "b . Không ..." (the stray space
            // is a punctuation finding), a cell holding only "B" or "ii", or an
            // upper-case group label "B Giằng thép".
            if (previous < 0)
            {
                var following = Significant(tokens, index, 1);
                if (following < 0 || tokens[following].Text == "." || tokens[following].Text == ")")
                    return true;
                if (letters == 1 && char.IsUpper(tokens[index].Text[0]) &&
                    tokens[index + 1].Kind == VietnameseTokenKind.Whitespace)
                    return true;
            }
            var opensItem = previous < 0 || IsAny(tokens[previous].Text, ":", ";", ",", ".");
            // "Hồ sơ gồm: a) Đơn đề nghị; b) Bản sao ..."
            if (Is(tokens, index + 1, ")") && opensItem && (index == 0 || tokens[index - 1].Kind ==
                    VietnameseTokenKind.Whitespace))
                return true;
            // "... như sau: a. Nội dung; b. Thời hạn."
            if (letters == 1 && Is(tokens, index + 1, ".") && opensItem && previous >= 0)
            {
                var next = Significant(tokens, index + 1, 1);
                return next >= 0 && tokens[next].Kind != VietnameseTokenKind.Number &&
                    tokens[next].Text.Length > 0 && char.IsUpper(tokens[next].Text[0]);
            }
            return false;
        }

        private static bool IsAbbreviation(IReadOnlyList<VietnameseTokenSpan> tokens, int index, string lower,
            int letters)
        {
            // Slash abbreviations: "V/v", "đ/c", "k/g", "T/M".
            if (letters <= 2 &&
                ((Is(tokens, index - 1, "/") && IsShortLetters(tokens, index - 2)) ||
                 (Is(tokens, index + 1, "/") && IsShortLetters(tokens, index + 2))))
                return true;
            // Dotted abbreviations: "v.v.", "đ.c.".
            if (letters == 1 &&
                ((Is(tokens, index + 1, ".") && IsSingleLetter(tokens, index + 2)) ||
                 (Is(tokens, index - 1, ".") && IsSingleLetter(tokens, index - 2))))
                return true;
            // A short abbreviation before a number: "tr. 15", "q.3", "p. 7".
            if (letters <= 3 && Is(tokens, index + 1, "."))
            {
                var next = Significant(tokens, index + 1, 1);
                if (next >= 0 && tokens[next].Kind == VietnameseTokenKind.Number) return true;
            }
            return KnownAbbreviations.Contains(lower) &&
                (Is(tokens, index + 1, ".") || Is(tokens, index + 1, ":"));
        }

        private static bool IsShortLetters(IReadOnlyList<VietnameseTokenSpan> tokens, int index) =>
            index >= 0 && index < tokens.Count &&
            (tokens[index].Kind == VietnameseTokenKind.Word || tokens[index].Kind == VietnameseTokenKind.AcronymOrUpper) &&
            CountLetters(tokens[index].Text) <= 3;

        private static bool IsSingleLetter(IReadOnlyList<VietnameseTokenSpan> tokens, int index) =>
            index >= 0 && index < tokens.Count && tokens[index].Kind == VietnameseTokenKind.Word &&
            CountLetters(tokens[index].Text) == 1;

        private static bool IsNumeric(VietnameseTokenSpan token)
        {
            if (token.Kind == VietnameseTokenKind.Number) return true;
            if (token.Kind != VietnameseTokenKind.Other || token.Text.Length == 0) return false;
            foreach (var character in token.Text)
                if (SuperscriptDigits.IndexOf(character) < 0) return false;
            return true;
        }

        private static bool Is(IReadOnlyList<VietnameseTokenSpan> tokens, int index, string text) =>
            index >= 0 && index < tokens.Count && tokens[index].Text == text;

        private static bool IsAny(string text, params string[] values) => Array.IndexOf(values, text) >= 0;

        /// <summary>Index of the nearest non-whitespace token in a direction, or -1.</summary>
        private static int Significant(IReadOnlyList<VietnameseTokenSpan> tokens, int index, int direction)
        {
            for (var position = index + direction; position >= 0 && position < tokens.Count; position += direction)
                if (tokens[position].Kind != VietnameseTokenKind.Whitespace) return position;
            return -1;
        }

        private static int CountLetters(string value)
        {
            var count = 0;
            foreach (var character in value)
                if (char.IsLetter(character)) count++;
            return count;
        }
    }
}
