namespace ChuanHoa.Client.Core.Text
{
    /// <summary>
    /// Latin Eth (Ð U+00D0, ð U+00F0) renders like the Vietnamese Đ/đ
    /// (U+0110/U+0111) but is another letter. It enters documents through text
    /// decoded with a Western code page (0xD0 is Đ in Windows-1258 but Ð in
    /// Windows-1252) and silently breaks search, sorting, role detection and
    /// spelling. No Vietnamese word contains it.
    /// </summary>
    public static class VietnameseLookalikeLetters
    {
        private static readonly char[] EthLetters = { 'Ð', 'ð' };

        public static bool ContainsEth(string? value) =>
            !string.IsNullOrEmpty(value) && value!.IndexOfAny(EthLetters) >= 0;

        /// <summary>Replaces Ð with Đ and ð with đ, keeping length and case.</summary>
        public static string FoldEth(string value) =>
            (value ?? string.Empty).Replace('Ð', 'Đ').Replace('ð', 'đ');
    }
}
