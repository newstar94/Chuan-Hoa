using System;

namespace ChuanHoa.Client.Core.Annotations
{
    public static class AnnotationSourceFamilyResolver
    {
        public static AnnotationSourceFamily FromRuleCode(string ruleCode)
        {
            if (string.IsNullOrWhiteSpace(ruleCode))
                return AnnotationSourceFamily.NotEvaluated;

            if (HasPrefix(ruleCode, "ND30-")) return AnnotationSourceFamily.Nd30;
            if (HasPrefix(ruleCode, "HD05-")) return AnnotationSourceFamily.Hd05;
            if (HasPrefix(ruleCode, "LATEX-")) return AnnotationSourceFamily.LatexTypst;
            if (HasPrefix(ruleCode, "LOCAL-") || HasPrefix(ruleCode, "SPELLING-"))
                return AnnotationSourceFamily.LocalLanguage;

            // Legacy administrative format rule codes predate the canonical ND30-
            // prefix but still belong to the mandatory administrative format lane.
            if (HasPrefix(ruleCode, "FORMAT-")) return AnnotationSourceFamily.Nd30;

            return AnnotationSourceFamily.Unknown;
        }

        private static bool HasPrefix(string value, string prefix) =>
            value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
    }
}
