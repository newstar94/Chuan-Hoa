using System;

namespace ChuanHoa.Client.Core.Annotations
{
    public static class AnnotationSeverityResolver
    {
        public static AnnotationSeverityLevel FromValue(string severity)
        {
            var value = (severity ?? string.Empty).Trim();
            if (value.Length == 0) return AnnotationSeverityLevel.NotEvaluated;
            if (Equals(value, "Error") || Equals(value, "High") || Equals(value, "Lỗi"))
                return AnnotationSeverityLevel.Error;
            if (Equals(value, "Warning")) return AnnotationSeverityLevel.Warning;
            if (Equals(value, "Suggestion") || Equals(value, "Medium"))
                return AnnotationSeverityLevel.Suggestion;
            if (Equals(value, "Unknown")) return AnnotationSeverityLevel.Unknown;
            if (Equals(value, "NotEvaluated")) return AnnotationSeverityLevel.NotEvaluated;
            return AnnotationSeverityLevel.Unknown;
        }

        private static bool Equals(string left, string right) =>
            string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }
}
