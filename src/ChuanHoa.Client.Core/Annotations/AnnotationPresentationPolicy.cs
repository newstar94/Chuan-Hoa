namespace ChuanHoa.Client.Core.Annotations
{
    internal static class AnnotationPresentationPolicy
    {
        internal const string LatexRecommendationLabel = "[Khuyến nghị LaTeX/Typst]";

        internal static bool ShouldAnnotate(AnnotationFinding finding)
        {
            return finding.SourceFamily != AnnotationSourceFamily.NotEvaluated &&
                finding.SourceFamily != AnnotationSourceFamily.Unknown &&
                finding.SeverityLevel != AnnotationSeverityLevel.NotEvaluated &&
                finding.SeverityLevel != AnnotationSeverityLevel.Unknown;
        }

        internal static bool ShouldMarkRed(AnnotationFinding finding)
        {
            if (!ShouldAnnotate(finding)) return false;

            // LaTeX/Typst findings are optional publishing recommendations. Red is
            // reserved for mandatory ND30/HD05 findings and local language defects.
            return finding.SourceFamily != AnnotationSourceFamily.LatexTypst;
        }

        internal static string CurrentIssue(AnnotationFinding finding)
        {
            return finding.SourceFamily == AnnotationSourceFamily.LatexTypst
                ? LatexRecommendationLabel + " " + finding.CurrentIssue.Trim()
                : finding.CurrentIssue.Trim();
        }
    }
}
