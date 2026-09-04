using System;
using System.Collections.Generic;

namespace ChuanHoa.VietnameseEngine
{
    public sealed class EngineRequest
    {
        public string RequestId { get; set; } = Guid.NewGuid().ToString("N");
        public string Action { get; set; } = "check";
        public string Text { get; set; } = string.Empty;
        public int ParagraphIndex { get; set; }
        public string DocumentHash { get; set; } = string.Empty;
        public bool CheckContext { get; set; } = true;
    }

    public sealed class EngineResponse
    {
        public string RequestId { get; set; } = string.Empty;
        public string EngineVersion { get; set; } = "1.0.0";
        public bool IsSuccess { get; set; } = true;
        public string? ErrorMessage { get; set; }
        public List<EngineFinding> Findings { get; set; } = new List<EngineFinding>();
    }

    public sealed class EngineFinding
    {
        public int StartOffset { get; set; }
        public int Length { get; set; }
        public string Original { get; set; } = string.Empty;
        public int Level { get; set; } // 0: Spacing/Punct, 1: Lexicon/Typo, 2: Ambiguous, 3: Contextual
        public double Confidence { get; set; }
        public string IssueDescription { get; set; } = string.Empty;
        public List<EngineSuggestion> Suggestions { get; set; } = new List<EngineSuggestion>();
    }

    public sealed class EngineSuggestion
    {
        public string Text { get; set; } = string.Empty;
        public double Score { get; set; }
        public string Source { get; set; } = string.Empty;
    }
}
