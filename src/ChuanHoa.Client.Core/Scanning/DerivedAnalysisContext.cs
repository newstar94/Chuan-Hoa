using System;
using System.Collections.Generic;
using System.Linq;

namespace ChuanHoa.Client.Core.Scanning
{
    /// <summary>
    /// Pure-.NET conclusions derived from an immutable raw document snapshot and a
    /// detector policy. None of these values participate in the Word-state fingerprint.
    /// </summary>
    public sealed class DerivedAnalysisContext
    {
        private DerivedAnalysisContext(
            IReadOnlyList<LogicalDocumentBlock> logicalBlocks,
            IReadOnlyDictionary<int, string> rolesByParagraphIndex,
            IReadOnlyDictionary<int, string> logicalBlockIdsByParagraphIndex,
            IReadOnlyList<DetectedHeading> headings)
        {
            LogicalBlocks = logicalBlocks;
            RolesByParagraphIndex = rolesByParagraphIndex;
            LogicalBlockIdsByParagraphIndex = logicalBlockIdsByParagraphIndex;
            Headings = headings;
        }

        public IReadOnlyList<LogicalDocumentBlock> LogicalBlocks { get; }
        public IReadOnlyDictionary<int, string> RolesByParagraphIndex { get; }
        public IReadOnlyDictionary<int, string> LogicalBlockIdsByParagraphIndex { get; }
        public IReadOnlyList<DetectedHeading> Headings { get; }

        public static DerivedAnalysisContext Create(
            LocalScanSnapshot snapshot,
            IReadOnlyList<LogicalDocumentBlock> logicalBlocks,
            IReadOnlyDictionary<int, string> rolesByParagraphIndex,
            HeadingDetector headingDetector)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (logicalBlocks == null) throw new ArgumentNullException(nameof(logicalBlocks));
            if (rolesByParagraphIndex == null) throw new ArgumentNullException(nameof(rolesByParagraphIndex));
            if (headingDetector == null) throw new ArgumentNullException(nameof(headingDetector));

            var blockIds = new Dictionary<int, string>();
            for (var blockIndex = 0; blockIndex < logicalBlocks.Count; blockIndex++)
            {
                var block = logicalBlocks[blockIndex];
                var blockId = "block:" + blockIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    ":" + block.StartParagraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    "-" + block.EndParagraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                foreach (var paragraph in snapshot.Paragraphs)
                {
                    if (paragraph.Index >= block.StartParagraphIndex &&
                        paragraph.Index <= block.EndParagraphIndex)
                    {
                        blockIds[paragraph.Index] = blockId;
                    }
                }
            }

            var headings = headingDetector.Detect(snapshot.Paragraphs,
                new HeadingDetectionContext(rolesByParagraphIndex, blockIds));
            return new DerivedAnalysisContext(logicalBlocks,
                rolesByParagraphIndex.ToDictionary(item => item.Key, item => item.Value),
                blockIds, headings);
        }

        public string RoleOf(LocalParagraphSnapshot paragraph)
        {
            string role;
            return RolesByParagraphIndex.TryGetValue(paragraph.Index, out role)
                ? role
                : paragraph.Role;
        }

        public string BlockIdOf(LocalParagraphSnapshot paragraph)
        {
            string blockId;
            return LogicalBlockIdsByParagraphIndex.TryGetValue(paragraph.Index, out blockId)
                ? blockId
                : HeadingDetectionContext.DefaultLogicalBlockId;
        }
    }
}
