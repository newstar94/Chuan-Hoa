using System.Collections.Generic;
using System.Linq;

namespace ChuanHoa.Client.Core.Scanning
{
    public static class HeaderLayoutPolicy
    {
        public static int DateAlignment(IReadOnlyList<LogicalDocumentBlock> blocks, int paragraphIndex)
        {
            var block = blocks.FirstOrDefault(b => b.ContainsParagraph(paragraphIndex));
            if (block == null) return 1;
            // A layout table can enumerate the right-hand date before the
            // left-hand issuing organization. Inspect the whole header region,
            // but never borrow an organization from the following document.
            var headerEnd = block.Roles.Where(r => r.Value == "typeName" || r.Value == "standaloneTitle")
                .Select(r => r.Key).DefaultIfEmpty(block.EndParagraphIndex + 1).Min();
            var header = block.Roles.Where(r => r.Key < headerEnd).Select(r => r.Value).ToArray();
            return header.Contains("nationalTitle") && header.Contains("nationalMotto") &&
                !header.Contains("organName") && !header.Contains("superiorOrganName") ? 2 : 1;
        }
    }
}
