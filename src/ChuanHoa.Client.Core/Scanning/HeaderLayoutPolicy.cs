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
            var header = block.Roles.Where(r => r.Key < paragraphIndex).Select(r => r.Value).ToArray();
            return header.Contains("nationalTitle") && header.Contains("nationalMotto") &&
                !header.Contains("organName") && !header.Contains("superiorOrganName") ? 2 : 1;
        }
    }
}
