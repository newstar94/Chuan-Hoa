using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ChuanHoa.Client.Core.Scanning
{
    public static class LocalDocumentTypeCodes
    {
        public const string Unknown = "UNKNOWN";
        public const string OfficialLetter = "OFFICIAL_LETTER";
        public const string Decision = "DECISION";
        public const string Report = "REPORT";
        public const string Plan = "PLAN";
        public const string Proposal = "PROPOSAL";
        public const string Regulation = "REGULATION";
        public const string Guidance = "GUIDANCE";
        public const string Notice = "NOTICE";
        public const string Minutes = "MINUTES";
        public const string Resolution = "RESOLUTION";
        public const string Directive = "DIRECTIVE";
        public const string Circular = "CIRCULAR";
        public const string Communique = "COMMUNIQUE";
        public const string Program = "PROGRAM";
        public const string Scheme = "SCHEME";
        public const string Project = "PROJECT";
        public const string Option = "OPTION";
        public const string Invitation = "INVITATION";
        public const string Telegram = "TELEGRAM";
        public const string IntroductionLetter = "INTRODUCTION_LETTER";
        public const string LeavePermit = "LEAVE_PERMIT";
        public const string AuthorizationLetter = "AUTHORIZATION_LETTER";
        public const string SendingSlip = "SENDING_SLIP";
        public const string TransferSlip = "TRANSFER_SLIP";
        public const string NotificationSlip = "NOTIFICATION_SLIP";
        public const string Conclusion = "CONCLUSION";

        public static string GetDisplayName(string code)
        {
            switch (code)
            {
                case OfficialLetter: return "Công văn";
                case Resolution: return "Nghị quyết";
                case Decision: return "Quyết định";
                case Directive: return "Chỉ thị";
                case Circular: return "Thông tư";
                case Communique: return "Thông cáo";
                case Notice: return "Thông báo";
                case Guidance: return "Hướng dẫn";
                case Program: return "Chương trình";
                case Plan: return "Kế hoạch";
                case Option: return "Phương án";
                case Scheme: return "Đề án";
                case Project: return "Dự án";
                case Report: return "Báo cáo";
                case Proposal: return "Tờ trình";
                case Regulation: return "Quy chế/Quy định";
                case Invitation: return "Giấy mời";
                case Telegram: return "Công điện";
                case IntroductionLetter: return "Giấy giới thiệu";
                case Minutes: return "Biên bản";
                case LeavePermit: return "Giấy nghỉ phép";
                case AuthorizationLetter: return "Giấy ủy quyền";
                case SendingSlip: return "Phiếu gửi";
                case TransferSlip: return "Phiếu chuyển";
                case NotificationSlip: return "Phiếu báo";
                case Conclusion: return "Kết luận";
                default: return "Chưa xác định";
            }
        }
    }

    /// <summary>
    /// A self-contained administrative document inside the main Word story. A Word
    /// file can contain several such documents, so roles and type are scoped here.
    /// </summary>
    public sealed class LogicalDocumentBlock
    {
        internal LogicalDocumentBlock(int startParagraphIndex, int endParagraphIndex,
            string documentTypeCode, IReadOnlyDictionary<int, string> roles)
        {
            StartParagraphIndex = startParagraphIndex;
            EndParagraphIndex = endParagraphIndex;
            DocumentTypeCode = documentTypeCode ?? LocalDocumentTypeCodes.Unknown;
            Roles = roles ?? throw new ArgumentNullException(nameof(roles));
        }

        public int StartParagraphIndex { get; }
        public int EndParagraphIndex { get; }
        public string DocumentTypeCode { get; }
        public IReadOnlyDictionary<int, string> Roles { get; }

        public bool ContainsParagraph(int paragraphIndex) =>
            paragraphIndex >= StartParagraphIndex && paragraphIndex <= EndParagraphIndex;
    }

    /// <summary>
    /// Resolves semantic roles before applying format rules. Document content is
    /// authoritative. A manual selection is retained only as an explicit fallback for
    /// unusual legacy documents whose type cannot be inferred from any textual signal.
    /// </summary>
    public sealed class DocumentRoleDetector
    {
        private sealed class RoleAssignment
        {
            public RoleAssignment(string role, int confidence, int priority, string evidence)
            {
                Role = role;
                Confidence = confidence;
                Priority = priority;
                Evidence = evidence;
            }

            public string Role { get; }
            public int Confidence { get; }
            public int Priority { get; }
            public string Evidence { get; }
        }

        private static readonly Regex NationalTitle = Rx(@"CỘNG\s+H(?:ÒA|OÀ)\s+XÃ\s+HỘI\s+CHỦ\s+NGHĨA\s+VIỆT\s+NAM", true);
        private static bool IsExtendedReportTitle(string text) =>
            text.Length <= 160 && Regex.IsMatch(text, @"^BÁO CÁO\s+[\p{Lu}\p{M}\d ,/()-]+$");
        private static bool IsTypeHeading(string text) => TypeName.IsMatch(text) ||
            IsExtendedReportTitle(text) || Eq(text, "BẢN CAM KẾT");
        private static readonly Regex PlaceDate = Rx(@"^\s*(?<place>[\p{L}][\p{L}\s.]{0,70}?)(?<comma>,?)\s+ngày\s+(?<day>\d{1,2})\s+tháng\s+(?<month>\d{1,2})\s+năm\s+(?<year>\d{4})\s*$", true);
        private static readonly Regex TypeName = Rx(@"^(?<type>NGHỊ QUYẾT|QUYẾT ĐỊNH|CHỈ THỊ|THÔNG TƯ|THÔNG CÁO|THÔNG BÁO|HƯỚNG DẪN|CHƯƠNG TRÌNH|KẾ HOẠCH|PHƯƠNG ÁN|ĐỀ ÁN|DỰ ÁN|BÁO CÁO|TỜ TRÌNH|QUY CHẾ|QUY ĐỊNH|GIẤY MỜI|CÔNG ĐIỆN|GIẤY GIỚI THIỆU|BIÊN BẢN|GIẤY NGHỈ PHÉP|GIẤY ỦY QUYỀN|PHIẾU GỬI|PHIẾU CHUYỂN|PHIẾU BÁO|KẾT LUẬN)(?:\s*[.:]?\s*\d{1,2})?$", true);
        private static readonly Regex LegalBasis = Rx(@"^(?:[-–—]\s*)?(Căn cứ|Xét|Xét đề nghị|Theo đề nghị)\b", true);
        private static readonly Regex FormalLegalBasis = Rx(
            @"^(?:[-–—]\s*)?(?:" +
            @"(?:Xét|Xét\s+đề\s+nghị|Theo\s+đề\s+nghị)\b|" +
            @"Căn\s+cứ(?:\s+vào)?\s+(?:" +
            VietnameseLegalDocumentVocabulary.RegexAlternation + @"|Thông\s+cáo|Hướng\s+dẫn|Quy\s+chế|Quy\s+định|Điều\s+lệ|" +
            @"Điều\s+\d+|Khoản\s+\d+|Điểm\s+[a-zđ]\)|Văn\s+bản|Công\s+văn|Tờ\s+trình|" +
            @"Kế\s+hoạch|Chương\s+trình|Đề\s+án|Dự\s+án|Hợp\s+đồng|Biên\s+bản|Giấy\s+phép|" +
            @"chức\s+năng|nhiệm\s+vụ|thẩm\s+quyền|đề\s+nghị|yêu\s+cầu|ý\s+kiến\s+chỉ\s+đạo|" +
            @"hồ\s+sơ\s+(?:đề\s+nghị|xin|đăng\s+ký)" +
            @")\b)", true);
        private static readonly Regex SignerAuthority = Rx(@"^(?:(?:TM|KT|TL|TUQ|T/M|K/T|T/L|Q)\.?\s+|CHỦ TỊCH|PHÓ CHỦ TỊCH|GIÁM ĐỐC|PHÓ GIÁM ĐỐC|BỘ TRƯỞNG|THỨ TRƯỞNG|BÍ THƯ|PHÓ BÍ THƯ|TRƯỞNG BAN|PHÓ TRƯỞNG BAN|CHÁNH VĂN PHÒNG|PHÓ CHÁNH VĂN PHÒNG)\b", true);
        private static readonly Regex CodeNumberPattern = Rx(@"^Số\s*:?\s*(?:\d|[.…]+\s*/)");
        private static readonly Regex InlineTableCodeNumberPattern = Rx(@"(?:^|[\r\n\v])\s*Số\s*:?\s*(?:\d|[.…]+\s*/)", true);
        private static readonly Regex CodeNumberSimple = Rx(@"^Số\s*:?\s*\d");
        private static readonly Regex AboutPrefixRegex = Rx(@"^Về\s+việc\b", true);
        private static readonly Regex OfficialLetterSubjectRegex = Rx(@"^(V/v|Về việc)\b", true);
        private static readonly Regex RecipientSalutationPrefixRegex = Rx(@"^Kính\s+(gửi|trình)\s*:", true);
        private static readonly Regex RecipientLabelPattern = Rx(@"^Nơi\s+nhận", true);
        private static readonly Regex RecipientListContinuationRegex = Rx(@"^(Như\s+Điều\b|Lưu\b)", true);
        private static readonly Regex AppendixLabelPattern = Rx(@"^Phụ\s+lục(?:\s+[IVXLCDM\d]+)?\b", true);
        private static readonly Regex PartChapterHeadingRegex = Rx(@"^(Phần|Chương)\s+(?:[IVXLCDM]+|thứ\s+\p{L}+)$", true);
        private static readonly Regex SectionHeadingRegex = Rx(@"^(Mục|Tiểu mục)\s+\d+$", true);
        private static readonly Regex StructuralTitlePrevRegex = Rx(@"^(Phần|Chương|Mục|Tiểu mục)\b", true);
        private static readonly Regex SubjectTerminatorRegex = Rx(@"^(Kính\s+(?:gửi|trình)|Nơi\s+nhận)\b", true);
        private static readonly Regex StructuralBodyStartRegex = Rx(@"^(?:Điều\s+\d+|(?:\d+(?:\.\d+)*|[IVXLCDM]+)[.)]\s+\p{L}|[a-zđ]\)\s+\p{L}|Phần\b|Chương\b|Mục\b|Tiểu\s+mục\b)", true);
        private static readonly Regex PlaceDateDraftPattern = Rx(@"^[\p{L}][\p{L}\s.]{0,70},\s*ngày\s+[.…]+\s+tháng\s+(?:\d{1,2}|[.…]+)\s+năm\s+(?:\d{4}|[.…]+)$", true);
        private static readonly Regex HeaderCodeNumberPattern = Rx(@"^Số\s*:?\s*\d", true);
        private static readonly Regex AppendixDigitalSignatureLabelPattern = Rx(@"^Số\s*:", true);
        private static readonly Regex DigitalSignatureUnitPattern = Rx(@"\b(giờ|phút|giây)\b", true);
        private static readonly Regex DigitalSignatureTimePattern = Rx(@"\b\d{1,2}:\d{2}(?::\d{2})?\b");
        private static readonly Regex AppendixReferencePattern = Rx(@"^\(\s*Kèm\s+theo\b", true);

        public Dictionary<int, string> Detect(LocalScanSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var result = new Dictionary<int, string>();
            foreach (var block in DetectBlocks(snapshot))
                foreach (var role in block.Roles)
                    result[role.Key] = role.Value;
            return result;
        }

        public IReadOnlyList<LogicalDocumentBlock> DetectBlocks(LocalScanSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var main = MainParagraphs(snapshot);
            if (main.Length == 0) return Array.Empty<LogicalDocumentBlock>();

            var starts = FindBlockStarts(main);
            var blocks = new List<LogicalDocumentBlock>(starts.Count);
            for (var blockIndex = 0; blockIndex < starts.Count; blockIndex++)
            {
                var startPosition = starts[blockIndex];
                var endPosition = blockIndex + 1 < starts.Count ? starts[blockIndex + 1] - 1 : main.Length - 1;
                var paragraphs = main.Skip(startPosition).Take(endPosition - startPosition + 1).ToArray();
                var documentType = ResolveDocumentTypeFromContent(paragraphs);
                if (Eq(documentType, LocalDocumentTypeCodes.Unknown) && starts.Count == 1 &&
                    snapshot.DocumentTypeWasSelectedManually &&
                    !Eq(snapshot.DocumentTypeCode, LocalDocumentTypeCodes.Unknown))
                    documentType = snapshot.DocumentTypeCode;
                var roles = DetectBlock(paragraphs, documentType);
                blocks.Add(new LogicalDocumentBlock(paragraphs[0].Index,
                    paragraphs[paragraphs.Length - 1].Index, documentType, roles));
            }
            return blocks;
        }

        private static Dictionary<int, string> DetectBlock(LocalParagraphSnapshot[] main,
            string documentType)
        {
            var result = new Dictionary<int, string>();
            var legalBasisWindowOpen = false;
            var legalBasisSequenceStarted = false;
            var legalBasisGapCount = 0;
            var typeNameAssigned = false;

            for (var i = 0; i < main.Length; i++)
            {
                var paragraph = main[i];
                if (!Eq(paragraph.Role, "Unknown") && !string.IsNullOrWhiteSpace(paragraph.Role))
                {
                    result[paragraph.Index] = paragraph.Role;
                    if (string.Equals(paragraph.Role, "typeName", StringComparison.Ordinal))
                        typeNameAssigned = true;
                    UpdateLegalBasisWindow(paragraph, paragraph.Role, ref legalBasisWindowOpen,
                        ref legalBasisSequenceStarted, ref legalBasisGapCount);
                    continue;
                }

                var text = Collapse(paragraph.Text);
                var previousRole = i > 0 && result.ContainsKey(main[i - 1].Index)
                    ? result[main[i - 1].Index]
                    : string.Empty;

                var assignment = SelectRoleAssignment(main, i, documentType, previousRole,
                    text, typeNameAssigned, legalBasisWindowOpen, legalBasisSequenceStarted);
                var assignedRole = assignment?.Role;
                if (assignedRole == "typeName") typeNameAssigned = true;

                if (assignedRole != null) result[paragraph.Index] = assignedRole;
                UpdateLegalBasisWindow(paragraph, assignedRole, ref legalBasisWindowOpen,
                    ref legalBasisSequenceStarted, ref legalBasisGapCount);
            }

            AssignAppendixRoles(main, result);
            AssignOrganRoles(main, result);
            return result;
        }

        private static RoleAssignment? SelectRoleAssignment(LocalParagraphSnapshot[] main, int index,
            string documentType, string previousRole, string text, bool typeNameAssigned,
            bool legalBasisWindowOpen, bool legalBasisSequenceStarted)
        {
            var paragraph = main[index];
            var candidates = new List<RoleAssignment>();
            AddCandidate(candidates, NationalTitle.IsMatch(text), "nationalTitle", 100, 100,
                "exact national-title text");
            AddCandidate(candidates, Contains(text, "Độc lập") && Contains(text, "Hạnh phúc"),
                "nationalMotto", 100, 99, "national-motto vocabulary");
            AddCandidate(candidates, Eq(text, "ĐẢNG CỘNG SẢN VIỆT NAM"), "partyTitle", 100, 100,
                "exact party title");
            AddCandidate(candidates, CodeNumberPattern.IsMatch(text) ||
                (paragraph.IsInTable && InlineTableCodeNumberPattern.IsMatch(paragraph.Text)),
                "codeNumber", 96, 95, paragraph.IsInTable
                    ? "code-number text in a table cell" : "code-number prefix");
            AddCandidate(candidates, legalBasisWindowOpen && (IsFormalLegalBasisParagraph(text) ||
                (legalBasisSequenceStarted && LegalBasis.IsMatch(text))), "legalBasis", 94, 94,
                "formal legal vocabulary inside the preamble window");
            AddCandidate(candidates, IsPlaceDate(text), "placeAndIssuedDate", 92, 92,
                "place-and-issued-date pattern");

            if (IsTypeHeading(text))
            {
                // A Decision commonly repeats its name as the operative formula before
                // Điều 1. Only the first occurrence is the document type.
                var role = Eq(text, "BẢN CAM KẾT") ? "standaloneTitle" : typeNameAssigned
                    ? "structuralTitle"
                    : "typeName";
                AddCandidate(candidates, true, role, 98, 98, "document-type heading");
            }

            if (previousRole == "typeName")
            {
                var subjectScore = ScoreSubjectCandidate(main, index, text);
                AddCandidate(candidates, subjectScore > 0, "subject", subjectScore, 80,
                    "adjacency, text, formatting and following boundary");
            }
            if ((previousRole == "subject" || previousRole == "subjectContinuation") && index > 0)
            {
                var continuationScore = ScoreSubjectContinuation(main, index, text);
                AddCandidate(candidates, continuationScore > 0, "subjectContinuation",
                    continuationScore, 79, "subject adjacency and style continuity");
            }

            AddCandidate(candidates, documentType == LocalDocumentTypeCodes.OfficialLetter &&
                OfficialLetterSubjectRegex.IsMatch(text), "officialLetterSubject", 96, 96,
                "official-letter subject prefix");
            AddCandidate(candidates, SignerAuthority.IsMatch(text), "signerAuthority", 91, 91,
                "signer-authority vocabulary");
            AddCandidate(candidates, RecipientSalutationPrefixRegex.IsMatch(text),
                text.EndsWith(":", StringComparison.Ordinal) ? "recipientSalutation" : "recipientSalutationInline",
                91, 90, "recipient salutation prefix");
            AddCandidate(candidates, (previousRole == "recipientSalutation" ||
                previousRole == "recipientSalutationList") && text.StartsWith("-", StringComparison.Ordinal),
                "recipientSalutationList", 86, 86, "recipient-list adjacency and marker");
            AddCandidate(candidates, RecipientLabelPattern.IsMatch(text), "recipientLabel", 92, 92,
                "recipient label");
            AddCandidate(candidates, (previousRole == "recipientLabel" || previousRole == "recipientList") &&
                index > 0 && main[index - 1].Index + 1 == paragraph.Index &&
                main[index - 1].TableIndex == paragraph.TableIndex &&
                main[index - 1].CellIndex == paragraph.CellIndex &&
                (text.StartsWith("-", StringComparison.Ordinal) ||
                 !string.IsNullOrWhiteSpace(paragraph.ListMarker) ||
                 RecipientListContinuationRegex.IsMatch(text)), "recipientList", 87, 87,
                "recipient-list marker and structural adjacency");
            AddCandidate(candidates, AppendixLabelPattern.IsMatch(text), "appendixLabel", 95, 95,
                "appendix label");
            AddCandidate(candidates, PartChapterHeadingRegex.IsMatch(text), "partChapterHeading", 93, 93,
                "part/chapter heading");
            AddCandidate(candidates, SectionHeadingRegex.IsMatch(text), "sectionHeading", 93, 93,
                "section heading");
            AddCandidate(candidates, IsStructuralTitle(main, index, text), "structuralTitle", 75, 70,
                "uppercase title following a structural label");

            return candidates.OrderByDescending(candidate => candidate.Confidence)
                .ThenByDescending(candidate => candidate.Priority)
                .FirstOrDefault();
        }

        private static void AddCandidate(ICollection<RoleAssignment> candidates, bool condition,
            string role, int confidence, int priority, string evidence)
        {
            if (condition)
                candidates.Add(new RoleAssignment(role, confidence, priority, evidence));
        }

        private static LocalParagraphSnapshot[] MainParagraphs(LocalScanSnapshot snapshot) =>
            snapshot.Paragraphs
                .Where(p => string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    !string.IsNullOrWhiteSpace(p.Text))
                .OrderBy(p => p.Index)
                .ToArray();

        private static List<int> FindBlockStarts(LocalParagraphSnapshot[] main)
        {
            var starts = new List<int> { 0 };
            var currentStart = 0;
            var currentHasDocumentIdentity = false;

            for (var position = 0; position < main.Length; position++)
            {
                var text = Collapse(main[position].Text);
                var explicitTypeName = string.Equals(main[position].Role, "typeName",
                    StringComparison.Ordinal);
                if (explicitTypeName)
                {
                    if (currentHasDocumentIdentity && position > currentStart)
                    {
                        starts.Add(position);
                        currentStart = position;
                    }
                    currentHasDocumentIdentity = true;
                    continue;
                }
                var isNationalOrPartyTitle = NationalTitle.IsMatch(text) ||
                    Eq(text, "ĐẢNG CỘNG SẢN VIỆT NAM");

                if (isNationalOrPartyTitle && currentHasDocumentIdentity &&
                    HasHeaderClusterAhead(main, position))
                {
                    var candidate = BacktrackHeaderStart(main, position, currentStart);
                    if (candidate > currentStart)
                    {
                        starts.Add(candidate);
                        currentStart = candidate;
                        currentHasDocumentIdentity = false;
                    }
                }

                if (IsTypeHeading(text) || OfficialLetterSubjectRegex.IsMatch(text))
                {
                    if (currentHasDocumentIdentity && IsTypeHeading(text))
                    {
                        var headerStart = FindHeaderClusterStartBefore(main, position, currentStart);
                        if (headerStart > currentStart)
                        {
                            starts.Add(headerStart);
                            currentStart = headerStart;
                            currentHasDocumentIdentity = false;
                        }
                    }
                    currentHasDocumentIdentity = true;
                }
            }

            return starts.Distinct().OrderBy(value => value).ToList();
        }

        private static bool HasHeaderClusterAhead(LocalParagraphSnapshot[] main, int titlePosition)
        {
            var signals = Contains(main[titlePosition].Text, "Độc lập") &&
                Contains(main[titlePosition].Text, "Hạnh phúc") ? 1 : 0;
            var end = Math.Min(main.Length - 1, titlePosition + 12);
            for (var position = titlePosition + 1; position <= end; position++)
            {
                var text = Collapse(main[position].Text);
                if (Contains(text, "Độc lập") && Contains(text, "Hạnh phúc")) signals++;
                else if (CodeNumberSimple.IsMatch(text)) signals++;
                else if (IsPlaceDate(text)) signals++;
                else if (IsTypeHeading(text)) signals++;
                else if (OfficialLetterSubjectRegex.IsMatch(text)) signals++;
                if (signals >= 2) return true;
                if (IsStructuralBodyStart(text)) break;
            }
            return false;
        }

        private static int BacktrackHeaderStart(LocalParagraphSnapshot[] main, int titlePosition,
            int currentStart)
        {
            var start = titlePosition;
            var titlePage = main[titlePosition].PageNumber;
            for (var position = titlePosition - 1;
                position >= currentStart && titlePosition - position <= 3;
                position--)
            {
                var paragraph = main[position];
                var text = Collapse(paragraph.Text);
                if (titlePage > 0 && paragraph.PageNumber > 0 && paragraph.PageNumber != titlePage) break;
                if (main[position + 1].Index - paragraph.Index > 2) break;
                if (!IsLikelyOrganHeading(text)) break;
                start = position;
            }
            return start;
        }

        private static int FindHeaderClusterStartBefore(LocalParagraphSnapshot[] main, int typePosition,
            int currentStart)
        {
            var earliestSignal = -1;
            var signalCount = 0;
            for (var position = typePosition - 1;
                position > currentStart && typePosition - position <= 12;
                position--)
            {
                var text = Collapse(main[position].Text);
                if (NationalTitle.IsMatch(text) ||
                    (Contains(text, "Độc lập") && Contains(text, "Hạnh phúc")) ||
                    HeaderCodeNumberPattern.IsMatch(text) || IsPlaceDate(text))
                {
                    earliestSignal = position;
                    signalCount++;
                }
            }
            if (signalCount < 2 || earliestSignal < 0) return currentStart;
            return BacktrackHeaderStart(main, earliestSignal, currentStart);
        }

        private static bool IsLikelyOrganHeading(string text)
        {
            if (text.Length < 4 || text.Length > 220 || !IsMostlyUppercase(text)) return false;
            if (SignerAuthority.IsMatch(text) || IsTypeHeading(text) || IsStructuralBodyStart(text)) return false;
            return !Eq(text, "QUỐC HỘI");
        }

        private static void AssignAppendixRoles(IReadOnlyList<LocalParagraphSnapshot> main,
            IDictionary<int, string> roles)
        {
            for (var i = 0; i < main.Count; i++)
            {
                string role;
                if (!roles.TryGetValue(main[i].Index, out role) || role != "appendixLabel") continue;

                for (var previous = Math.Max(0, i - 3); previous < i; previous++)
                {
                    var text = Collapse(main[previous].Text);
                    if (AppendixDigitalSignatureLabelPattern.IsMatch(text) &&
                        (DigitalSignatureUnitPattern.IsMatch(text) ||
                         DigitalSignatureTimePattern.IsMatch(text)))
                        roles[main[previous].Index] = "appendixDigitalSignatureInfo";
                }

                var candidates = main.Skip(i + 1)
                    .TakeWhile(item => item.SectionIndex == main[i].SectionIndex)
                    .Take(5)
                    .ToArray();
                var reference = candidates.FirstOrDefault(item =>
                    AppendixReferencePattern.IsMatch(Collapse(item.Text)));
                if (reference != null)
                {
                    foreach (var title in candidates.TakeWhile(item => item.Index < reference.Index))
                        if (!roles.ContainsKey(title.Index)) roles[title.Index] = "appendixTitle";
                    roles[reference.Index] = "appendixReference";
                }
                else if (candidates.Length > 0 && !roles.ContainsKey(candidates[0].Index))
                {
                    // Model 2.1/2.2 places the title immediately below the label. Keep
                    // that role even when the mandatory reference line is missing so
                    // the scanner can report and 1-Click can format both defects.
                    roles[candidates[0].Index] = "appendixTitle";
                }
            }
        }

        public string ResolveDocumentType(LocalScanSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var main = MainParagraphs(snapshot);
            return ResolveDocumentType(snapshot, main);
        }

        public string ResolveDocumentTypeFromContent(LocalScanSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            return ResolveDocumentTypeFromContent(snapshot.Paragraphs
                .Where(p => string.Equals(p.StoryType, "wdMainTextStory", StringComparison.Ordinal) &&
                    !string.IsNullOrWhiteSpace(p.Text))
                .OrderBy(p => p.Index)
                .ToArray());
        }

        private static string ResolveDocumentType(LocalScanSnapshot snapshot, LocalParagraphSnapshot[] main)
        {
            var detected = ResolveDocumentTypeFromContent(main);
            if (!Eq(detected, LocalDocumentTypeCodes.Unknown)) return detected;

            if (snapshot.DocumentTypeWasSelectedManually &&
                !Eq(snapshot.DocumentTypeCode, LocalDocumentTypeCodes.Unknown))
                return snapshot.DocumentTypeCode;
            return LocalDocumentTypeCodes.Unknown;
        }

        private static string ResolveDocumentTypeFromContent(LocalParagraphSnapshot[] main)
        {
            var typeParagraph = main.FirstOrDefault(p => IsTypeHeading(Collapse(p.Text)));
            if (typeParagraph != null)
            {
                if (IsExtendedReportTitle(Collapse(typeParagraph.Text))) return LocalDocumentTypeCodes.Report;
                if (Eq(Collapse(typeParagraph.Text), "BẢN CAM KẾT")) return LocalDocumentTypeCodes.Unknown;
                var match = TypeName.Match(Collapse(typeParagraph.Text));
                var type = match.Groups["type"].Value.ToUpper(CultureInfo.GetCultureInfo("vi-VN"));
                return MapTypeName(type);
            }

            if (main.Any(p => OfficialLetterSubjectRegex.IsMatch(Collapse(p.Text))) &&
                main.All(p => !IsTypeHeading(Collapse(p.Text))))
                return LocalDocumentTypeCodes.OfficialLetter;
            return LocalDocumentTypeCodes.Unknown;
        }

        private static string MapTypeName(string type)
        {
            switch (type)
            {
                case "NGHỊ QUYẾT": return LocalDocumentTypeCodes.Resolution;
                case "QUYẾT ĐỊNH": return LocalDocumentTypeCodes.Decision;
                case "CHỈ THỊ": return LocalDocumentTypeCodes.Directive;
                case "THÔNG TƯ": return LocalDocumentTypeCodes.Circular;
                case "THÔNG CÁO": return LocalDocumentTypeCodes.Communique;
                case "THÔNG BÁO": return LocalDocumentTypeCodes.Notice;
                case "HƯỚNG DẪN": return LocalDocumentTypeCodes.Guidance;
                case "CHƯƠNG TRÌNH": return LocalDocumentTypeCodes.Program;
                case "KẾ HOẠCH": return LocalDocumentTypeCodes.Plan;
                case "PHƯƠNG ÁN": return LocalDocumentTypeCodes.Option;
                case "ĐỀ ÁN": return LocalDocumentTypeCodes.Scheme;
                case "DỰ ÁN": return LocalDocumentTypeCodes.Project;
                case "BÁO CÁO": return LocalDocumentTypeCodes.Report;
                case "TỜ TRÌNH": return LocalDocumentTypeCodes.Proposal;
                case "QUY CHẾ":
                case "QUY ĐỊNH": return LocalDocumentTypeCodes.Regulation;
                case "GIẤY MỜI": return LocalDocumentTypeCodes.Invitation;
                case "CÔNG ĐIỆN": return LocalDocumentTypeCodes.Telegram;
                case "GIẤY GIỚI THIỆU": return LocalDocumentTypeCodes.IntroductionLetter;
                case "BIÊN BẢN": return LocalDocumentTypeCodes.Minutes;
                case "GIẤY NGHỈ PHÉP": return LocalDocumentTypeCodes.LeavePermit;
                case "GIẤY ỦY QUYỀN": return LocalDocumentTypeCodes.AuthorizationLetter;
                case "PHIẾU GỬI": return LocalDocumentTypeCodes.SendingSlip;
                case "PHIẾU CHUYỂN": return LocalDocumentTypeCodes.TransferSlip;
                case "PHIẾU BÁO": return LocalDocumentTypeCodes.NotificationSlip;
                case "KẾT LUẬN": return LocalDocumentTypeCodes.Conclusion;
                default: return LocalDocumentTypeCodes.Unknown;
            }
        }

        private static void AssignOrganRoles(LocalParagraphSnapshot[] main, IDictionary<int, string> roles)
        {
            var firstCode = main.FirstOrDefault(p => roles.ContainsKey(p.Index) && roles[p.Index] == "codeNumber");
            var candidates = firstCode != null
                ? main.Where(p => p.Index < firstCode.Index && !roles.ContainsKey(p.Index) &&
                    Collapse(p.Text).Length < 220 && IsMostlyUppercase(p.Text) &&
                    !Eq(Collapse(p.Text), "QUỐC HỘI")).ToArray()
                : OrganCandidatesInSiblingHeaderCells(main, roles);
            var preceding = candidates.Skip(Math.Max(0, candidates.Length - 2)).ToArray();
            if (preceding.Length == 2) roles[preceding[0].Index] = "superiorOrganName";
            if (preceding.Length > 0) roles[preceding[preceding.Length - 1].Index] = "organName";
        }

        private static LocalParagraphSnapshot[] OrganCandidatesInSiblingHeaderCells(
            LocalParagraphSnapshot[] main, IDictionary<int, string> roles)
        {
            var nationalHeaderCells = main.Where(p => p.IsInTable && p.TableIndex.HasValue &&
                p.RowIndex.HasValue && p.CellIndex.HasValue && roles.TryGetValue(p.Index, out var role) &&
                (role == "nationalTitle" || role == "nationalMotto")).ToArray();
            if (nationalHeaderCells.Length == 0) return Array.Empty<LocalParagraphSnapshot>();

            // Some valid identity layouts omit the document-number row and put the
            // two organ headings in the cell immediately left of the national header.
            // Restrict the fallback to that exact table row/cell relationship so an
            // unrelated uppercase paragraph elsewhere in the document is never used.
            return main.Where(p => p.IsInTable && p.TableIndex.HasValue && p.RowIndex.HasValue &&
                p.CellIndex.HasValue && !roles.ContainsKey(p.Index) &&
                Collapse(p.Text).Length < 220 && IsMostlyUppercase(p.Text) &&
                !Eq(Collapse(p.Text), "QUỐC HỘI") && nationalHeaderCells.Any(header =>
                    header.TableIndex == p.TableIndex && header.RowIndex == p.RowIndex &&
                    header.CellIndex == p.CellIndex + 1)).ToArray();
        }

        private static bool IsStructuralTitle(LocalParagraphSnapshot[] main, int index, string text)
        {
            if (!IsMostlyUppercase(text) || text.Length < 4 || text.Length > 180 || index == 0) return false;
            var previous = Collapse(main[index - 1].Text);
            return StructuralTitlePrevRegex.IsMatch(previous);
        }

        private static int ScoreSubjectContinuation(LocalParagraphSnapshot[] main,
            int index, string text)
        {
            var previous = main[index - 1];
            var current = main[index];
            // Multi-line subjects are frequently stored as consecutive Word
            // paragraphs. A blank paragraph (visible through the source index gap)
            // terminates the component, as do normal body/legal/operative starts.
            if (current.Index != previous.Index + 1 || text.Length == 0)
                return 0;
            if (LegalBasis.IsMatch(text) || IsTypeHeading(text) ||
                IsStructuralBodyStart(text) ||
                SubjectTerminatorRegex.IsMatch(text))
                return 0;
            if (IsMostlyUppercase(text)) return 0;

            var score = 25; // consecutive paragraph in the same logical block
            if (current.Alignment == 1) score += 15;
            if (current.Bold.GetValueOrDefault()) score += 15;
            if (current.FontSizePoints >= 12.5d && current.FontSizePoints <= 14.5d) score += 8;
            if (previous.Alignment.HasValue && current.Alignment == previous.Alignment) score += 12;
            if (previous.Bold.HasValue && current.Bold == previous.Bold) score += 12;
            if (text.Length <= 240) score += 5;
            if (HasSubjectBoundaryAhead(main, index)) score += 8;
            if (text.EndsWith(".", StringComparison.Ordinal) ||
                text.EndsWith(";", StringComparison.Ordinal)) score -= 20;
            return score >= 50 ? score : 0;
        }

        private static int ScoreSubjectCandidate(LocalParagraphSnapshot[] main, int index, string text)
        {
            var paragraph = main[index];
            if (string.IsNullOrWhiteSpace(text) || LegalBasis.IsMatch(text) ||
                IsStructuralBodyStart(text) || SubjectTerminatorRegex.IsMatch(text))
                return 0;

            var score = 30; // immediately follows the document type
            if (AboutPrefixRegex.IsMatch(text)) score += 40;
            if (paragraph.Alignment == 1) score += 18;
            if (paragraph.Bold.GetValueOrDefault()) score += 18;
            if (paragraph.FontSizePoints >= 12.5d && paragraph.FontSizePoints <= 14.5d) score += 8;
            if (HasSubjectBoundaryAhead(main, index)) score += 18;
            if (text.Length <= 240) score += 6;
            else if (text.Length <= 600) score += 2;
            else score -= 30;
            if (text.EndsWith(".", StringComparison.Ordinal) ||
                text.EndsWith(";", StringComparison.Ordinal)) score -= 15;
            return score >= 55 ? score : 0;
        }

        private static bool HasSubjectBoundaryAhead(LocalParagraphSnapshot[] main, int index)
        {
            var end = Math.Min(main.Length - 1, index + 6);
            for (var candidateIndex = index + 1; candidateIndex <= end; candidateIndex++)
            {
                var candidate = Collapse(main[candidateIndex].Text);
                if (IsFormalLegalBasisParagraph(candidate) || IsStructuralBodyStart(candidate) ||
                    SignerAuthority.IsMatch(candidate) || RecipientSalutationPrefixRegex.IsMatch(candidate))
                    return true;
                if (IsTypeHeading(candidate) && candidateIndex > index + 1) return true;
            }
            return false;
        }

        private static void UpdateLegalBasisWindow(LocalParagraphSnapshot paragraph, string? assignedRole,
            ref bool windowOpen, ref bool sequenceStarted, ref int gapCount)
        {
            if (string.Equals(assignedRole, "typeName", StringComparison.Ordinal))
            {
                windowOpen = true;
                sequenceStarted = false;
                gapCount = 0;
                return;
            }
            if (!windowOpen) return;
            if (string.Equals(assignedRole, "legalBasis", StringComparison.Ordinal))
            {
                sequenceStarted = true;
                gapCount = 0;
                return;
            }
            if (!sequenceStarted && IsPreambleBridge(paragraph, assignedRole)) return;
            // After the legal-basis sequence has started, allow up to 2 gap
            // paragraphs that look like continuations of a previous basis entry
            // (e.g., a line that doesn't start with "Căn cứ" but continues
            // a citation, or a short blank/formatting paragraph between entries).
            if (sequenceStarted && IsLegalBasisContinuationBridge(paragraph, assignedRole))
            {
                gapCount++;
                if (gapCount <= 2) return;
            }
            windowOpen = false;
            sequenceStarted = false;
            gapCount = 0;
        }

        private static bool IsPreambleBridge(LocalParagraphSnapshot paragraph, string? assignedRole)
        {
            if (paragraph.IsInTable) return false;
            if (string.Equals(assignedRole, "subject", StringComparison.Ordinal) ||
                string.Equals(assignedRole, "subjectContinuation", StringComparison.Ordinal) ||
                string.Equals(assignedRole, "signerAuthority", StringComparison.Ordinal))
                return true;
            var text = Collapse(paragraph.Text);
            return text.Length <= 220 && !IsStructuralBodyStart(text) && IsMostlyUppercase(text);
        }

        /// <summary>
        /// After a legal-basis sequence has started, short continuation
        /// paragraphs may appear between "Căn cứ..." entries.  These are
        /// typically: (1) the tail of a long citation that wrapped onto a
        /// new Word paragraph, (2) a blank paragraph used for spacing, or
        /// (3) a short structural formula such as "QUYẾT ĐỊNH:" that
        /// immediately follows the last basis before the operative part.
        /// Only up to 2 such gaps are tolerated before the window closes.
        /// </summary>
        private static bool IsLegalBasisContinuationBridge(LocalParagraphSnapshot paragraph, string? assignedRole)
        {
            if (paragraph.IsInTable) return false;
            // A structural body start (Điều, Khoản, …) definitively ends the preamble.
            var text = Collapse(paragraph.Text);
            if (IsStructuralBodyStart(text)) return false;
            // A type-name repetition (e.g., "QUYẾT ĐỊNH:" before Điều 1) is a
            // legitimate bridge between the last basis and the operative section.
            if (string.Equals(assignedRole, "structuralTitle", StringComparison.Ordinal)) return true;
            // Short paragraphs (≤300 chars) that continue a citation or are
            // blank formatting paragraphs are acceptable bridges.
            if (text.Length <= 300) return true;
            return false;
        }

        private static bool IsFormalLegalBasisParagraph(string text)
        {
            // “Căn cứ” is also an ordinary causal phrase in report content, for
            // example: “Căn cứ các tài liệu được cung cấp, kết quả thẩm định…”.
            // A position inside the preamble is therefore necessary but not enough:
            // the phrase must also introduce a recognized normative/administrative
            // source, authority, task or formal proposal.
            return FormalLegalBasis.IsMatch(text);
        }

        private static bool IsStructuralBodyStart(string text)
        {
            return StructuralBodyStartRegex.IsMatch(text);
        }

        private static bool IsPlaceDate(string text)
        {
            if (LegalBasis.IsMatch(text) || text.IndexOf(';') >= 0) return false;
            var match = PlaceDate.Match(text);
            if (!match.Success && PlaceDateDraftPattern.IsMatch(text))
                return true;
            if (!match.Success) return false;
            var place = match.Groups["place"].Value.Trim();
            return place.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length <= 8;
        }

        private static bool Contains(string value, string token) =>
            value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool Eq(string left, string right) =>
            string.Equals(Collapse(left), Collapse(right), StringComparison.OrdinalIgnoreCase);

        private static string Collapse(string value) =>
            Regex.Replace(value ?? string.Empty, @"\s+", " ").Trim();

        private static bool IsMostlyUppercase(string value)
        {
            var letters = value.Where(char.IsLetter).ToArray();
            return letters.Length >= 4 && letters.Count(char.IsUpper) >= Math.Ceiling(letters.Length * .8d);
        }

        private static Regex Rx(string pattern, bool ignoreCase = false) => new Regex(pattern,
            RegexOptions.CultureInvariant | RegexOptions.Compiled | (ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None),
            TimeSpan.FromMilliseconds(200));
    }
}
