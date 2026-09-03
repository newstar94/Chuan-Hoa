# Local Rule Port Ledger

Nguồn sự thật: canonical draft 96 rule. Phạm vi sản phẩm gồm 73 `BaselineLogicPath`; 19 `HardwiredNotChecked`, 2 `Unrouted` và 2 route tone/i-y đã loại không được tính là implementation.

| Rule code | VBA function | Lane | C# detector | Positive / negative / boundary evidence | Status |
| --- | --- | --- | --- | --- | --- |
| `LOCAL-TYPO-DICT` | `CheckDictionaryTypo` | spelling | `CheckDictionary` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `LOCAL-TYPO-HIDDEN` | `CheckHiddenCharacters` | spelling | `CheckHiddenCharacters` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `LOCAL-TYPO-PUNCT` | `CheckPunctuationSpacingTypo` | spelling | `AddMatches` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `LOCAL-TYPO-SPACE` | `CheckExtraSpaceTypo` | spelling | `AddMatches` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `LOCAL-TYPO-TELEX` | `CheckTelexLeftoverTypo` | spelling | `CheckTelex` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M1-K1` | `CheckPageSizeA4` | format | `CheckPageSetup` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M1-K2` | `CheckPageOrientation` | format | `CheckPageSetup` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M1-K3` | `CheckPageMargins` | format | `CheckPageSetup` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M1-K4-COLOR` | `CheckBodyTextFontColor` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M1-K4-FONT` | `CheckBodyTextFontName` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M1-K7` | `CheckPageNumbering` | format | `CheckPageSetup` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K1-C` | `CheckNationalMottoSpacing` | format | `CheckComponents` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K1-QH` | `CheckNationalTitleStyle` | format | `CheckComponents` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K1-TN` | `CheckNationalMottoStyle` | format | `CheckComponents` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K1-TN-SEP` | `CheckNationalMottoSeparator` | format | `CheckComponents` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K2-ORG` | `CheckOrganNameStyle` | format | `CheckComponents` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K2-SUP` | `CheckSuperiorOrganNameStyle` | format | `CheckComponents` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K3-ABBR` | `CheckCodeNumberAbbreviation` | format | `CheckCodeNumber` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K3-CASE` | `CheckCodeNumberNotationUppercase` | format | `CheckCodeNumber` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K3-PAD` | `CheckCodeNumberPad` | format | `CheckCodeNumber` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K3-PREFIX` | `CheckCodeNumberColon` | format | `CheckCodeNumber` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K3-SEP` | `CheckCodeNumberSeparators` | format | `CheckCodeNumber` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K3-SPACE` | `CheckCodeNumberNoSpace` | format | `CheckCodeNumber` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K4-CASE` | `CheckPlaceNameLetterCase` | format | `CheckPlaceDate` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K4-COMMA` | `CheckPlaceDateComma` | format | `CheckPlaceDate` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K4-PAD` | `CheckPlaceDatePad` | format | `CheckPlaceDate` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K4-STYLE` | `CheckPlaceDateStyle` | format | `CheckPlaceDate` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K5A-SUBJ` | `CheckSubjectStyle` | format | `CheckTypeAndSubject` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K5A-TYPE` | `CheckTypeNameStyle` | format | `CheckTypeAndSubject` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K5B-SPACE` | `CheckSubjectOfficialLetterSpacing` | format | `CheckTypeAndSubject` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K5B-STYLE` | `CheckSubjectOfficialLetterStyle` | format | `CheckTypeAndSubject` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6A-PUNCT` | `CheckLegalBasisPunctuation` | format | `CheckLegalBasisAndCitations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6A-STYLE` | `CheckLegalBasisStyle` | format | `CheckLegalBasisAndCitations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6B-CITE` | `CheckCitationFullFirstCitation` | format | `CheckLegalBasisAndCitations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6B-DATE` | `CheckCitationAbbreviatedDate` | spelling | `CheckBareShortDates` | `LOCAL-RULE-PORT-001`, `LEGAL-BASIS-SHORT-DATE-001` | `ngày 05/03/2020` được chấp nhận; `05/03/2020` thiếu từ `ngày` được neo chính xác và hướng dẫn thêm từ `ngày` |
| `ND30-PL1-M2-K6B-SO` | `CheckCitationMissingSoKeyword` | format | `CheckLegalBasisAndCitations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6D-ALPHABET` | `CheckPointAlphabetOrder` | format | `CheckStructure` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6D-ARTICLE` | `CheckArticleFormat` | format | `CheckStructure` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6D-CLAUSE` | `CheckClauseFormat` | format | `CheckStructure` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6D-POINT` | `CheckPointFormat` | format | `CheckStructure` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6D-TITLE` | `CheckStructureTitlePresence` | format | `CheckStructure` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6E-ALIGN` | `CheckBodyTextAlignment` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6E-DOTSLASH` | `CheckContentEndsWithDotSlash` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6E-INDENT` | `CheckBodyTextFirstLineIndent` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6E-LINESPACING` | `CheckBodyTextLineSpacing` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K6E-SPACEAFTER` | `CheckBodyTextSpaceAfter` | format | `CheckBodyTypography` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K7B-AUTH` | `CheckSignerAuthorityAbbreviation` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K7D-STYLE` | `CheckSignerAuthorityStyle` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9A-COLON` | `CheckRecipientSalutationColon` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9A-INLINE-END` | `CheckRecipientSalutationInlineEnd` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9A-LAYOUT` | `CheckRecipientSalutationLayout` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9A-PUNCT` | `CheckRecipientSalutationPunctuation` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9B-LABEL` | `CheckRecipientLabelStyle` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9B-LIST` | `CheckRecipientListStyle` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M2-K9B-LUU` | `CheckRecipientListLuuLine` | format | `CheckSignerAndRecipients` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M3-K1A-NUM` | `CheckAppendixRomanNumbering` | format | `CheckAppendices` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M3-K1A-REF` | `CheckAppendixReferenceMentioned` | format | `CheckAppendices` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M3-K1B` | `CheckAppendixTitleStyle` | format | `CheckAppendices` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M3-K1C` | `CheckAppendixReferenceInfoStyle` | format | `CheckAppendices` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-M3-K1D` | `CheckAppendixPageNumberingRestart` | format | `CheckAppendices` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL1-MV-CT1` | `CheckFontSizeConsistency` | format | `CheckFontSizeConsistency` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M1` | `CheckSentenceCapitalization` | spelling | `CheckSentenceCapitalization` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M2-K1` | `CheckPersonNameCapitalizationWarn` | spelling | `CheckPersonNames` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M3-K1A` | `CheckAdministrativeUnitNameWarn` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M3-K1B` | `CheckAdministrativeUnitNumeralCase` | spelling | `CheckAdministrativeNumerals` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M3-K1C` | `CheckSpecialGeographicCapitalization` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M3-K1D` | `CheckTerrainPlaceNameCapitalization` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M3-K1E` | `CheckRegionNameCapitalization` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M4-K1A` | `CheckCommonOrganNameWarn` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M4-K1B` | `CheckSpecialOrganCapitalization` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M5-K5` | `CheckHolidayNameCapitalization` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M5-K7` | `CheckArticleClauseCapitalization` | spelling | `CheckArticleClauseCapitalization` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |
| `ND30-PL2-M5-K8A` | `CheckLunarYearCapitalization` | spelling | `CheckConfiguredCapitalizations` | `LOCAL-RULE-PORT-001` | IMPLEMENTED_LOCAL |

## Detector bổ sung trực tiếp từ tài liệu chuẩn

Các detector dưới đây không được tính ngược vào 73 route VBA baseline. Chúng được bổ sung vì audit ba tài liệu chuẩn xác nhận đường kẻ là đối tượng Line Shape, trong khi route VBA cũ `CheckComponentUnderline` là hardwired `NotChecked`.

| Rule code | Nguồn | Lane | C# detector | Positive / negative / boundary evidence | Status |
| --- | --- | --- | --- | --- | --- |
| `ND30-PL1-M2-K1-TN-LINE` | Phụ lục I NĐ30, Phần I.I.1.b | format | `CheckRequiredLineShapes` | `WORD-LINE-SHAPE-001` | IMPLEMENTED_LOCAL_SOURCE_BACKED |
| `ND30-PL1-M2-K2-ORG-LINE` | Phụ lục I NĐ30, Phần I.I.2.b | format | `CheckRequiredLineShapes` | `WORD-LINE-SHAPE-001` | IMPLEMENTED_LOCAL_SOURCE_BACKED |
| `ND30-PL1-M2-K5A-SUBJ-LINE` | Phụ lục I NĐ30, Phần I.I.5.b | format | `CheckRequiredLineShapes` | `WORD-LINE-SHAPE-001` | IMPLEMENTED_LOCAL_SOURCE_BACKED |
| `HD05-M1-TITLE-LINE` | HD05, Phần II.I.1.2 | format | `CheckRequiredLineShapes` | `WORD-LINE-SHAPE-001` | IMPLEMENTED_LOCAL_SOURCE_BACKED |

## Gate còn mở

- `IMPLEMENTED_LOCAL` xác nhận source, synthetic positive/negative/boundary corpus, exact anchor và Word 16 x64 DOC/DOCX smoke.
- Trạng thái này không thay thế legal sign-off, golden corpus văn bản thật, Word 2010/x86 matrix hoặc production signing.
- Gói quy tắc Development có chữ ký; Release không tin khóa Development.
