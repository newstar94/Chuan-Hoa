export type RegimeType = 'ND30' | 'DANG_HD05' | 'VIETTEL';

export type ComponentRole = 
  | 'nationalTitle'
  | 'motto'
  | 'partyTitle'
  | 'parentOrganName'
  | 'organName'
  | 'codeNumberNotation'
  | 'placeDate'
  | 'typeName'
  | 'subject'
  | 'officialLetterSubject'
  | 'legalBases'
  | 'bodyText'
  | 'partChapterTitle'
  | 'sectionTitle'
  | 'article'
  | 'clause'
  | 'point'
  | 'signAuthority'
  | 'signPosition'
  | 'signFullName'
  | 'recipientsKinhGui'
  | 'recipientsNoiNhan'
  | 'recipientListItem'
  | 'recipientLuuLine'
  | 'appendixHeader'
  | 'appendixTitle'
  | 'appendixBody'
  | 'confidentiality'
  | 'urgency'
  | 'contactInfo'
  | 'authorAndCopies'
  | 'tableContent'
  | 'unknown';

export interface ParagraphSnapshot {
  index: number;
  text: string;
  cleanText: string;
  fontName?: string;
  fontSize?: number;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  alignment?: 'Left' | 'Center' | 'Right' | 'Justify';
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacingPt?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  isInTable?: boolean;
  tableIndex?: number;
  detectedRole?: ComponentRole;
  confidence?: number;
}

export interface DocumentSnapshot {
  documentTitle?: string;
  regime: RegimeType;
  pageSetup: {
    paperSize: string;
    widthMm: number;
    heightMm: number;
    orientation: 'Portrait' | 'Landscape';
    marginsMm: {
      top: number;
      bottom: number;
      left: number;
      right: number;
    };
  };
  hasHeaderPageNumber: boolean;
  firstPageHasNumber: boolean;
  paragraphs: ParagraphSnapshot[];
  tableCount: number;
}

export type IssueSeverity = 'ERROR' | 'WARNING' | 'INFO';

export interface ComplianceIssue {
  id: string;
  checkCode: string;
  paragraphIndex?: number;
  componentRole?: ComponentRole;
  category: string;
  severity: IssueSeverity;
  title: string;
  description: string;
  currentValue?: any;
  expectedValue?: any;
  autoFixable: boolean;
  fixAction?: string;
}

export interface AutoFixAction {
  paragraphIndex: number;
  componentRole: ComponentRole;
  targetFormat: {
    fontName?: string;
    fontSize?: number;
    isBold?: boolean;
    isItalic?: boolean;
    isUnderline?: boolean;
    alignment?: 'Left' | 'Center' | 'Right' | 'Justify';
    firstLineIndentCm?: number;
    spaceBeforePt?: number;
    spaceAfterPt?: number;
    lineSpacingRule?: string;
    lineSpacingPt?: number;
    lineSpacingExactPt?: number;
  };
  textReplacement?: {
    originalText: string;
    normalizedText: string;
  };
  underlineAction?: {
    required: boolean;
    lengthRatio: number;
    lineType: 'SolidSingle';
  };
  separatorAction?: {
    type: 'StarSymbol' | 'FiveHyphens';
    symbol: string;
  };
}

export interface DocumentAnalysisResult {
  detectedRegime: RegimeType;
  regimeName: string;
  totalParagraphs: number;
  totalIssues: number;
  errorCount: number;
  warningCount: number;
  infoCount: number;
  issues: ComplianceIssue[];
  autoFixActions: AutoFixAction[];
  summary: {
    pageSetupValid: boolean;
    typographyScore: number;
    structureScore: number;
    isCompliant: boolean;
  };
}
