import { DocumentSnapshot, AutoFixAction, DocumentAnalysisResult } from '../types';
import { ComponentDetector } from '../detector/ComponentDetector';
import { ComplianceChecker } from '../checker/ComplianceChecker';
import { TextNormalizer } from '../normalizer/TextNormalizer';
import { RuleLoader } from '../rules/RuleLoader';

export class AutoFixEngine {
  private static ruleLoader = RuleLoader.getInstance();

  /**
   * Phân tích và tạo kế hoạch 1-Click Auto-Fix toàn diện
   */
  public static analyzeAndPlanAutoFix(snapshot: DocumentSnapshot): DocumentAnalysisResult {
    // 1. Phân loại vai trò các đoạn
    const classifiedParagraphs = ComponentDetector.classifyParagraphs(snapshot.paragraphs, snapshot.regime);
    snapshot.paragraphs = classifiedParagraphs;

    // 2. Chạy 82 mã kiểm tra
    const issues = ComplianceChecker.checkDocument(snapshot);

    // 3. Lấy cấu hình của Regime
    const rule = this.ruleLoader.getRegimeRule(snapshot.regime);
    const regime = snapshot.regime;

    // 4. Sinh danh sách hành động Auto-Fix
    const autoFixActions: AutoFixAction[] = [];

    for (const p of snapshot.paragraphs) {
      const role = p.detectedRole || 'bodyText';
      const originalText = p.text;
      const normalizedText = TextNormalizer.normalizeAll(originalText);

      const action: AutoFixAction = {
        paragraphIndex: p.index,
        componentRole: role,
        targetFormat: {
          fontName: 'Times New Roman'
        }
      };

      if (originalText !== normalizedText) {
        action.textReplacement = {
          originalText,
          normalizedText
        };
      }

      // Format specs theo từng vai trò
      switch (role) {
        case 'nationalTitle':
          action.targetFormat.fontSize = regime === 'VIETTEL' ? 12 : 12;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          action.targetFormat.spaceBeforePt = 0;
          action.targetFormat.spaceAfterPt = 0;
          action.textReplacement = {
            originalText,
            normalizedText: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'
          };
          break;

        case 'motto':
          action.targetFormat.fontSize = regime === 'VIETTEL' ? 13 : 13;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          action.targetFormat.spaceBeforePt = 0;
          action.targetFormat.spaceAfterPt = 0;
          action.underlineAction = {
            required: true,
            lengthRatio: 1.0,
            lineType: 'SolidSingle'
          };
          action.textReplacement = {
            originalText,
            normalizedText: 'Độc lập - Tự do - Hạnh phúc'
          };
          break;

        case 'partyTitle':
          action.targetFormat.fontSize = 15;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          action.underlineAction = {
            required: true,
            lengthRatio: 1.0,
            lineType: 'SolidSingle'
          };
          action.textReplacement = {
            originalText,
            normalizedText: 'ĐẢNG CỘNG SẢN VIỆT NAM'
          };
          break;

        case 'parentOrganName':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 12;
          action.targetFormat.isBold = false;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          break;

        case 'organName':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 12;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          if (regime === 'DANG_HD05') {
            action.separatorAction = {
              type: 'StarSymbol',
              symbol: '*'
            };
          } else {
            action.underlineAction = {
              required: true,
              lengthRatio: 0.4,
              lineType: 'SolidSingle'
            };
          }
          break;

        case 'codeNumberNotation':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 13;
          action.targetFormat.isBold = false;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          break;

        case 'placeDate':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 13;
          action.targetFormat.isBold = false;
          action.targetFormat.isItalic = true;
          action.targetFormat.alignment = 'Center';
          break;

        case 'typeName':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 15 : 14;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          action.targetFormat.spaceBeforePt = 12;
          action.targetFormat.spaceAfterPt = 0;
          break;

        case 'subject':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 14;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          action.targetFormat.spaceAfterPt = 6;
          if (regime === 'DANG_HD05') {
            action.separatorAction = {
              type: 'FiveHyphens',
              symbol: '-----'
            };
          } else {
            action.underlineAction = {
              required: true,
              lengthRatio: 0.4,
              lineType: 'SolidSingle'
            };
          }
          break;

        case 'officialLetterSubject':
          action.targetFormat.fontSize = 12;
          action.targetFormat.isBold = false;
          action.targetFormat.isItalic = regime === 'DANG_HD05';
          action.targetFormat.alignment = 'Center';
          action.targetFormat.spaceBeforePt = 6;
          break;

        case 'legalBases':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 14;
          action.targetFormat.isBold = false;
          action.targetFormat.isItalic = true;
          action.targetFormat.alignment = 'Justify';
          action.targetFormat.firstLineIndentCm = 1.0;
          break;

        case 'bodyText':
        case 'article':
        case 'clause':
        case 'point':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 14;
          action.targetFormat.isBold = (role === 'article');
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Justify';
          action.targetFormat.firstLineIndentCm = 1.0;
          action.targetFormat.spaceAfterPt = 6;
          if (regime === 'DANG_HD05') {
            action.targetFormat.lineSpacingExactPt = 20; // 18-22pt
          } else if (regime === 'VIETTEL') {
            action.targetFormat.lineSpacingRule = 'Single';
          } else {
            action.targetFormat.lineSpacingPt = 1.15; // 1.0 - 1.5 lines
          }
          break;

        case 'signAuthority':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 14;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          break;

        case 'signPosition':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 14;
          action.targetFormat.isBold = (regime !== 'DANG_HD05');
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          break;

        case 'signFullName':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 14;
          action.targetFormat.isBold = true;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Center';
          break;

        case 'recipientsNoiNhan':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 14 : 12;
          action.targetFormat.isBold = (regime !== 'DANG_HD05');
          action.targetFormat.isItalic = (regime !== 'DANG_HD05');
          action.targetFormat.isUnderline = (regime === 'DANG_HD05');
          action.targetFormat.alignment = 'Left';
          break;

        case 'recipientListItem':
        case 'recipientLuuLine':
          action.targetFormat.fontSize = regime === 'DANG_HD05' ? 12 : 11;
          action.targetFormat.isBold = false;
          action.targetFormat.isItalic = false;
          action.targetFormat.alignment = 'Left';
          break;
      }

      autoFixActions.push(action);
    }

    const errors = issues.filter(i => i.severity === 'ERROR').length;
    const warnings = issues.filter(i => i.severity === 'WARNING').length;
    const infos = issues.filter(i => i.severity === 'INFO').length;

    return {
      detectedRegime: regime,
      regimeName: rule.name || regime,
      totalParagraphs: snapshot.paragraphs.length,
      totalIssues: issues.length,
      errorCount: errors,
      warningCount: warnings,
      infoCount: infos,
      issues,
      autoFixActions,
      summary: {
        pageSetupValid: !issues.some(i => i.category === 'PAGE_SETUP' && i.severity === 'ERROR'),
        typographyScore: Math.max(0, 100 - (errors * 10 + warnings * 3)),
        structureScore: Math.max(0, 100 - (errors * 15)),
        isCompliant: errors === 0
      }
    };
  }
}
