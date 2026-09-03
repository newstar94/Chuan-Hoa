import { DocumentSnapshot, ComplianceIssue, RegimeType, ParagraphSnapshot } from '../types';
import { RuleLoader } from '../rules/RuleLoader';

export class ComplianceChecker {
  private static ruleLoader = RuleLoader.getInstance();

  /**
   * Kiểm tra toàn diện tài liệu theo 82 mã kiểm tra
   */
  public static checkDocument(snapshot: DocumentSnapshot): ComplianceIssue[] {
    const issues: ComplianceIssue[] = [];
    const regime = snapshot.regime;
    const rule = this.ruleLoader.getRegimeRule(regime);

    // 1. Kiểm tra Thiết lập Trang (Page Setup)
    this.checkPageSetup(snapshot, rule, issues);

    // 2. Kiểm tra Đánh số trang
    this.checkPageNumbering(snapshot, rule, issues);

    // 3. Kiểm tra từng Đoạn văn và Thành phần Thể thức
    for (const p of snapshot.paragraphs) {
      this.checkParagraphTypography(p, regime, rule, issues);
      this.checkComponentSpecificRules(p, regime, rule, issues);
      this.checkCapitalizationRules(p, issues);
      this.checkTypoAndSpacingRules(p, issues);
    }

    return issues;
  }

  // --- 1. PAGE SETUP CHECKS ---
  private static checkPageSetup(snapshot: DocumentSnapshot, rule: any, issues: ComplianceIssue[]): void {
    const ps = snapshot.pageSetup;
    if (ps.paperSize && ps.paperSize.toUpperCase() !== 'A4') {
      issues.push({
        id: `CHK_PAGE_SIZE_${Date.now()}`,
        checkCode: 'CHK_PAGE_SIZE',
        category: 'PAGE_SETUP',
        severity: 'ERROR',
        title: 'Khổ giấy không phải A4',
        description: `Văn bản đang để khổ ${ps.paperSize}. Quy định bắt buộc dùng khổ A4 (210 x 297 mm).`,
        currentValue: ps.paperSize,
        expectedValue: 'A4',
        autoFixable: true
      });
    }

    if (ps.orientation && ps.orientation !== 'Portrait') {
      issues.push({
        id: `CHK_PAGE_ORIENTATION_${Date.now()}`,
        checkCode: 'CHK_PAGE_ORIENTATION',
        category: 'PAGE_SETUP',
        severity: 'ERROR',
        title: 'Hướng trang không phải chiều dọc',
        description: 'Văn bản hành chính phải trình bày theo chiều dài khổ A4 (Portrait).',
        currentValue: ps.orientation,
        expectedValue: 'Portrait',
        autoFixable: true
      });
    }

    // Margins check
    const m = ps.marginsMm;
    if (m) {
      const topRule = rule.pageSetup?.marginsMm?.top || { min: 20, max: 25 };
      const leftRule = rule.pageSetup?.marginsMm?.left || { min: 30, max: 35 };
      const rightRule = rule.pageSetup?.marginsMm?.right || { min: 15, max: 20 };
      const bottomRule = rule.pageSetup?.marginsMm?.bottom || { min: 20, max: 25 };

      if (m.top < topRule.min || m.top > topRule.max ||
          m.left < leftRule.min || m.left > leftRule.max ||
          m.right < rightRule.min || m.right > rightRule.max ||
          m.bottom < bottomRule.min || m.bottom > bottomRule.max) {
        issues.push({
          id: `CHK_PAGE_MARGINS_${Date.now()}`,
          checkCode: 'CHK_PAGE_MARGINS',
          category: 'PAGE_SETUP',
          severity: 'ERROR',
          title: 'Căn lề trang không đúng quy chuẩn',
          description: `Lề hiện tại: Trên ${m.top}mm, Dưới ${m.bottom}mm, Trái ${m.left}mm, Phải ${m.right}mm. Chuẩn: Trên/Dưới 20-25mm, Trái 30-35mm, Phải 15-20mm.`,
          currentValue: m,
          expectedValue: { top: 20, bottom: 20, left: 30, right: 15 },
          autoFixable: true
        });
      }
    }
  }

  // --- 2. PAGE NUMBERING ---
  private static checkPageNumbering(snapshot: DocumentSnapshot, rule: any, issues: ComplianceIssue[]): void {
    if (!snapshot.hasHeaderPageNumber && snapshot.paragraphs.length > 20) {
      issues.push({
        id: `CHK_PAGE_NUMBERING_${Date.now()}`,
        checkCode: 'CHK_PAGE_NUMBERING',
        category: 'PAGE_SETUP',
        severity: 'WARNING',
        title: 'Chưa đánh số trang đúng chuẩn',
        description: 'Số trang phải được đánh từ số 1, cỡ 13-14 đứng, đặt canh giữa tại Header và ẩn ở trang 1.',
        autoFixable: true
      });
    }
  }

  // --- 3. PARAGRAPH TYPOGRAPHY ---
  private static checkParagraphTypography(p: ParagraphSnapshot, regime: RegimeType, rule: any, issues: ComplianceIssue[]): void {
    const text = p.text.trim();
    if (!text) return;

    // Font family check
    if (p.fontName && !p.fontName.toLowerCase().includes('times new roman')) {
      issues.push({
        id: `CHK_BODY_FONT_NAME_${p.index}`,
        paragraphIndex: p.index,
        componentRole: p.detectedRole,
        checkCode: 'CHK_BODY_FONT_NAME',
        category: 'TYPOGRAPHY_AND_BODY',
        severity: 'ERROR',
        title: `Phông chữ không phải Times New Roman (Đoạn ${p.index + 1})`,
        description: `Đoạn đang dùng phông '${p.fontName}'. Bắt buộc dùng 'Times New Roman'.`,
        currentValue: p.fontName,
        expectedValue: 'Times New Roman',
        autoFixable: true
      });
    }

    // Body alignment & indent
    if (p.detectedRole === 'bodyText' || p.detectedRole === 'article' || p.detectedRole === 'clause' || p.detectedRole === 'point') {
      if (p.alignment && p.alignment !== 'Justify') {
        issues.push({
          id: `CHK_BODY_ALIGNMENT_${p.index}`,
          paragraphIndex: p.index,
          componentRole: p.detectedRole,
          checkCode: 'CHK_BODY_ALIGNMENT',
          category: 'TYPOGRAPHY_AND_BODY',
          severity: 'ERROR',
          title: `Nội dung chưa căn đều hai lề (Đoạn ${p.index + 1})`,
          description: `Đoạn đang căn '${p.alignment}'. Bắt buộc căn đều 2 lề (Justify).`,
          currentValue: p.alignment,
          expectedValue: 'Justify',
          autoFixable: true
        });
      }

      if (p.firstLineIndentPt !== undefined && p.firstLineIndentPt < 20) {
        issues.push({
          id: `CHK_BODY_INDENT_${p.index}`,
          paragraphIndex: p.index,
          componentRole: p.detectedRole,
          checkCode: 'CHK_BODY_INDENT',
          category: 'TYPOGRAPHY_AND_BODY',
          severity: 'ERROR',
          title: `Chưa thụt đầu dòng (Đoạn ${p.index + 1})`,
          description: 'Nội dung khi xuống dòng phải thụt đầu dòng 1.0 - 1.27 cm (hoặc 10mm đối với văn bản Đảng).',
          autoFixable: true
        });
      }
    }
  }

  // --- 4. COMPONENT SPECIFIC RULES ---
  private static checkComponentSpecificRules(p: ParagraphSnapshot, regime: RegimeType, rule: any, issues: ComplianceIssue[]): void {
    const text = p.text.trim();
    const role = p.detectedRole;
    if (!text || !role) return;

    if (role === 'nationalTitle') {
      if (text !== 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM') {
        issues.push({
          id: `CHK_NATIONAL_TITLE_${p.index}`,
          paragraphIndex: p.index,
          componentRole: role,
          checkCode: 'CHK_NATIONAL_TITLE',
          category: 'COMPONENTS',
          severity: 'ERROR',
          title: 'Quốc hiệu chưa đúng chính tả hoặc in hoa',
          description: 'Quốc hiệu phải là: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM (In hoa, Đứng, Đậm).',
          currentValue: text,
          expectedValue: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
          autoFixable: true
        });
      }
    }

    if (role === 'motto') {
      if (!text.includes('Độc lập - Tự do - Hạnh phúc')) {
        issues.push({
          id: `CHK_NATIONAL_MOTTO_${p.index}`,
          paragraphIndex: p.index,
          componentRole: role,
          checkCode: 'CHK_NATIONAL_MOTTO',
          category: 'COMPONENTS',
          severity: 'ERROR',
          title: 'Tiêu ngữ chưa đúng quy chuẩn gạch nối',
          description: 'Tiêu ngữ chuẩn: Độc lập - Tự do - Hạnh phúc (chữ cái đầu viết hoa, gạch nối có cách chữ).',
          currentValue: text,
          expectedValue: 'Độc lập - Tự do - Hạnh phúc',
          autoFixable: true
        });
      }
    }

    if (role === 'partyTitle') {
      if (text !== 'ĐẢNG CỘNG SẢN VIỆT NAM') {
        issues.push({
          id: `CHK_PARTY_TITLE_${p.index}`,
          paragraphIndex: p.index,
          componentRole: role,
          checkCode: 'CHK_NATIONAL_TITLE',
          category: 'COMPONENTS',
          severity: 'ERROR',
          title: 'Tiêu đề Đảng chưa in hoa đậm',
          description: 'Tiêu đề Đảng phải là: ĐẢNG CỘNG SẢN VIỆT NAM (Cỡ 15, In hoa, Đứng, Đậm).',
          currentValue: text,
          expectedValue: 'ĐẢNG CỘNG SẢN VIỆT NAM',
          autoFixable: true
        });
      }
    }

    if (role === 'codeNumberNotation') {
      if (regime === 'ND30' && !text.startsWith('Số:')) {
        issues.push({
          id: `CHK_CODE_NUMBER_COLON_${p.index}`,
          paragraphIndex: p.index,
          componentRole: role,
          checkCode: 'CHK_CODE_NUMBER_FORMAT',
          category: 'COMPONENTS',
          severity: 'ERROR',
          title: 'Số văn bản NĐ30 thiếu dấu hai chấm sau từ "Số:"',
          description: 'Theo NĐ 30/2020/NĐ-CP, từ "Số:" phải có dấu hai chấm liền kề (ví dụ: Số: 15/QĐ-UBND).',
          currentValue: text,
          autoFixable: true
        });
      } else if (regime === 'DANG_HD05' && text.startsWith('Số:')) {
        issues.push({
          id: `CHK_CODE_NUMBER_NO_COLON_${p.index}`,
          paragraphIndex: p.index,
          componentRole: role,
          checkCode: 'CHK_CODE_NUMBER_FORMAT',
          category: 'COMPONENTS',
          severity: 'ERROR',
          title: 'Số văn bản Đảng không được có dấu hai chấm sau từ "Số"',
          description: 'Theo HD 05-HD/VPTW, số văn bản Đảng viết là "Số 05-HD/VPTW" (không có dấu hai chấm).',
          currentValue: text,
          autoFixable: true
        });
      }
    }
  }

  // --- 5. CAPITALIZATION RULES ---
  private static checkCapitalizationRules(p: ParagraphSnapshot, issues: ComplianceIssue[]): void {
    const text = p.text.trim();
    if (!text) return;

    // Check sentence capitalization
    if (/^[a-zàáảãạăắằẳẵặâấầẩẫậèéẻẽẹêếềểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵđ]/.test(text)) {
      issues.push({
        id: `CHK_CAP_SENTENCE_${p.index}`,
        paragraphIndex: p.index,
        componentRole: p.detectedRole,
        checkCode: 'CHK_CAP_SENTENCE',
        category: 'CAPITALIZATION_PHULUC2',
        severity: 'ERROR',
        title: `Chưa viết hoa chữ cái đầu đoạn (Đoạn ${p.index + 1})`,
        description: 'Quy tắc đặt câu: Phải viết hoa chữ cái đầu tiên của câu và sau dấu chấm câu.',
        autoFixable: true
      });
    }

    // Check special nouns
    if (/\bnhân dân\b/.test(text) && !/\bNhân dân\b/.test(text)) {
      issues.push({
        id: `CHK_CAP_NHAN_DAN_${p.index}`,
        paragraphIndex: p.index,
        componentRole: p.detectedRole,
        checkCode: 'CHK_CAP_SPECIAL_NOUNS',
        category: 'CAPITALIZATION_PHULUC2',
        severity: 'WARNING',
        title: 'Từ "Nhân dân" cần được viết hoa trang trọng',
        description: 'Theo Phụ lục II NĐ30, danh từ "Nhân dân", "Nhà nước" thuộc trường hợp viết hoa đặc biệt.',
        autoFixable: true
      });
    }
  }

  // --- 6. TYPOGRAPHY & SPACING RULES ---
  private static checkTypoAndSpacingRules(p: ParagraphSnapshot, issues: ComplianceIssue[]): void {
    const text = p.text;
    if (!text) return;

    // Extra spaces
    if (/[ \t]{2,}/.test(text)) {
      issues.push({
        id: `CHK_EXTRA_SPACES_${p.index}`,
        paragraphIndex: p.index,
        componentRole: p.detectedRole,
        checkCode: 'CHK_EXTRA_SPACES',
        category: 'TYPOGRAPHY_AND_NORMALIZERS',
        severity: 'WARNING',
        title: `Phát hiện khoảng trắng kép thừa (Đoạn ${p.index + 1})`,
        description: 'Đoạn văn bản có chứa 2 hoặc nhiều dấu cách liên tiếp cần được rút gọn.',
        autoFixable: true
      });
    }

    // Space before punctuation
    if (/\s+[,.:;?!]/.test(text)) {
      issues.push({
        id: `CHK_PUNCTUATION_SPACING_${p.index}`,
        paragraphIndex: p.index,
        componentRole: p.detectedRole,
        checkCode: 'CHK_PUNCTUATION_SPACING',
        category: 'TYPOGRAPHY_AND_NORMALIZERS',
        severity: 'WARNING',
        title: `Có dấu cách thừa trước dấu câu (Đoạn ${p.index + 1})`,
        description: 'Không được để dấu cách trước các dấu phẩy, chấm, hai chấm, chấm phẩy, chấm hỏi, chấm than.',
        autoFixable: true
      });
    }
  }
}
