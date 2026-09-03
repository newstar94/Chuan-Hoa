import { ParagraphSnapshot, ComponentRole, RegimeType } from '../types';

export class ComponentDetector {
  /**
   * Tự động phát hiện Regime dựa vào nội dung tiêu đề và cơ quan
   */
  public static detectRegime(paragraphs: ParagraphSnapshot[]): RegimeType {
    const fullTextUpper = paragraphs.slice(0, 10).map(p => p.text.toUpperCase()).join(' ');

    if (fullTextUpper.includes('ĐẢNG CỘNG SẢN VIỆT NAM') || 
        fullTextUpper.includes('BAN CHẤP HÀNH TRUNG ƯƠNG') ||
        fullTextUpper.includes('TỈNH UỶ') || fullTextUpper.includes('THÀNH UỶ') ||
        fullTextUpper.includes('ĐẢNG BỘ') || fullTextUpper.includes('CHI BỘ')) {
      return 'DANG_HD05';
    }

    if (fullTextUpper.includes('VIỄN THÔNG QUÂN ĐỘI') || 
        fullTextUpper.includes('TẬP ĐOÀN CÔNG NGHIỆP - VIỄN THÔNG') ||
        fullTextUpper.includes('VIETTEL')) {
      return 'VIETTEL';
    }

    return 'ND30';
  }

  /**
   * Phân loại vai trò 12+ thành phần cho từng đoạn văn bản
   */
  public static classifyParagraphs(paragraphs: ParagraphSnapshot[], regime: RegimeType): ParagraphSnapshot[] {
    const total = paragraphs.length;
    let foundTypeNameOrSubject = false;
    let foundSignerZone = false;
    let inRecipientZone = false;

    for (let i = 0; i < total; i++) {
      const p = paragraphs[i];
      const text = p.text.trim();
      const upper = text.toUpperCase();

      if (!text) {
        p.detectedRole = 'unknown';
        continue;
      }

      // --- 1. ZONE ĐẦU VĂN BẢN (TRANG 1, KHOẢNG 10-15 ĐOẠN ĐẦU) ---
      if (i < 15 && !foundTypeNameOrSubject) {
        // Quốc hiệu (NĐ30, Viettel)
        if (upper.includes('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM') || upper.includes('CONG HOA XA HOI CHU NGHIA VIET NAM')) {
          p.detectedRole = 'nationalTitle';
          continue;
        }

        // Tiêu ngữ (NĐ30, Viettel)
        if (upper.includes('ĐỘC LẬP - TỰ DO - HẠNH PHÚC') || upper.includes('DOC LAP - TU DO - HANH PHUC') || upper.includes('ĐỘC LẬP-TỰ DO-HẠNH PHÚC')) {
          p.detectedRole = 'motto';
          continue;
        }

        // Tiêu đề Đảng (HD05)
        if (upper === 'ĐẢNG CỘNG SẢN VIỆT NAM' || upper === 'DANG CONG SAN VIET NAM') {
          p.detectedRole = 'partyTitle';
          continue;
        }

        // Số và Ký hiệu
        if (/^Số\s*:?\s*\d+/i.test(text) || /^So\s*:?\s*\d+/i.test(text)) {
          p.detectedRole = 'codeNumberNotation';
          continue;
        }

        // Địa danh, ngày tháng
        if (/,\s*ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+/i.test(text) || /,\s*ngay\s+\d+\s+thang\s+\d+\s+nam\s+\d+/i.test(text)) {
          p.detectedRole = 'placeDate';
          continue;
        }

        // Trích yếu công văn (V/v...)
        if (/^V\/v\s+/i.test(text) || /^Về\s+việc\s+/i.test(text)) {
          p.detectedRole = 'officialLetterSubject';
          foundTypeNameOrSubject = true;
          continue;
        }

        // Tên cơ quan cấp trên / cơ quan ban hành (ở đầu văn bản)
        if (i < 6 && (p.isBold || upper === text)) {
          if (upper.includes('BỘ ') || upper.includes('ỦY BAN NHÂN DÂN') || upper.includes('TỈNH') || upper.includes('THÀNH PHỐ') || upper.includes('CỤC ') || upper.includes('SỞ ') || upper.includes('TẬP ĐOÀN')) {
            p.detectedRole = 'organName';
            continue;
          }
        }

        // Tên loại văn bản (QUYẾT ĐỊNH, CHỈ THỊ, THÔNG TƯ, NGHỊ QUYẾT, BÁO CÁO, KẾ HOẠCH, QUY ĐỊNH, HƯỚNG DẪN, TỜ TRÌNH...)
        if (/^(QUYẾT ĐỊNH|CHỈ THỊ|THÔNG TƯ|NGHỊ QUYẾT|BÁO CÁO|KẾ HOẠCH|QUY ĐỊNH|QUY CHẾ|HƯỚNG DẪN|TỜ TRÌNH|THÔNG BÁO|CÔNG VĂN|BIÊN BẢN)$/i.test(upper) ||
            /^(QUYET DINH|CHI THI|THONG TU|NGHI QUYET|BAO CAO|KE HOACH|QUY DINH|QUY CHE|HUONG DAN|TO TRINH|THONG BAO|BIEN BAN)$/i.test(upper)) {
          p.detectedRole = 'typeName';
          foundTypeNameOrSubject = true;
          // Đoạn tiếp theo thường là trích yếu
          if (i + 1 < total && paragraphs[i + 1].text.trim()) {
            paragraphs[i + 1].detectedRole = 'subject';
          }
          continue;
        }
      }

      // --- 2. CĂN CỨ PHÁP LÝ ---
      if (/^Căn\s+cứ\s+/i.test(text) || /^Can\s+cu\s+/i.test(text) || /^-\s*Căn\s+cứ\s+/i.test(text)) {
        p.detectedRole = 'legalBases';
        continue;
      }

      // --- 3. KÍNH GỬI / KÍNH TRÌNH ---
      if (/^Kính\s+gửi\s*:/i.test(text) || /^Kính\s+trình\s*:/i.test(text) || /^Kinh\s+gui\s*:/i.test(text)) {
        p.detectedRole = 'recipientsKinhGui';
        continue;
      }

      // --- 4. BỐ CỤC: PHẦN, CHƯƠNG, MỤC, TIỂU MỤC, ĐIỀU, KHOẢN, ĐIỂM ---
      if (/^(Phần|Phan)\s+([IVXLCDM\d]+|thứ\s+[a-z]+)/i.test(text) || /^(Chương|Chuong)\s+([IVXLCDM\d]+|thứ\s+[a-z]+)/i.test(text)) {
        p.detectedRole = 'partChapterTitle';
        continue;
      }

      if (/^(Mục|Muc)\s+\d+/i.test(text) || /^(Tiểu\s+mục|Tieu\s+muc)\s+\d+/i.test(text)) {
        p.detectedRole = 'sectionTitle';
        continue;
      }

      if (/^Điều\s+\d+\.?/i.test(text) || /^Dieu\s+\d+\.?/i.test(text)) {
        p.detectedRole = 'article';
        continue;
      }

      if (/^\d+\.\s+[^\d]/.test(text)) {
        p.detectedRole = 'clause';
        continue;
      }

      if (/^[a-zđ]\)\s+/i.test(text)) {
        p.detectedRole = 'point';
        continue;
      }

      // --- 5. ZONE CUỐI VĂN BẢN: QUYỀN HẠN, CHỨC VỤ, CHỮ KÝ, NƠI NHẬN ---
      if (i > total - 25) {
        // Nơi nhận
        if (/^Nơi\s+nhận\s*:/i.test(text) || /^Noi\s+nhan\s*:/i.test(text)) {
          p.detectedRole = 'recipientsNoiNhan';
          inRecipientZone = true;
          continue;
        }

        if (inRecipientZone) {
          if (/^-\s*Lưu\s*:/i.test(text) || /^-\s*Luu\s*:/i.test(text) || /Lưu\s+Văn\s+phòng/i.test(text)) {
            p.detectedRole = 'recipientLuuLine';
            inRecipientZone = false;
            continue;
          }
          if (/^-\s*/.test(text)) {
            p.detectedRole = 'recipientListItem';
            continue;
          }
        }

        // Quyền hạn ký (TM., KT., TL., TUQ., Q. / T/M, K/T, T/L)
        if (/^(TM\.|KT\.|TL\.|TUQ\.|Q\.|T\/M|K\/T|T\/L)\s+/i.test(upper) || /^(TM\.|KT\.|TL\.|TUQ\.|Q\.|T\/M|K\/T|T\/L)$/i.test(upper)) {
          p.detectedRole = 'signAuthority';
          foundSignerZone = true;
          continue;
        }

        // Chức vụ người ký
        if (/^(CHỦ TỊCH|PHÓ CHỦ TỊCH|GIÁM ĐỐC|PHÓ GIÁM ĐỐC|BÍ THƯ|PHÓ BÍ THƯ|BỘ TRƯỞNG|THỨ TRƯỞNG|CHÁNH VĂN PHÒNG|TRƯỞNG BAN)$/i.test(upper)) {
          p.detectedRole = 'signPosition';
          foundSignerZone = true;
          continue;
        }

        // Họ và tên người ký (Title case, in đậm ở gần cuối)
        if (foundSignerZone && i >= total - 5 && p.isBold) {
          p.detectedRole = 'signFullName';
          continue;
        }
      }

      // Mặc định là nội dung văn bản chính
      if (!p.detectedRole) {
        p.detectedRole = 'bodyText';
      }
    }

    return paragraphs;
  }
}
