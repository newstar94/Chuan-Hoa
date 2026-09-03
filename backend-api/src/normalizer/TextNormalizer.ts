import { RuleLoader } from '../rules/RuleLoader';

export class TextNormalizer {
  private static ruleLoader = RuleLoader.getInstance();

  // 1. Dấu thanh chuẩn mới vs chuẩn cũ (oà/hòa, uý/thúy)
  private static newToneMap: Record<string, string> = {
    'hoà': 'hòa', 'hoá': 'hóa', 'hoả': 'hỏa', 'hoã': 'hõa', 'hoạ': 'họa',
    'toà': 'tòa', 'toán': 'toán', 'toả': 'tỏa', 'toã': 'tõa', 'toạ': 'tọa',
    'xoà': 'xòa', 'xoá': 'xóa', 'xoả': 'xỏa', 'xoã': 'xõa', 'xoạ': 'xọa',
    'doà': 'dòa', 'doá': 'dóa', 'doả': 'dỏa', 'doã': 'dõa', 'doạ': 'dọa',
    'loà': 'lòa', 'loá': 'lóa', 'loả': 'lỏa', 'loã': 'lõa', 'loạ': 'lọa',
    'thuý': 'thúy', 'thuỷ': 'thủy', 'thuỵ': 'thụy', 'quyế': 'quế',
    'Thuý': 'Thúy', 'Thuỷ': 'Thủy', 'Thuỵ': 'Thụy',
    'Hoà': 'Hòa', 'Hoá': 'Hóa', 'Hoả': 'Hỏa', 'Hoã': 'Hõa', 'Hoạ': 'Họa',
    'Toà': 'Tòa', 'Toán': 'Toán', 'Toả': 'Tỏa', 'Toã': 'Tõa', 'Toạ': 'Tọa'
  };

  /**
   * Chuẩn hóa dấu thanh kiểu mới (hòa, thúy)
   */
  public static normalizeTone(text: string): string {
    let result = text;
    for (const [oldTone, newTone] of Object.entries(this.newToneMap)) {
      const regex = new RegExp(`\\b${oldTone}\\b`, 'g');
      result = result.replace(regex, newTone);
    }
    return result;
  }

  /**
   * Chuẩn hóa i / y theo từ điển chuẩn
   */
  public static normalizeIy(text: string): string {
    const iyDict = this.ruleLoader.iyDict;
    let result = text;
    for (const [wrong, right] of Object.entries(iyDict)) {
      const regex = new RegExp(`\\b${wrong}\\b`, 'gi');
      result = result.replace(regex, (match) => {
        // preserve case
        if (match === match.toUpperCase()) return right.toUpperCase();
        if (match[0] === match[0].toUpperCase()) return right.charAt(0).toUpperCase() + right.slice(1);
        return right;
      });
    }
    return result;
  }

  /**
   * Sửa các lỗi chính tả phổ biến trong từ điển
   */
  public static fixTypos(text: string): string {
    const typoDict = this.ruleLoader.typoDict;
    let result = text;
    for (const [wrong, right] of Object.entries(typoDict)) {
      const regex = new RegExp(`\\b${wrong}\\b`, 'gi');
      result = result.replace(regex, (match) => {
        if (match === match.toUpperCase()) return right.toUpperCase();
        if (match[0] === match[0].toUpperCase()) return right.charAt(0).toUpperCase() + right.slice(1);
        return right;
      });
    }
    return result;
  }

  /**
   * Dọn sạch khoảng trắng thừa và chuẩn hóa khoảng cách dấu câu
   */
  public static normalizeSpacesAndPunctuation(text: string): string {
    let result = text;
    // 1. Khử ký tự ẩn / non-breaking space (U+00A0, U+FEFF, U+200B)
    result = result.replace(/[\u00A0\uFEFF\u200B\u200E\u200F]/g, ' ');

    // 2. Khử nhiều dấu cách liên tiếp
    result = result.replace(/[ \t]+/g, ' ');

    // 3. Xóa khoảng cách trước các dấu câu (. , : ; ? ! ) ] } )
    result = result.replace(/\s+([.,:;?!)\\]}])/g, '$1');

    // 4. Thêm khoảng cách sau dấu câu (. , : ; ? ! ) nếu dính liền chữ cái
    result = result.replace(/([.,:;?!])([^\s\d.,:;?!_\\/"'])/g, '$1 $2');

    // 5. Chuẩn hóa dấu mở ngoặc ( [ { không có khoảng trắng sau
    result = result.replace(/([(\\[{])\s+/g, '$1');

    // 6. Xóa khoảng trắng đầu & cuối dòng
    result = result.trim();

    return result;
  }

  /**
   * Chuẩn hóa các loại dấu gạch ngang (hyphen, en-dash, em-dash)
   */
  public static normalizeDashes(text: string): string {
    let result = text;
    // Tiêu ngữ: Độc lập - Tự do - Hạnh phúc (gạch nối có khoảng trắng hai bên)
    result = result.replace(/Độc\s*lập\s*[-–—]\s*Tự\s*do\s*[-–—]\s*Hạnh\s*phúc/gi, 'Độc lập - Tự do - Hạnh phúc');
    return result;
  }

  /**
   * Chuẩn hóa dấu ba chấm (3 dots ...) thay vì ký tự ellipsis đơn (U+2026 …)
   */
  public static normalizeEllipsis(text: string): string {
    return text.replace(/\u2026/g, '...').replace(/\.{4,}/g, '...');
  }

  /**
   * Chuẩn hóa toàn diện 1 chuỗi text (Tổng hợp)
   */
  public static normalizeAll(text: string): string {
    let s = text;
    s = this.normalizeSpacesAndPunctuation(s);
    s = this.normalizeTone(s);
    s = this.normalizeIy(s);
    s = this.fixTypos(s);
    s = this.normalizeDashes(s);
    s = this.normalizeEllipsis(s);
    return s;
  }
}
