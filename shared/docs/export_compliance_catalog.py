import os
import json

base_dir = r'd:\chuan-hoa-the-thuc-workspace'
rules_dir = os.path.join(base_dir, 'shared', 'rules')

# 1. Viettel Regime Rules
rules_viettel = {
  "regime": "VIETTEL_QD11095",
  "name": "Quy chế Thể thức Văn bản Tập đoàn Viettel (QĐ 11095)",
  "documentScope": "Văn bản nội bộ và giao dịch của Tập đoàn Công nghiệp - Viễn thông Quân đội (Viettel)",
  "pageSetup": {
    "paperSize": "A4",
    "widthMm": 210,
    "heightMm": 297,
    "orientation": "Portrait",
    "marginsMm": {
      "top": { "min": 20, "max": 25, "recommended": 20 },
      "bottom": { "min": 20, "max": 25, "recommended": 20 },
      "left": { "min": 30, "max": 35, "recommended": 30 },
      "right": { "min": 15, "max": 20, "recommended": 15 }
    }
  },
  "defaultTypography": {
    "fontFamily": "Times New Roman",
    "color": "#000000",
    "encoding": "Unicode TCVN 6909:2001 (NFC)"
  },
  "lineSpacing": {
    "bodyZoneRule": "Single",
    "spacingAfterPt": 6
  },
  "components": {
    "nationalTitle": {
      "text": "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
      "fontSize": 12,
      "fontWeight": "Bold",
      "alignment": "Center"
    },
    "nationalMotto": {
      "text": "Độc lập - Tự do - Hạnh phúc",
      "fontSize": 13,
      "fontWeight": "Bold",
      "alignment": "Center",
      "borderBottom": {
        "required": True,
        "lineType": "SolidSingle",
        "lengthRatio": 1.0
      }
    },
    "parentOrganName": {
      "defaultText": "TẬP ĐOÀN CÔNG NGHIỆP - VIỄN THÔNG QUÂN ĐỘI",
      "fontSize": 12,
      "fontWeight": "Normal",
      "alignment": "Center"
    },
    "organName": {
      "fontSize": 12,
      "fontWeight": "Bold",
      "alignment": "Center",
      "borderBottom": {
        "required": True,
        "lengthRatioMin": 0.333,
        "lengthRatioMax": 0.5
      }
    },
    "placeDate": {
      "fontSize": 13,
      "fontStyle": "Italic",
      "alignment": "Center"
    },
    "officialLetterSubject": {
      "prefix": "V/v",
      "fontSize": 12,
      "fontWeight": "Normal",
      "alignment": "Center",
      "spaceBeforePt": 6
    },
    "recipientList": {
      "archivePattern": "^-?\\s*Lưu\\s*:\\s*VT\\s*,\\s*[^\\s.]+\\.[^.]+\\.$",
      "example": "- Lưu: VT, TCHC.02."
    },
    "legalBasis": {
      "decree": {
        "bold": False,
        "italic": True,
        "underline": False
      },
      "other": {
        "bold": False,
        "italic": False,
        "underline": False,
        "lastLineEndChar": ","
      }
    }
  }
}

with open(os.path.join(rules_dir, 'rules_viettel.json'), 'w', encoding='utf-8') as f:
    json.dump(rules_viettel, f, ensure_ascii=False, indent=2)

# 2. Compliance Checks Catalog
checks_catalog = {
  "totalChecks": 82,
  "categories": [
    {
      "category": "PAGE_SETUP",
      "name": "Thiết lập Trang & Định dạng Khổ giấy",
      "checks": [
        { "code": "CHK_PAGE_SIZE", "name": "Kiểm tra khổ giấy A4 (210x297mm)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_PAGE_ORIENTATION", "name": "Kiểm tra hướng giấy dọc (Portrait)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_PAGE_MARGINS", "name": "Kiểm tra căn lề trang (Trên/Dưới 20-25mm, Trái 30-35mm, Phải 15-20mm)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_PAGE_NUMBERING", "name": "Kiểm tra đánh số trang Header giữa, cỡ 13-14 đứng, ẩn trang 1", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_APPENDIX_PAGE_NUM", "name": "Kiểm tra đánh số trang phụ lục tách biệt", "severity": "INFO", "autoFix": True }
      ]
    },
    {
      "category": "TYPOGRAPHY_AND_BODY",
      "name": "Kiểu chữ, Đoạn văn & Thụt lề Nội dung",
      "checks": [
        { "code": "CHK_BODY_FONT_NAME", "name": "Kiểm tra font chữ Times New Roman", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_BODY_FONT_COLOR", "name": "Kiểm tra màu chữ đen (Auto/Black)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_BODY_ALIGNMENT", "name": "Kiểm tra căn lề 2 bên (Justified)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_BODY_INDENT", "name": "Kiểm tra thụt đầu dòng (1.0 - 1.27 cm)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_BODY_SPACE_AFTER", "name": "Kiểm tra khoảng cách đoạn (>= 6pt)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_BODY_LINE_SPACING", "name": "Kiểm tra giãn dòng (NĐ30: 1.0-1.5 lines; Đảng: Exactly 18-22pt)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_BODY_END_DOT", "name": "Kiểm tra kết thúc nội dung có dấu chấm (.)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_FONT_SIZE_CONSISTENCY", "name": "Kiểm tra tính nhất quán cỡ chữ theo bộ quy tắc", "severity": "ERROR", "autoFix": True }
      ]
    },
    {
      "category": "COMPONENTS",
      "name": "12+ Thành phần Thể thức Văn bản",
      "checks": [
        { "code": "CHK_NATIONAL_TITLE", "name": "Quốc hiệu / Tiêu đề Đảng", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_NATIONAL_MOTTO", "name": "Tiêu ngữ (Độc lập - Tự do - Hạnh phúc)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_UNDERLINE_LENGTH", "name": "Độ dài đường kẻ dưới Tiêu ngữ (100%), Cơ quan (1/3-1/2), Trích yếu (1/3-1/2)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_ORGAN_NAMES", "name": "Tên cơ quan cấp trên & cơ quan ban hành", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_CODE_NUMBER_FORMAT", "name": "Cú pháp Số & Ký hiệu (Số: ... / Số ...)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_CODE_NUMBER_PAD", "name": "Số < 10 phải có số 0 đứng trước (01, 02...)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_PLACE_DATE", "name": "Địa danh và ngày, tháng, năm ban hành", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_TYPE_NAME_SUBJECT", "name": "Tên loại văn bản & Trích yếu nội dung", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_OFFICIAL_LETTER_SUBJECT", "name": "Trích yếu công văn (V/v... / về việc...)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_LEGAL_BASES", "name": "Căn cứ pháp lý (in nghiêng, ngắt dòng, dấu chấm phẩy)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_SIGN_AUTHORITY", "name": "Quyền hạn ký (TM., KT., TL., TUQ., Q. / T/M, K/T, T/L)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_SIGN_POSITION", "name": "Chức vụ người ký (In hoa, đậm NĐ30 / thường Đảng)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_SIGN_FULLNAME", "name": "Họ và tên người ký (Title Case, in đậm)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_RECIPIENTS_KINH_GUI", "name": "Kính gửi / Kính trình", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_RECIPIENTS_NOI_NHAN", "name": "Nơi nhận (NĐ30: nghiêng đậm / Đảng: đứng gạch chân)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_RECIPIENT_LUU_LINE", "name": "Dòng lưu văn bản (- Lưu: VT, ...)", "severity": "WARNING", "autoFix": True }
      ]
    },
    {
      "category": "STRUCTURE_HIERARCHY",
      "name": "Bố cục: Phần, Chương, Mục, Điều, Khoản, Điểm",
      "checks": [
        { "code": "CHK_PART_CHAPTER", "name": "Phần / Chương (Chữ La Mã, in đậm, căn giữa)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_SECTION", "name": "Mục / Tiểu mục (Số Ả Rập, in đậm, căn giữa)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_ARTICLE", "name": "Điều (Điều 1., in đậm, thụt đầu dòng)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_CLAUSE", "name": "Khoản (1., 2., 3.)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_POINT", "name": "Điểm (a), b), c)...)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_POINT_ORDER", "name": "Thứ tự điểm theo bảng chữ cái tiếng Việt", "severity": "WARNING", "autoFix": False },
        { "code": "CHK_LEGAL_CITATION", "name": "Định dạng viện dẫn văn bản quy phạm pháp luật", "severity": "WARNING", "autoFix": False }
      ]
    },
    {
      "category": "CAPITALIZATION_PHULUC2",
      "name": "Quy tắc Viết hoa (Phụ lục II Nghị định 30)",
      "checks": [
        { "code": "CHK_CAP_SENTENCE", "name": "Viết hoa đầu câu và sau dấu chấm câu", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_CAP_PERSON_NAME", "name": "Viết hoa tên người Việt Nam & nước ngoài", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_CAP_ADMIN_UNIT", "name": "Viết hoa đơn vị hành chính (tỉnh, huyện, xã, Quận 1...)", "severity": "ERROR", "autoFix": True },
        { "code": "CHK_CAP_GEOGRAPHY", "name": "Viết hoa tên địa lý, sông núi vùng miền (Tây Bắc, Bắc Bộ...)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_CAP_ORGANIZATION", "name": "Viết hoa tên cơ quan, tổ chức, chức năng", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_CAP_SPECIAL_NOUNS", "name": "Viết hoa danh từ đặc biệt (Nhân dân, Nhà nước, Đảng, Bác...)", "severity": "ERROR", "autoFix": True }
      ]
    },
    {
      "category": "TYPOGRAPHY_AND_NORMALIZERS",
      "name": "Chuẩn hóa Tiếng Việt, Chính tả & Ký tự",
      "checks": [
        { "code": "CHK_TONE_MARK_MIX", "name": "Kiểm tra đồng nhất vị trí dấu thanh tiếng Việt", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_IY_MIX", "name": "Kiểm tra đồng nhất chính tả i/y", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_EXTRA_SPACES", "name": "Loại bỏ khoảng trắng kép, khoảng trắng trước dấu câu", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_PUNCTUATION_SPACING", "name": "Chuẩn hóa khoảng trắng sau dấu phẩy, chấm, chấm phẩy", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_DASH_NORMALIZATION", "name": "Chuẩn hóa các loại dấu gạch ngang (hyphen, en-dash, em-dash)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_DECIMAL_SEPARATOR", "name": "Chuẩn hóa dấu phẩy thập phân và dấu chấm hàng nghìn", "severity": "INFO", "autoFix": True },
        { "code": "CHK_SOFT_LINEBREAKS", "name": "Chuẩn hóa ngắt dòng mềm (Shift+Enter sang Enter chuẩn)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_TRAILING_EMPTY_PAGES", "name": "Xóa trang trắng thừa ở cuối tài liệu", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_TABLE_HEADER_REPEAT", "name": "Tự động lặp lại Header hàng đầu của bảng khi tràn trang (RepeatHeaderRow)", "severity": "WARNING", "autoFix": True },
        { "code": "CHK_TABLE_ROW_SPLIT", "name": "Ngăn gãy dòng bảng giữa 2 trang (CantSplit)", "severity": "WARNING", "autoFix": True }
      ]
    }
  ]
}

with open(os.path.join(rules_dir, 'rules_compliance_checks.json'), 'w', encoding='utf-8') as f:
    json.dump(checks_catalog, f, ensure_ascii=False, indent=2)

print("Viettel rules and Compliance check catalog exported successfully!")
