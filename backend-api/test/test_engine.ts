import { DocumentSnapshot } from '../src/types';
import { AutoFixEngine } from '../src/autofix/AutoFixEngine';
import { TextNormalizer } from '../src/normalizer/TextNormalizer';
import { LicenseManager } from '../src/auth/LicenseManager';

console.log('=== TEST 1: TYPOGRAPHY NORMALIZATION ===');
const messyText = '  UBND   tỉnh    hoà  bình   đã   xử  lí   kĩ  thuật  cho   nhân dân  , ngày  05  tháng 1 năm 2026. ';
const cleanText = TextNormalizer.normalizeAll(messyText);
console.log('Original:', messyText);
console.log('Cleaned :', cleanText);

console.log('\n=== TEST 2: DOCUMENT SCAN & AUTO-FIX (ND30) ===');
const testDoc: DocumentSnapshot = {
  regime: 'ND30',
  pageSetup: {
    paperSize: 'Letter', // Error (must be A4)
    widthMm: 215.9,
    heightMm: 279.4,
    orientation: 'Portrait',
    marginsMm: {
      top: 15, // Error (< 20)
      bottom: 15,
      left: 20, // Error (< 30)
      right: 10 // Error (< 15)
    }
  },
  hasHeaderPageNumber: false,
  firstPageHasNumber: false,
  tableCount: 2,
  paragraphs: [
    {
      index: 0,
      text: 'ubnd tỉnh hà tĩnh',
      cleanText: 'ubnd tỉnh hà tĩnh',
      fontName: 'Arial', // Error (must be Times New Roman)
      fontSize: 12,
      isBold: false,
      alignment: 'Left'
    },
    {
      index: 1,
      text: 'CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM',
      cleanText: 'CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM',
      fontName: 'Arial',
      fontSize: 12,
      isBold: true,
      alignment: 'Left'
    },
    {
      index: 2,
      text: 'Độc lập-Tự do-Hạnh phúc',
      cleanText: 'Độc lập-Tự do-Hạnh phúc',
      fontName: 'Arial',
      fontSize: 12,
      isBold: true,
      alignment: 'Left'
    },
    {
      index: 3,
      text: 'QUYẾT ĐỊNH',
      cleanText: 'QUYẾT ĐỊNH',
      fontName: 'Arial',
      fontSize: 14,
      isBold: true,
      alignment: 'Center'
    },
    {
      index: 4,
      text: 'Về việc thành lập tổ công tác',
      cleanText: 'Về việc thành lập tổ công tác',
      fontName: 'Arial',
      fontSize: 14,
      isBold: true,
      alignment: 'Center'
    },
    {
      index: 5,
      text: 'Điều 1. Thành lập tổ công tác gồm các thành viên sau đây...',
      cleanText: 'Điều 1. Thành lập tổ công tác gồm các thành viên sau đây...',
      fontName: 'Arial',
      fontSize: 13,
      isBold: true,
      alignment: 'Left', // Error (must be Justify)
      firstLineIndentPt: 0 // Error (must be indent)
    }
  ]
};

const result = AutoFixEngine.analyzeAndPlanAutoFix(testDoc);
console.log(`Detected Regime: ${result.detectedRegime} (${result.regimeName})`);
console.log(`Total Issues Found: ${result.totalIssues} (Errors: ${result.errorCount}, Warnings: ${result.warningCount})`);
console.log(`AutoFix Actions Generated: ${result.autoFixActions.length}`);
for (const issue of result.issues) {
  console.log(` - [${issue.severity}] [${issue.checkCode}]: ${issue.title}`);
}

console.log('\n=== TEST 3: LICENSE KEY GENERATION ===');
const license = LicenseManager.generateLicenseKey('HWID-9876-5432-10', 'UBND Tinh Ha Tinh', 365);
console.log('Generated License Key:', license);
const isValid = LicenseManager.verifyLicense(license, 'HWID-9876-5432-10', 'UBND Tinh Ha Tinh', '2027-09-01');
console.log('License Valid:', isValid);

console.log('\nALL TESTS PASSED SUCCESSFULLY!');
