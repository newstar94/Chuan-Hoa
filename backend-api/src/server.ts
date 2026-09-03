import express, { Request, Response } from 'express';
import cors from 'cors';
import { DocumentSnapshot } from './types';
import { AutoFixEngine } from './autofix/AutoFixEngine';
import { TextNormalizer } from './normalizer/TextNormalizer';
import { TableFormatter } from './formatter/TableFormatter';
import { LicenseManager } from './auth/LicenseManager';
import { RuleLoader } from './rules/RuleLoader';

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

const ruleLoader = RuleLoader.getInstance();

// Health Check
app.get('/api/health', (req: Request, res: Response) => {
  res.json({
    status: 'ok',
    version: '1.0.0',
    service: 'Document Standardizer Core Engine (ND30, HD05, Viettel)'
  });
});

// 1. Quét lỗi tài liệu
app.post('/api/scan', (req: Request, res: Response) => {
  try {
    const snapshot: DocumentSnapshot = req.body;
    if (!snapshot || !snapshot.paragraphs) {
      return res.status(400).json({ error: 'Invalid document snapshot' });
    }
    const result = AutoFixEngine.analyzeAndPlanAutoFix(snapshot);
    res.json(result);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// 2. Kế hoạch 1-Click Auto-Fix
app.post('/api/autofix', (req: Request, res: Response) => {
  try {
    const snapshot: DocumentSnapshot = req.body;
    if (!snapshot || !snapshot.paragraphs) {
      return res.status(400).json({ error: 'Invalid document snapshot' });
    }
    const analysis = AutoFixEngine.analyzeAndPlanAutoFix(snapshot);
    const tablePlans = TableFormatter.planTableFormatting(snapshot.tableCount || 0);

    res.json({
      regime: analysis.detectedRegime,
      regimeName: analysis.regimeName,
      actions: analysis.autoFixActions,
      tablePlans,
      summary: analysis.summary
    });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// 3. Chuẩn hóa chuỗi văn bản
app.post('/api/normalize', (req: Request, res: Response) => {
  try {
    const { text } = req.body;
    if (typeof text !== 'string') {
      return res.status(400).json({ error: 'Text string is required' });
    }
    const normalized = TextNormalizer.normalizeAll(text);
    res.json({ original: text, normalized });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// 4. Lấy danh sách quy tắc & từ điển
app.get('/api/rules', (req: Request, res: Response) => {
  res.json({
    regimes: [
      { code: 'ND30', name: 'Nghị định 30/2020/NĐ-CP (Hành chính Nhà nước)' },
      { code: 'DANG_HD05', name: 'Hướng dẫn 05-HD/VPTW (Văn bản của Đảng)' },
      { code: 'VIETTEL', name: 'Quy chế Thể thức Viettel (QĐ 11095)' }
    ],
    checksCount: 82,
    dictionaryStats: {
      administrativeUnits: ruleLoader.adminUnits.length,
      typoEntries: Object.keys(ruleLoader.typoDict).length,
      iyEntries: Object.keys(ruleLoader.iyDict).length,
      nonEndingAbbreviations: ruleLoader.nonEndingAbbrs.length
    }
  });
});

// 5. Quản lý bản quyền & VietQR
app.post('/api/license/verify', (req: Request, res: Response) => {
  const { licenseKey, deviceId, customerName, expiryDate } = req.body;
  const isValid = LicenseManager.verifyLicense(licenseKey, deviceId, customerName, expiryDate);
  res.json({ isValid, licenseKey, deviceId });
});

app.post('/api/license/create-qr', async (req: Request, res: Response) => {
  try {
    const { accountNo, bankCode, amount, memo } = req.body;
    const qrUrl = await LicenseManager.generateVietQR(accountNo || '123456789', bankCode || 'MB', amount || 299000, memo || 'KICH HOAT ADDIN');
    res.json({ qrUrl });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`[Core API] Server running at http://localhost:${port}`);
});
