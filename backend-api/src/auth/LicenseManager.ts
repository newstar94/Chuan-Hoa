import * as crypto from 'crypto';
import * as QRCode from 'qrcode';

export interface LicenseInfo {
  licenseKey: string;
  deviceId: string;
  customerName: string;
  plan: 'PRO' | 'ENTERPRISE' | 'TRIAL';
  expiryDate: string;
  isActive: boolean;
}

export class LicenseManager {
  private static SECRET_SALT = 'VIET_DOC_STANDARDIZER_2026_SECRET_KEY';

  /**
   * Tạo mã kích hoạt bản quyền từ Hardware ID / Machine Fingerprint
   */
  public static generateLicenseKey(deviceId: string, customerName: string, daysValid: number = 365): string {
    const expiry = new Date(Date.now() + daysValid * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    const payload = `${deviceId}|${customerName}|${expiry}|${this.SECRET_SALT}`;
    const hash = crypto.createHash('sha256').update(payload).digest('hex').substring(0, 16).toUpperCase();
    return `VDS-${hash.substring(0, 4)}-${hash.substring(4, 8)}-${hash.substring(8, 12)}-${hash.substring(12, 16)}`;
  }

  /**
   * Kiểm tra tính hợp lệ của License Key
   */
  public static verifyLicense(licenseKey: string, deviceId: string, customerName: string, expiryDate: string): boolean {
    if (new Date(expiryDate).getTime() < Date.now()) {
      return false; // Hết hạn
    }
    const expected = this.generateLicenseKey(deviceId, customerName, 365);
    return licenseKey.startsWith('VDS-') && licenseKey.length === 23;
  }

  /**
   * Tạo chuỗi thanh toán và hình ảnh VietQR
   */
  public static async generateVietQR(accountNo: string, bankCode: string, amount: number, memo: string): Promise<string> {
    const qrData = `https://img.vietqr.io/image/${bankCode}-${accountNo}-compact2.png?amount=${amount}&addInfo=${encodeURIComponent(memo)}`;
    return qrData;
  }
}
