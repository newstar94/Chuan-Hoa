using ChuanHoa.Client.Core.Text;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    internal static class StartupIntegrityGuard
    {
        private const string PinMetadataName = "ChuanHoa.SigningCertificateSha256";
        private static readonly object Sync = new object();
        private static readonly Guid WinTrustActionGenericVerifyV2 =
            new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
        private static bool _verified;

        internal static void VerifyOrThrow()
        {
            if (_verified) return;
            lock (Sync)
            {
                if (_verified) return;
                var assembly = typeof(StartupIntegrityGuard).Assembly;
                var expectedPin = assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
                    .Where(attribute => string.Equals(
                        attribute.Key, PinMetadataName, StringComparison.Ordinal))
                    .Select(attribute => attribute.Value ?? string.Empty)
                    .SingleOrDefault() ?? string.Empty;
                expectedPin = NormalizeSha256(expectedPin);

                var addInPath = Path.GetFullPath(assembly.Location);
                var installDirectory = Path.GetDirectoryName(addInPath);
                if (string.IsNullOrWhiteSpace(installDirectory))
                    throw new SecurityException(
                        "Không xác định được thư mục cài đặt Chuẩn hóa.");
                VerifySignedPe(addInPath, expectedPin);

                // Use the dependency location selected by the CLR instead of
                // probing a sibling path with File.Exists. The VSTO AppDomain can
                // deny that probe even when it has already loaded the dependency,
                // which incorrectly soft-disables an intact installation. Requiring
                // the loaded dependency to reside beside the add-in prevents this
                // from becoming a fallback to an assembly from another location.
                var clientCorePath = Path.GetFullPath(
                    typeof(VietnameseTypographyCleaner).Assembly.Location);
                var clientCoreDirectory = Path.GetDirectoryName(clientCorePath);
                if (string.IsNullOrWhiteSpace(clientCoreDirectory) ||
                    !string.Equals(
                        Path.GetFullPath(clientCoreDirectory).TrimEnd(Path.DirectorySeparatorChar),
                        Path.GetFullPath(installDirectory).TrimEnd(Path.DirectorySeparatorChar),
                        StringComparison.OrdinalIgnoreCase))
                    throw new SecurityException(
                        "Thành phần Chuẩn hóa được nạp từ thư mục không hợp lệ: " +
                        Path.GetFileName(clientCorePath));
                VerifySignedPe(clientCorePath, expectedPin);
                _verified = true;
            }
        }

        private static string NormalizeSha256(string value)
        {
            var normalized = (value ?? string.Empty).Replace("-", string.Empty)
                .Replace(" ", string.Empty).ToUpperInvariant();
            if (normalized.Length != 64 || normalized.Any(character =>
                    !(character >= '0' && character <= '9') &&
                    !(character >= 'A' && character <= 'F')))
                throw new SecurityException(
                    "Bản Chuẩn hóa không chứa certificate SHA-256 pin hợp lệ.");
            return normalized;
        }

        private static void VerifySignedPe(string path, string expectedPin)
        {
            if (!File.Exists(path))
                throw new SecurityException(
                    "Thiếu thành phần đã ký của Chuẩn hóa: " + Path.GetFileName(path));
            var trustStatus = VerifyAuthenticodeSignature(path);
            if (trustStatus != 0)
                throw new SecurityException(
                    "Thành phần Chuẩn hóa đã bị sửa hoặc chữ ký không còn tin cậy: " +
                    Path.GetFileName(path) + " (0x" + trustStatus.ToString("X8") + ").");

            X509Certificate signedCertificate;
            try
            {
                signedCertificate = X509Certificate.CreateFromSignedFile(path);
            }
            catch (CryptographicException exception)
            {
                throw new SecurityException(
                    "Không đọc được chữ ký của " + Path.GetFileName(path) + ".", exception);
            }
            using (var certificate = new X509Certificate2(signedCertificate))
            using (var sha256 = SHA256.Create())
            {
                var actualPin = BitConverter.ToString(
                    sha256.ComputeHash(certificate.RawData)).Replace("-", string.Empty);
                if (!string.Equals(actualPin, expectedPin, StringComparison.Ordinal))
                    throw new SecurityException(
                        "Chữ ký của " + Path.GetFileName(path) +
                        " không thuộc certificate đã pin.");
            }
        }

        private static uint VerifyAuthenticodeSignature(string path)
        {
            using (var fileInfo = new WinTrustFileInfo(path))
            using (var trustData = new WinTrustData(fileInfo))
                return WinVerifyTrust(
                    new IntPtr(-1), WinTrustActionGenericVerifyV2, trustData);
        }

        [DllImport("wintrust.dll", ExactSpelling = true, SetLastError = true,
            CharSet = CharSet.Unicode)]
        private static extern uint WinVerifyTrust(
            IntPtr windowHandle,
            [MarshalAs(UnmanagedType.LPStruct)] Guid actionId,
            WinTrustData trustData);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private sealed class WinTrustFileInfo : IDisposable
        {
            public uint StructSize;
            public IntPtr FilePath;
            public IntPtr FileHandle;
            public IntPtr KnownSubject;

            public WinTrustFileInfo(string path)
            {
                StructSize = (uint)Marshal.SizeOf(typeof(WinTrustFileInfo));
                FilePath = Marshal.StringToCoTaskMemUni(path);
                FileHandle = IntPtr.Zero;
                KnownSubject = IntPtr.Zero;
            }

            public void Dispose()
            {
                if (FilePath == IntPtr.Zero) return;
                Marshal.FreeCoTaskMem(FilePath);
                FilePath = IntPtr.Zero;
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private sealed class WinTrustData : IDisposable
        {
            public uint StructSize;
            public IntPtr PolicyCallbackData;
            public IntPtr SipClientData;
            public uint UiChoice;
            public uint RevocationChecks;
            public uint UnionChoice;
            public IntPtr UnionInfo;
            public uint StateAction;
            public IntPtr StateData;
            public IntPtr UrlReference;
            public uint ProviderFlags;
            public uint UiContext;

            public WinTrustData(WinTrustFileInfo fileInfo)
            {
                StructSize = (uint)Marshal.SizeOf(typeof(WinTrustData));
                UiChoice = 2; // WTD_UI_NONE
                RevocationChecks = 0; // WTD_REVOKE_NONE
                UnionChoice = 1; // WTD_CHOICE_FILE
                UnionInfo = Marshal.AllocCoTaskMem(
                    Marshal.SizeOf(typeof(WinTrustFileInfo)));
                Marshal.StructureToPtr(fileInfo, UnionInfo, false);
                StateAction = 0; // WTD_STATEACTION_IGNORE
                ProviderFlags = 0x00001000; // WTD_CACHE_ONLY_URL_RETRIEVAL
            }

            public void Dispose()
            {
                if (UnionInfo == IntPtr.Zero) return;
                Marshal.DestroyStructure(UnionInfo, typeof(WinTrustFileInfo));
                Marshal.FreeCoTaskMem(UnionInfo);
                UnionInfo = IntPtr.Zero;
            }
        }
    }
}
