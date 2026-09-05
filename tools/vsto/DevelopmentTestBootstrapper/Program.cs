using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;

namespace ChuanHoa.DevelopmentTestBootstrapper
{
    internal static class Program
    {
        private const string Title = "Chu\u1EA9n h\u00F3a - C\u00E0i \u0111\u1EB7t th\u1EED nghi\u1EC7m";
        private const string PayloadResource = "ChuanHoa.DevelopmentInstaller.Payload.zip";
        private const string VersionResource = "ChuanHoa.DevelopmentInstaller.Version.txt";
        private const string TrustedPublicKeyPinResource =
            "ChuanHoa.DevelopmentInstaller.TrustedPublicKey.sha256";
        private const string SigningCertificatePinResource =
            "ChuanHoa.DevelopmentInstaller.SigningCertificate.sha256";
        private const string SigningRootCertificatePinResource =
            "ChuanHoa.DevelopmentInstaller.SigningRootCertificate.sha256";
        private const string ExpectedCertificateSubject = "CN=Chuan Hoa Local Development";
        private const string ExpectedRootCertificateSubject =
            "CN=Chuan Hoa Local Development Root";
        private const string WordRibbonValidationValueName =
            "ChuanHoa.AddIn.Vsto.Microsoft.Word.Document";
        private const string AppsFeaturesRegistryPath =
            @"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChuanHoa.DevelopmentTest";
        private const string CachedInstallerFileName =
            "ChuanHoa_Development_Test_Setup.exe";
        private static readonly Guid WinTrustActionGenericVerifyV2 =
            new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
        private const uint ErrorSuccess = 0x00000000;
        private const uint CertEUntrustedRoot = 0x800B0109;
        private const uint CertEChaining = 0x800B010A;
        private const uint TrustESubjectNotTrusted = 0x800B0004;

        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var quiet = Array.Exists(
                args ?? new string[0],
                value => string.Equals(value, "/quiet", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(value, "/silent", StringComparison.OrdinalIgnoreCase));
            var uninstall = Array.Exists(args ?? new string[0],
                value => string.Equals(value, "/uninstall", StringComparison.OrdinalIgnoreCase));
            var cleanupCache = Array.Exists(args ?? new string[0],
                value => string.Equals(value, "/cleanup-cache", StringComparison.OrdinalIgnoreCase));

            try
            {
                var signingCertificatePin = ReadSha256Resource(SigningCertificatePinResource);
                var signingRootCertificatePin =
                    ReadSha256Resource(SigningRootCertificatePinResource);
                VerifyRunningInstallerSigner(signingCertificatePin);
                if (cleanupCache)
                    return CleanupInstallerCache(
                        ReadVersionResource(), ReadParentProcessId(args ?? new string[0]));

                if (Process.GetProcessesByName("WINWORD").Length > 0)
                {
                    if (!quiet)
                    {
                        MessageBox.Show(
                            "H\u00E3y \u0111\u00F3ng ho\u00E0n to\u00E0n Microsoft Word tr\u01B0\u1EDBc khi c\u00E0i \u0111\u1EB7t.",
                            Title,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                    return 2;
                }

                var baseDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ChuanHoa",
                    "DevelopmentInstaller");
                if (uninstall)
                {
                    UninstallDevelopmentChannel(baseDirectory);
                    ScheduleInstallerCacheCleanup(ReadVersionResource(), signingCertificatePin);
                    if (!quiet)
                        MessageBox.Show(
                            "Đã gỡ add-in Chuẩn hóa Development Test. Từ điển cá nhân và tài liệu của bạn được giữ nguyên.",
                            Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return 0;
                }

                if (!quiet)
                {
                    var answer = MessageBox.Show(
                        "\u0110\u00E2y l\u00E0 b\u1EA3n Development Test, ch\u1EC9 d\u00F9ng tr\u00EAn m\u00E1y th\u1EED nghi\u1EC7m.\n\n" +
                        "Tr\u00ECnh c\u00E0i \u0111\u1EB7t s\u1EBD th\u00EAm ch\u1EE9ng th\u01B0 Development v\u00E0o kho tin c\u1EADy c\u1EE7a t\u00E0i kho\u1EA3n Windows hi\u1EC7n t\u1EA1i.\n\n" +
                        "N\u1EBFu \u0111\u00E3 c\u00F3 b\u1EA3n Chu\u1EA9n h\u00F3a Development c\u0169, tr\u00ECnh c\u00E0i \u0111\u1EB7t s\u1EBD t\u1EF1 thay th\u1EBF b\u1EA3n \u0111\u00F3.\n\n" +
                        "Ti\u1EBFp t\u1EE5c?",
                        Title,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2);
                    if (answer != DialogResult.Yes) return 1;
                }

                var version = ReadVersionResource();
                var trustedPublicKeyPin = ReadSha256Resource(TrustedPublicKeyPinResource);
                EnsureNoLegacyClickOnceDevelopmentAddIn();
                var installDirectory = Path.Combine(baseDirectory, "Current");
                var stagingDirectory = Path.Combine(baseDirectory, "Staging-" + Guid.NewGuid().ToString("N"));
                PrepareStagingDirectory(baseDirectory, stagingDirectory);
                using (var transaction = new InstallTransaction(
                    baseDirectory, stagingDirectory, installDirectory, version))
                {
                    try
                    {
                        ExtractPayload(stagingDirectory);
                        VerifyStagedPayload(stagingDirectory, version,
                            trustedPublicKeyPin, signingCertificatePin,
                            signingRootCertificatePin);
                        var certificatePath = Path.Combine(
                            stagingDirectory, "ChuanHoa.LocalDevelopment.Public.cer");
                        var rootCertificatePath = Path.Combine(
                            stagingDirectory, "ChuanHoa.LocalDevelopment.Root.cer");
                        transaction.TrustDevelopmentCertificate(
                            rootCertificatePath, signingRootCertificatePin,
                            certificatePath, signingCertificatePin);
                        InjectFault("after-trusted-certificate");
                        transaction.InstallTrustedDevelopmentKey(stagingDirectory, trustedPublicKeyPin);
                        InjectFault("after-trusted-key");
                        EnsureRequiredRuntimes();
                        EnsureDevelopmentAccess(stagingDirectory, version);
                        InjectFault("after-access-smoke");

                        transaction.Activate();
                        InjectFault("after-current-switch");
                        var cachedInstallerPath = transaction.CacheRunningInstaller(
                            signingCertificatePin);
                        RegisterInstalledDevelopmentManifest(installDirectory);
                        RegisterInstallerState(version, installDirectory,
                            cachedInstallerPath, trustedPublicKeyPin,
                            signingCertificatePin, signingRootCertificatePin);
                        RegisterAppsFeatures(version, installDirectory, cachedInstallerPath);
                        InjectFault("after-registry");
                        VerifyDirectRegistration(version, installDirectory,
                            cachedInstallerPath, trustedPublicKeyPin,
                            signingCertificatePin, signingRootCertificatePin);
                        InjectFault("after-verification");
                        transaction.Commit();
                    }
                    catch (Exception installError)
                    {
                        try
                        {
                            transaction.Rollback();
                        }
                        catch (Exception rollbackError)
                        {
                            throw new AggregateException(
                                "Cài đặt thất bại và rollback cũng không hoàn tất.",
                                installError,
                                rollbackError);
                        }
                        throw;
                    }
                }

                TryClearWordRibbonValidationCache();
                TryClearWordResiliencyAndVstoSolutionMetadata();

                if (!quiet)
                {
                    MessageBox.Show(
                        "\u0110\u00E3 c\u00E0i add-in Chu\u1EA9n h\u00F3a Development Test " + version + ".\n\n" +
                        "H\u00E3y m\u1EDF Microsoft Word. C\u00E1c ch\u1EE9c n\u0103ng h\u1ED7 tr\u1EE3 c\u1EA3 t\u00E0i li\u1EC7u .doc/.docx \u0111\u00E3 l\u01B0u v\u00E0 v\u0103n b\u1EA3n m\u1EDBi ch\u01B0a l\u01B0u.",
                        Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception exception)
            {
                if (quiet) Console.Error.WriteLine(exception);
                if (!quiet)
                {
                    MessageBox.Show(
                        "C\u00E0i \u0111\u1EB7t kh\u00F4ng th\u00E0nh c\u00F4ng.\n\n" + exception.Message,
                        Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                return 10;
            }
        }

        private static void UninstallDevelopmentChannel(string baseDirectory)
        {
            UninstallClickOnceDevelopmentAddIn(baseDirectory);
            var ownedCertificates = ReadOwnedDevelopmentCertificates(baseDirectory);
            RemoveDirectRegistrationIfOwned(baseDirectory);
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            {
                currentUser.DeleteSubKeyTree(@"Software\ChuanHoa\DevelopmentInstaller", false);
                currentUser.DeleteSubKeyTree(AppsFeaturesRegistryPath, false);
            }

            if (Directory.Exists(baseDirectory))
            {
                var normalizedBase = Path.GetFullPath(baseDirectory)
                    .TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                foreach (var directory in Directory.GetDirectories(baseDirectory))
                {
                    var normalized = Path.GetFullPath(directory)
                        .TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                    var name = Path.GetFileName(directory);
                    if (!normalized.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase) ||
                        !(string.Equals(name, "Current", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(name, "Previous", StringComparison.OrdinalIgnoreCase) ||
                          name.StartsWith("Staging-", StringComparison.OrdinalIgnoreCase)))
                        continue;
                    Directory.Delete(directory, true);
                }
                if (Directory.GetFileSystemEntries(baseDirectory).Length == 0)
                    Directory.Delete(baseDirectory);
            }

            var trustedKeyPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ChuanHoa", "Development", "trusted-key.xml");
            if (File.Exists(trustedKeyPath)) File.Delete(trustedKeyPath);
            foreach (var ownedCertificate in ownedCertificates)
                RemoveCertificate(ownedCertificate.Item1, ownedCertificate.Item2,
                    ownedCertificate.Item3);
            ClearWordRibbonValidationCache();
        }

        private static string GetInstallerCacheRoot()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ChuanHoa", "InstallerCache", "Development");
        }

        private static int ReadParentProcessId(IEnumerable<string> args)
        {
            const string prefix = "/parent-pid=";
            foreach (var argument in args)
            {
                if (!argument.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                int value;
                if (int.TryParse(argument.Substring(prefix.Length), out value) && value > 0)
                    return value;
                throw new InvalidOperationException("Cleanup parent PID is invalid.");
            }
            throw new InvalidOperationException("Cleanup parent PID is missing.");
        }

        private static int CleanupInstallerCache(string version, int parentProcessId)
        {
            try
            {
                using (var parent = Process.GetProcessById(parentProcessId))
                {
                    if (!parent.WaitForExit(30000))
                        throw new TimeoutException("Installer cleanup waited more than 30 seconds.");
                }
            }
            catch (ArgumentException)
            {
                // The parent has already exited, so cleanup can continue.
            }

            var cacheRoot = GetInstallerCacheRoot();
            var versionDirectory = Path.Combine(cacheRoot, version);
            AssertDirectChildDirectory(cacheRoot, versionDirectory, version);
            if (Directory.Exists(versionDirectory)) Directory.Delete(versionDirectory, true);
            if (Directory.Exists(cacheRoot) && Directory.GetFileSystemEntries(cacheRoot).Length == 0)
                Directory.Delete(cacheRoot);
            return 0;
        }

        private static void ScheduleInstallerCacheCleanup(string version,
            string signingCertificateSha256)
        {
            var source = Assembly.GetExecutingAssembly().Location;
            var helper = Path.Combine(Path.GetTempPath(),
                "ChuanHoaInstallerCleanup-" + Guid.NewGuid().ToString("N") + ".exe");
            File.Copy(source, helper, false);
            if (!string.Equals(ComputeFileSha256(source), ComputeFileSha256(helper),
                    StringComparison.Ordinal))
                throw new InvalidOperationException("Installer cleanup helper does not match the signed installer.");
            VerifyInstallerSigner(helper, signingCertificateSha256);
            var startInfo = new ProcessStartInfo
            {
                FileName = helper,
                Arguments = "/cleanup-cache /parent-pid=" +
                    Process.GetCurrentProcess().Id + " /quiet",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            if (Process.Start(startInfo) == null)
                throw new InvalidOperationException("Could not start the installer cache cleanup helper.");
        }

        private static void AssertDirectChildDirectory(string root, string target,
            string expectedName)
        {
            var normalizedRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar);
            var normalizedTarget = Path.GetFullPath(target).TrimEnd(Path.DirectorySeparatorChar);
            var parent = Directory.GetParent(normalizedTarget);
            if (parent == null ||
                !string.Equals(parent.FullName.TrimEnd(Path.DirectorySeparatorChar),
                    normalizedRoot, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetFileName(normalizedTarget), expectedName,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("Unsafe installer cache directory.");
        }

        private static Tuple<StoreName, string, string>[] ReadOwnedDevelopmentCertificates(
            string baseDirectory)
        {
            var result = new List<Tuple<StoreName, string, string>>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var expectedCertificateSha256 =
                ReadSha256Resource(SigningCertificatePinResource);
            var expectedRootCertificateSha256 =
                ReadSha256Resource(SigningRootCertificatePinResource);
            foreach (var name in new[] { "Current", "Previous" })
            {
                foreach (var owned in new[]
                {
                    Tuple.Create("ChuanHoa.LocalDevelopment.Public.cer",
                        ExpectedCertificateSubject, expectedCertificateSha256,
                        StoreName.TrustedPublisher),
                    Tuple.Create("ChuanHoa.LocalDevelopment.Root.cer",
                        ExpectedRootCertificateSubject,
                        expectedRootCertificateSha256, StoreName.Root)
                })
                {
                    var path = Path.Combine(baseDirectory, name, owned.Item1);
                    if (!File.Exists(path)) continue;
                    using (var certificate = new X509Certificate2(path))
                    {
                        if (!string.Equals(certificate.Subject, owned.Item2,
                                StringComparison.Ordinal) ||
                            !string.Equals(ComputeSha256(certificate.RawData),
                                owned.Item3, StringComparison.Ordinal) ||
                            string.IsNullOrWhiteSpace(certificate.Thumbprint))
                            continue;
                        var key = owned.Item4 + ":" + certificate.Thumbprint;
                        if (seen.Add(key))
                            result.Add(Tuple.Create(owned.Item4,
                                certificate.Thumbprint, owned.Item3));
                    }
                }
            }
            return result.ToArray();
        }

        private static void RemoveDirectRegistrationIfOwned(string baseDirectory)
        {
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var addIn = currentUser.OpenSubKey(
                @"Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto"))
            {
                if (addIn == null) return;
                var manifest = Convert.ToString(addIn.GetValue("Manifest")) ?? string.Empty;
                if (manifest.EndsWith("|vstolocal", StringComparison.OrdinalIgnoreCase))
                    manifest = manifest.Substring(0, manifest.Length - "|vstolocal".Length);
                Uri uri;
                if (!Uri.TryCreate(manifest, UriKind.Absolute, out uri) || !uri.IsFile) return;
                var normalizedBase = Path.GetFullPath(baseDirectory)
                    .TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                var registeredPath = Path.GetFullPath(uri.LocalPath);
                if (!registeredPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase)) return;
            }
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
                currentUser.DeleteSubKeyTree(
                    @"Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto", false);
        }

        private static void RemoveCertificate(StoreName storeName, string thumbprint,
            string expectedSha256)
        {
            if (!HasPinnedCertificate(storeName, thumbprint, expectedSha256)) return;
            RunCertificateUtility("-f -user -delstore " + storeName + " " + thumbprint);
            if (HasPinnedCertificate(storeName, thumbprint, expectedSha256))
                throw new InvalidOperationException(
                    "KhÃ´ng xÃ³a Ä‘Æ°á»£c certificate Development Ä‘Ã£ pin khá»i " + storeName + ".");
        }

        private static bool HasPinnedCertificate(StoreName storeName, string thumbprint,
            string expectedSha256)
        {
            using (var store = new X509Store(storeName, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var matches = store.Certificates.Find(
                    X509FindType.FindByThumbprint, thumbprint, false);
                foreach (X509Certificate2 certificate in matches)
                {
                    try
                    {
                        if (!string.Equals(certificate.Subject, ExpectedCertificateSubject,
                                StringComparison.Ordinal) ||
                            !string.Equals(ComputeSha256(certificate.RawData), expectedSha256,
                                StringComparison.Ordinal))
                            throw new InvalidOperationException(
                                "Certificate Development trong " + storeName +
                                " khÃ´ng khá»›p subject/SHA-256 Ä‘Ã£ pin.");
                        return true;
                    }
                    finally { certificate.Dispose(); }
                }
                return false;
            }
        }

        private static void RunCertificateUtility(string arguments)
        {
            var executable = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System), "certutil.exe");
            if (!File.Exists(executable))
                throw new FileNotFoundException("KhÃ´ng tÃ¬m tháº¥y certutil.exe cá»§a Windows.", executable);
            var startInfo = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            var result = RunCapturedProcess(startInfo, 15000, "certutil.exe");
            if (result.ExitCode != 0)
                throw new InvalidOperationException(
                    "certutil.exe tháº¥t báº¡i vá»›i mÃ£ " + result.ExitCode + ".\n" +
                    (string.IsNullOrWhiteSpace(result.StandardError)
                        ? result.StandardOutput
                        : result.StandardError));
        }

        private static ProcessResult RunCapturedProcess(ProcessStartInfo startInfo,
            int timeoutMilliseconds, string label)
        {
            var output = new StringBuilder();
            var error = new StringBuilder();
            using (var process = new Process { StartInfo = startInfo })
            {
                process.OutputDataReceived += (sender, eventArgs) =>
                {
                    if (eventArgs.Data != null) output.AppendLine(eventArgs.Data);
                };
                process.ErrorDataReceived += (sender, eventArgs) =>
                {
                    if (eventArgs.Data != null) error.AppendLine(eventArgs.Data);
                };
                if (!process.Start())
                    throw new InvalidOperationException(
                        "KhÃ´ng khá»Ÿi Ä‘á»™ng Ä‘Æ°á»£c " + label + ".");
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                if (!process.WaitForExit(timeoutMilliseconds))
                {
                    try { process.Kill(); } catch { }
                    try { process.WaitForExit(2000); } catch { }
                    throw new TimeoutException(
                        label + " quÃ¡ thá»i gian " +
                        (timeoutMilliseconds / 1000) + " giÃ¢y.");
                }
                // Wait once without a timeout after the process has exited so
                // asynchronous stdout/stderr handlers flush their final lines.
                process.WaitForExit();
                return new ProcessResult(
                    process.ExitCode, output.ToString().Trim(), error.ToString().Trim());
            }
        }

        private static string ReadTextResource(string resourceName)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new InvalidOperationException("Installer resource is missing: " + resourceName);
                using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
            }
        }

        private static string ReadVersionResource()
        {
            var version = ReadTextResource(VersionResource).Trim();
            Version parsed;
            if (!Version.TryParse(version, out parsed) || parsed.Build < 0 ||
                parsed.Revision < 0 || parsed.ToString(4) != version)
                throw new InvalidOperationException(
                    "Phiên bản bộ cài phải có đúng bốn thành phần số.");
            return version;
        }

        private static string ReadSha256Resource(string resourceName)
        {
            var value = ReadTextResource(resourceName).Trim().Replace("-", string.Empty)
                .Replace(" ", string.Empty).ToUpperInvariant();
            if (value.Length != 64 || value.Any(character =>
                    !(character >= '0' && character <= '9') &&
                    !(character >= 'A' && character <= 'F')))
                throw new InvalidOperationException(
                    "Bộ cài chứa SHA-256 pin không hợp lệ: " + resourceName);
            return value;
        }

        private static string ComputeSha256(byte[] value)
        {
            using (var sha256 = SHA256.Create())
                return BitConverter.ToString(sha256.ComputeHash(value)).Replace("-", string.Empty);
        }

        private static string ComputeFileSha256(string path)
        {
            using (var input = File.OpenRead(path))
            using (var sha256 = SHA256.Create())
                return BitConverter.ToString(sha256.ComputeHash(input)).Replace("-", string.Empty);
        }

        private static void VerifyRunningInstallerSigner(string expectedSha256)
        {
            VerifyInstallerSigner(Assembly.GetExecutingAssembly().Location, expectedSha256);
        }

        private static void VerifyInstallerSigner(string installerPath, string expectedSha256)
        {
            var trustStatus = VerifyAuthenticodeSignature(installerPath);
            // The Development certificate is self-signed and is intentionally
            // pinned by SHA-256. On a clean machine it is not trusted until the
            // installer transaction adds it, so accept only chain-trust errors
            // here. Digest/signature errors such as TRUST_E_BAD_DIGEST remain
            // fatal and are never hidden by the certificate pin.
            if (trustStatus != ErrorSuccess &&
                trustStatus != CertEUntrustedRoot &&
                trustStatus != CertEChaining &&
                trustStatus != TrustESubjectNotTrusted)
                throw new InvalidOperationException(
                    "Chữ ký Authenticode không hợp lệ: 0x" +
                    trustStatus.ToString("X8") + ".");

            X509Certificate signedCertificate;
            try
            {
                signedCertificate = X509Certificate.CreateFromSignedFile(installerPath);
            }
            catch (CryptographicException exception)
            {
                throw new InvalidOperationException(
                    "Bộ cài Development chưa có chữ ký Authenticode hợp lệ.", exception);
            }

            using (var certificate = new X509Certificate2(signedCertificate))
            {
                var actualSha256 = ComputeSha256(certificate.RawData);
                if (!string.Equals(actualSha256, expectedSha256, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Chữ ký bộ cài không khớp certificate SHA-256 đã pin.");
            }
        }

        private static uint VerifyAuthenticodeSignature(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("Không tìm thấy PE cần xác minh chữ ký.", path);
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

        private static void InjectFault(string point)
        {
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("CHUANHOA_INSTALLER_ENABLE_FAULT_INJECTION"),
                    "1",
                    StringComparison.Ordinal) ||
                !string.Equals(
                    Environment.GetEnvironmentVariable("CHUANHOA_INSTALLER_FAULT_POINT"),
                    point,
                    StringComparison.OrdinalIgnoreCase))
                return;
            throw new InvalidOperationException("Injected installer fault: " + point);
        }

        private static void PrepareStagingDirectory(string baseDirectory, string stagingDirectory)
        {
            Directory.CreateDirectory(baseDirectory);
            var normalizedBase = Path.GetFullPath(baseDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            var normalizedTarget = Path.GetFullPath(stagingDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!normalizedTarget.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Unsafe installer extraction directory.");
            Directory.CreateDirectory(stagingDirectory);
        }

        private static void VerifyStagedPayload(string stagingDirectory, string expectedVersion,
            string trustedPublicKeySha256, string signingCertificateSha256,
            string signingRootCertificateSha256)
        {
            var expectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ChuanHoa.AddIn.Vsto.vsto",
                "ChuanHoa.AddIn.Vsto.dll.manifest",
                "ChuanHoa.AddIn.Vsto.dll",
                "ChuanHoa.Client.Core.dll",
                "ChuanHoa.DevelopmentAccessSmoke.exe",
                "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
                "ChuanHoa.LocalDevelopment.Public.cer",
                "ChuanHoa.LocalDevelopment.Root.cer",
                "DevelopmentSupport/trusted-key.xml"
            };
            var actualFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in Directory.GetFiles(stagingDirectory, "*", SearchOption.AllDirectories))
            {
                var relativePath = path.Substring(stagingDirectory.Length)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    .Replace(Path.DirectorySeparatorChar, '/');
                actualFiles.Add(relativePath);
            }
            var missing = expectedFiles.Except(actualFiles).ToArray();
            var unexpected = actualFiles.Except(expectedFiles).ToArray();
            if (missing.Length != 0 || unexpected.Length != 0)
                throw new InvalidOperationException(
                    "Payload không đúng allowlist. Thiếu=[" + string.Join(", ", missing) +
                    "]; Thừa=[" + string.Join(", ", unexpected) + "].");

            var trustedKeyPath = Path.Combine(
                stagingDirectory, "DevelopmentSupport", "trusted-key.xml");
            if (!string.Equals(ComputeFileSha256(trustedKeyPath), trustedPublicKeySha256,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Khóa công khai Development không khớp SHA-256 đã pin.");

            var certificatePath = Path.Combine(
                stagingDirectory, "ChuanHoa.LocalDevelopment.Public.cer");
            var rootCertificatePath = Path.Combine(
                stagingDirectory, "ChuanHoa.LocalDevelopment.Root.cer");
            using (var certificate = new X509Certificate2(certificatePath))
            using (var rootCertificate = new X509Certificate2(rootCertificatePath))
            {
                if (!string.Equals(ComputeSha256(certificate.RawData), signingCertificateSha256,
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Certificate Development không khớp SHA-256 đã pin.");
                if (!string.Equals(ComputeSha256(rootCertificate.RawData),
                        signingRootCertificateSha256, StringComparison.Ordinal) ||
                    !string.Equals(rootCertificate.Subject,
                        ExpectedRootCertificateSubject, StringComparison.Ordinal) ||
                    !string.Equals(rootCertificate.Subject, rootCertificate.Issuer,
                        StringComparison.Ordinal) ||
                    !string.Equals(certificate.Issuer, rootCertificate.Subject,
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Chuỗi certificate Development không khớp root SHA-256 đã pin.");
            }
            foreach (var ownedPeName in new[]
            {
                "ChuanHoa.AddIn.Vsto.dll",
                "ChuanHoa.Client.Core.dll",
                "ChuanHoa.DevelopmentAccessSmoke.exe"
            })
                VerifyInstallerSigner(
                    Path.Combine(stagingDirectory, ownedPeName), signingCertificateSha256);

            var assemblyPath = Path.Combine(stagingDirectory, "ChuanHoa.AddIn.Vsto.dll");
            var assemblyName = AssemblyName.GetAssemblyName(assemblyPath);
            var versionInfo = FileVersionInfo.GetVersionInfo(assemblyPath);
            if (!string.Equals(assemblyName.Version.ToString(), expectedVersion,
                    StringComparison.Ordinal) ||
                !string.Equals(versionInfo.FileVersion, expectedVersion,
                    StringComparison.Ordinal) ||
                !string.Equals(versionInfo.ProductVersion, expectedVersion,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "DLL Development không khớp ProductVersion " + expectedVersion + ".");

            AssertManifestIdentityVersion(
                Path.Combine(stagingDirectory, "ChuanHoa.AddIn.Vsto.vsto"),
                "ChuanHoa.AddIn.Vsto.vsto",
                expectedVersion,
                certificatePath);
            AssertManifestIdentityVersion(
                Path.Combine(stagingDirectory, "ChuanHoa.AddIn.Vsto.dll.manifest"),
                "ChuanHoa.AddIn.Vsto.dll",
                expectedVersion,
                certificatePath);
        }

        private static void AssertManifestIdentityVersion(string path, string name,
            string expectedVersion, string certificatePath)
        {
            VerifySignedManifest(path, certificatePath);
            VerifyManifestFileHashes(path, Path.GetDirectoryName(path));
            var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
            var identities = document.Descendants()
                .Where(element => element.Name.LocalName == "assemblyIdentity" &&
                    string.Equals((string)element.Attribute("name"), name,
                        StringComparison.Ordinal))
                .ToArray();
            if (identities.Length == 0 || identities.Any(identity =>
                    !string.Equals((string)identity.Attribute("version"), expectedVersion,
                        StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "Manifest " + Path.GetFileName(path) +
                    " không khớp ProductVersion " + expectedVersion + ".");
            }
        }

        private static void VerifyManifestFileHashes(string manifestPath,
            string payloadDirectory)
        {
            var normalizedPayloadDirectory = Path.GetFullPath(payloadDirectory)
                .TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            var document = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
            foreach (var element in document.Descendants().Where(candidate =>
                candidate.Name.LocalName == "file" ||
                candidate.Name.LocalName == "dependentAssembly"))
            {
                var codebase = (string)element.Attribute("codebase");
                if (string.IsNullOrWhiteSpace(codebase)) continue;
                if (Uri.IsWellFormedUriString(codebase, UriKind.Absolute))
                    throw new InvalidOperationException(
                        "Manifest chứa codebase tuyệt đối không an toàn: " + codebase);
                var relativePath = Uri.UnescapeDataString(codebase)
                    .Replace('/', Path.DirectorySeparatorChar);
                var payloadPath = Path.GetFullPath(
                    Path.Combine(payloadDirectory, relativePath));
                if (!payloadPath.StartsWith(normalizedPayloadDirectory,
                        StringComparison.OrdinalIgnoreCase) || !File.Exists(payloadPath))
                    throw new InvalidOperationException(
                        "Manifest tham chiếu payload thiếu hoặc ngoài staging: " + codebase);

                // Mage records the pre-Authenticode size/digest for project PE
                // files. Authenticode appends a certificate table after that
                // manifest is generated, so a raw-file comparison would reject
                // a correctly signed build. Those two owned PE files are
                // instead covered by WinVerifyTrust plus the signer pin above.
                // Non-PE dependencies and the application manifest still use
                // the exact manifest size/SHA-256 checks below.
                var payloadFileName = Path.GetFileName(payloadPath);
                if (string.Equals(payloadFileName, "ChuanHoa.AddIn.Vsto.dll",
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(payloadFileName, "ChuanHoa.Client.Core.dll",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                long expectedSize;
                if (!long.TryParse((string)element.Attribute("size"), out expectedSize) ||
                    expectedSize < 0 || new FileInfo(payloadPath).Length != expectedSize)
                    throw new InvalidOperationException(
                        "Kích thước payload không khớp manifest: " + codebase);

                var hash = element.Elements().FirstOrDefault(candidate =>
                    candidate.Name.LocalName == "hash");
                var digestMethod = hash == null ? null : hash.Descendants()
                    .FirstOrDefault(candidate => candidate.Name.LocalName == "DigestMethod");
                var digestValue = hash == null ? null : hash.Descendants()
                    .FirstOrDefault(candidate => candidate.Name.LocalName == "DigestValue");
                if (digestMethod == null || digestValue == null ||
                    !string.Equals((string)digestMethod.Attribute("Algorithm"),
                        "http://www.w3.org/2000/09/xmldsig#sha256",
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Manifest thiếu SHA-256 digest hợp lệ cho: " + codebase);
                var actualDigest = Convert.ToBase64String(
                    ComputeSha256Bytes(File.ReadAllBytes(payloadPath)));
                if (!string.Equals(actualDigest, digestValue.Value.Trim(),
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "SHA-256 payload không khớp manifest: " + codebase);
            }
        }

        private static byte[] ComputeSha256Bytes(byte[] value)
        {
            using (var sha256 = SHA256.Create())
                return sha256.ComputeHash(value);
        }

        private static void VerifySignedManifest(string manifestPath, string certificatePath)
        {
            CryptoConfig.AddAlgorithm(
                typeof(RsaSha256SignatureDescription),
                "http://www.w3.org/2000/09/xmldsig#rsa-sha256");
            CryptoConfig.AddAlgorithm(
                typeof(SHA256CryptoServiceProvider),
                "http://www.w3.org/2000/09/xmldsig#sha256",
                "http://www.w3.org/2001/04/xmlenc#sha256");

            var document = new XmlDocument { PreserveWhitespace = true };
            document.Load(manifestPath);
            var signatureElement = document.SelectSingleNode(
                "//*[local-name()='Signature' and @Id='StrongNameSignature']") as XmlElement;
            if (signatureElement == null)
                throw new InvalidOperationException(
                    "Manifest thiếu chữ ký StrongNameSignature: " +
                    Path.GetFileName(manifestPath));

            var signedXml = new SignedXml(document);
            signedXml.LoadXml(signatureElement);
            using (var certificate = new X509Certificate2(certificatePath))
            {
                if (!signedXml.CheckSignature(certificate, true))
                    throw new InvalidOperationException(
                        "Chữ ký manifest không hợp lệ hoặc không khớp certificate đã pin: " +
                        Path.GetFileName(manifestPath));
            }
        }

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
                PolicyCallbackData = IntPtr.Zero;
                SipClientData = IntPtr.Zero;
                UiChoice = 2; // WTD_UI_NONE
                RevocationChecks = 0; // WTD_REVOKE_NONE
                UnionChoice = 1; // WTD_CHOICE_FILE
                UnionInfo = Marshal.AllocCoTaskMem(
                    Marshal.SizeOf(typeof(WinTrustFileInfo)));
                Marshal.StructureToPtr(fileInfo, UnionInfo, false);
                StateAction = 0; // WTD_STATEACTION_IGNORE
                StateData = IntPtr.Zero;
                UrlReference = IntPtr.Zero;
                // Do not allow signature validation to trigger network access.
                ProviderFlags = 0x00001000; // WTD_CACHE_ONLY_URL_RETRIEVAL
                UiContext = 0;
            }

            public void Dispose()
            {
                if (UnionInfo == IntPtr.Zero) return;
                Marshal.DestroyStructure(UnionInfo, typeof(WinTrustFileInfo));
                Marshal.FreeCoTaskMem(UnionInfo);
                UnionInfo = IntPtr.Zero;
            }
        }

        private static void EnsureRequiredRuntimes()
        {
            if (!IsNetFramework48Installed())
                throw new InvalidOperationException(
                    "Máy chưa có .NET Framework 4.8. Hãy cài thành phần này rồi chạy lại bộ cài.");
            if (!IsVstoRuntimeInstalled())
                throw new InvalidOperationException(
                    "Máy chưa có Microsoft Visual Studio Tools for Office Runtime. " +
                    "Hãy cài VSTO Runtime rồi chạy lại bộ cài.");
        }

        private static bool IsNetFramework48Installed()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (var framework = machine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (framework != null && Convert.ToInt32(framework.GetValue("Release", 0)) >= 528040)
                        return true;
                }
            }
            return false;
        }

        private static void UninstallClickOnceDevelopmentAddIn(string baseDirectory)
        {
            InstalledClickOnceAddIn installed;
            if (!TryFindInstalledClickOnceAddIn(out installed)) return;

            Uri manifestUri;
            if (!Uri.TryCreate(installed.ManifestUrl, UriKind.Absolute, out manifestUri) || !manifestUri.IsFile)
                throw new InvalidOperationException(
                    "Bản Chuẩn hóa đang cài không có địa chỉ Development cục bộ hợp lệ.");
            var normalizedBase = Path.GetFullPath(baseDirectory).TrimEnd(Path.DirectorySeparatorChar) +
                Path.DirectorySeparatorChar;
            var manifestPath = Path.GetFullPath(manifestUri.LocalPath);
            if (!manifestPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetFileName(manifestPath), "ChuanHoa.AddIn.Vsto.vsto", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Bản Chuẩn hóa đang cài không thuộc kênh Development cục bộ; không tự gỡ để tránh ảnh hưởng bản khác.");
            if (!File.Exists(manifestPath))
                throw new FileNotFoundException(
                    "Manifest của bản Development cũ đã bị mất; không thể gỡ an toàn.", manifestPath);

            var vstoInstaller = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                "Microsoft Shared", "VSTO", "10.0", "VSTOInstaller.exe");
            if (!File.Exists(vstoInstaller))
                throw new FileNotFoundException("Không tìm thấy VSTOInstaller.exe để gỡ bản Development cũ.", vstoInstaller);
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = vstoInstaller,
                Arguments = "/Uninstall \"" + manifestUri.AbsoluteUri + "\" /Silent",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            if (process == null) throw new InvalidOperationException("Không khởi động được trình gỡ bản Development cũ.");
            if (!process.WaitForExit(60000))
            {
                try { process.Kill(); } catch { }
                throw new TimeoutException(
                    "VSTOInstaller.exe quá thời gian 60 giây khi gỡ bản Development cũ.");
            }
            if (process.ExitCode != 0)
                throw new InvalidOperationException(
                    "Không gỡ được bản Chuẩn hóa Development " + installed.Version +
                    " (mã " + process.ExitCode + ").");
            InstalledClickOnceAddIn remaining;
            if (TryFindInstalledClickOnceAddIn(out remaining))
                throw new InvalidOperationException(
                    "Windows vẫn ghi nhận bản Chuẩn hóa Development " + remaining.Version + " sau khi gỡ.");
        }

        private static void EnsureNoLegacyClickOnceDevelopmentAddIn()
        {
            InstalledClickOnceAddIn installed;
            if (!TryFindInstalledClickOnceAddIn(out installed)) return;
            throw new InvalidOperationException(
                "Máy vẫn còn bản ClickOnce Development " + installed.Version +
                ". Hãy gỡ bản đó trong Apps & Features trước khi chạy bộ cài giao dịch này; " +
                "trình cài không tự gỡ vì ClickOnce không thể được khôi phục an toàn nếu bước sau thất bại.");
        }

        private static bool IsVstoRuntimeInstalled()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (var runtime = machine.OpenSubKey(@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4"))
                {
                    Version installed;
                    if (runtime != null && Convert.ToInt32(runtime.GetValue("Install", 0)) == 1 &&
                        Version.TryParse(Convert.ToString(runtime.GetValue("Version")), out installed) &&
                        installed >= new Version(10, 0, 50903))
                        return true;
                }
            }
            return false;
        }

        private static void RegisterInstalledDevelopmentManifest(string installDirectory)
        {
            var manifestPath = Path.Combine(installDirectory, "ChuanHoa.AddIn.Vsto.vsto");
            if (!File.Exists(manifestPath))
                throw new FileNotFoundException("Không tìm thấy manifest VSTO sau khi cài.", manifestPath);
            var manifest = new Uri(manifestPath).AbsoluteUri + "|vstolocal";
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var addIn = currentUser.CreateSubKey(
                @"Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto",
                RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                if (addIn == null) throw new InvalidOperationException("Không tạo được đăng ký Word Add-in.");
                addIn.SetValue("FriendlyName", "Chuẩn hóa", RegistryValueKind.String);
                addIn.SetValue("Description", "Chuẩn hóa", RegistryValueKind.String);
                addIn.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                addIn.SetValue("Manifest", manifest, RegistryValueKind.String);
            }
        }

        private static void ClearWordRibbonValidationCache()
        {
            // Office caches the Custom UI validation result by add-in ProgID and
            // Ribbon ID. Replacing a vstolocal payload in place can otherwise leave
            // Word reporting COMAddIn.Connect=true while omitting the custom tab.
            // Word must already be closed (enforced at installer startup), so it is
            // safe to remove only this add-in's cache entry and let Office validate
            // the newly installed Ribbon XML on the next launch.
            foreach (var officeVersion in new[] { "14.0", "15.0", "16.0" })
            {
                using (var currentUser = RegistryKey.OpenBaseKey(
                    RegistryHive.CurrentUser, RegistryView.Default))
                using (var validationCache = currentUser.OpenSubKey(
                    @"Software\Microsoft\Office\" + officeVersion +
                    @"\Common\CustomUIValidationCache", true))
                {
                    if (validationCache != null)
                        validationCache.DeleteValue(WordRibbonValidationValueName, false);
                }
            }
        }

        private static void TryClearWordRibbonValidationCache()
        {
            try { ClearWordRibbonValidationCache(); }
            catch { }
        }

        private static void ClearWordResiliencyAndVstoSolutionMetadata()
        {
            foreach (var officeVersion in new[] { "14.0", "15.0", "16.0" })
            {
                using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
                using (var disabledItems = currentUser.OpenSubKey(
                    @"Software\Microsoft\Office\" + officeVersion + @"\Word\Resiliency\DisabledItems", true))
                {
                    if (disabledItems != null)
                    {
                        foreach (var valueName in disabledItems.GetValueNames())
                        {
                            var bytes = disabledItems.GetValue(valueName) as byte[];
                            if (bytes != null && bytes.Length > 0)
                            {
                                var decoded = Encoding.Unicode.GetString(bytes);
                                if (decoded.IndexOf("chuanhoa", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    disabledItems.DeleteValue(valueName, false);
                                }
                            }
                        }
                    }
                }

                using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
                using (var crashingList = currentUser.OpenSubKey(
                    @"Software\Microsoft\Office\" + officeVersion + @"\Word\Resiliency\CrashingAddinList", true))
                {
                    if (crashingList != null)
                    {
                        foreach (var valueName in crashingList.GetValueNames())
                        {
                            if (valueName.IndexOf("chuanhoa", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                crashingList.DeleteValue(valueName, false);
                            }
                        }
                    }
                }
            }

            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var solutionMetadata = currentUser.OpenSubKey(@"Software\Microsoft\VSTO\SolutionMetadata", true))
            {
                if (solutionMetadata != null)
                {
                    foreach (var valueName in solutionMetadata.GetValueNames())
                    {
                        if (valueName.IndexOf("chuanhoa", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            var subKeyGuid = solutionMetadata.GetValue(valueName) as string;
                            solutionMetadata.DeleteValue(valueName, false);
                            if (!string.IsNullOrWhiteSpace(subKeyGuid))
                            {
                                try
                                {
                                    solutionMetadata.DeleteSubKeyTree(subKeyGuid, false);
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void TryClearWordResiliencyAndVstoSolutionMetadata()
        {
            try { ClearWordResiliencyAndVstoSolutionMetadata(); }
            catch { }
        }

        private static void RegisterInstallerState(string version, string installDirectory,
            string cachedInstallerPath, string trustedPublicKeySha256,
            string signingCertificateSha256, string signingRootCertificateSha256)
        {
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var state = currentUser.CreateSubKey(
                @"Software\ChuanHoa\DevelopmentInstaller",
                RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                if (state == null) throw new InvalidOperationException("Không ghi được trạng thái bản cài Development.");
                state.SetValue("Version", version, RegistryValueKind.String);
                state.SetValue("InstallDirectory", installDirectory, RegistryValueKind.String);
                state.SetValue("CachedInstallerPath", cachedInstallerPath, RegistryValueKind.String);
                state.SetValue("TrustedPublicKeySha256", trustedPublicKeySha256,
                    RegistryValueKind.String);
                state.SetValue("SigningCertificateSha256", signingCertificateSha256,
                    RegistryValueKind.String);
                state.SetValue("SigningRootCertificateSha256",
                    signingRootCertificateSha256, RegistryValueKind.String);
            }
        }

        private static void RegisterAppsFeatures(string version, string installDirectory,
            string cachedInstallerPath)
        {
            var command = "\"" + cachedInstallerPath + "\"";
            var estimatedBytes = Directory.GetFiles(installDirectory, "*", SearchOption.AllDirectories)
                .Sum(path => new FileInfo(path).Length) + new FileInfo(cachedInstallerPath).Length;
            using (var currentUser = RegistryKey.OpenBaseKey(
                RegistryHive.CurrentUser, RegistryView.Default))
            using (var entry = currentUser.CreateSubKey(
                AppsFeaturesRegistryPath, RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                if (entry == null)
                    throw new InvalidOperationException("Cannot register Chuẩn hóa in Apps & Features.");
                entry.SetValue("DisplayName", "Chuẩn hóa - Development Test",
                    RegistryValueKind.String);
                entry.SetValue("DisplayVersion", version, RegistryValueKind.String);
                entry.SetValue("Publisher", "Chuẩn hóa", RegistryValueKind.String);
                entry.SetValue("InstallLocation", installDirectory, RegistryValueKind.String);
                entry.SetValue("DisplayIcon", cachedInstallerPath + ",0", RegistryValueKind.String);
                entry.SetValue("UninstallString", command + " /uninstall", RegistryValueKind.String);
                entry.SetValue("QuietUninstallString", command + " /uninstall /quiet",
                    RegistryValueKind.String);
                entry.SetValue("ModifyPath", command + " /repair", RegistryValueKind.String);
                entry.SetValue("NoModify", 0, RegistryValueKind.DWord);
                entry.SetValue("NoRepair", 0, RegistryValueKind.DWord);
                entry.SetValue("EstimatedSize", Convert.ToInt32(Math.Max(
                    1L, Math.Min(int.MaxValue, (estimatedBytes + 1023L) / 1024L))),
                    RegistryValueKind.DWord);
                entry.SetValue("InstallDate", DateTime.UtcNow.ToString("yyyyMMdd"),
                    RegistryValueKind.String);
            }
        }

        private static void VerifyDirectRegistration(string expectedVersion, string installDirectory,
            string cachedInstallerPath, string trustedPublicKeySha256,
            string signingCertificateSha256, string signingRootCertificateSha256)
        {
            var manifestPath = Path.Combine(installDirectory, "ChuanHoa.AddIn.Vsto.vsto");
            var expectedManifest = new Uri(manifestPath).AbsoluteUri + "|vstolocal";
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var addIn = currentUser.OpenSubKey(@"Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto"))
            using (var state = currentUser.OpenSubKey(@"Software\ChuanHoa\DevelopmentInstaller"))
            using (var appsFeatures = currentUser.OpenSubKey(AppsFeaturesRegistryPath))
            {
                var command = "\"" + cachedInstallerPath + "\"";
                if (addIn == null || state == null || appsFeatures == null ||
                    !string.Equals(Convert.ToString(addIn.GetValue("Manifest")), expectedManifest, StringComparison.OrdinalIgnoreCase) ||
                    Convert.ToInt32(addIn.GetValue("LoadBehavior", 0)) != 3 ||
                    !string.Equals(Convert.ToString(state.GetValue("Version")), expectedVersion, StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(state.GetValue("InstallDirectory")), installDirectory, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Convert.ToString(state.GetValue("CachedInstallerPath")), cachedInstallerPath, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Convert.ToString(state.GetValue("TrustedPublicKeySha256")), trustedPublicKeySha256, StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(state.GetValue("SigningCertificateSha256")), signingCertificateSha256, StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(state.GetValue("SigningRootCertificateSha256")), signingRootCertificateSha256, StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(appsFeatures.GetValue("DisplayVersion")), expectedVersion, StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(appsFeatures.GetValue("InstallLocation")), installDirectory, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Convert.ToString(appsFeatures.GetValue("UninstallString")), command + " /uninstall", StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(appsFeatures.GetValue("QuietUninstallString")), command + " /uninstall /quiet", StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(appsFeatures.GetValue("ModifyPath")), command + " /repair", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Không xác minh được đăng ký trực tiếp của add-in Development.");
            }
            var runningInstaller = Assembly.GetExecutingAssembly().Location;
            if (!File.Exists(cachedInstallerPath) ||
                !string.Equals(ComputeFileSha256(cachedInstallerPath),
                    ComputeFileSha256(runningInstaller), StringComparison.Ordinal))
                throw new InvalidOperationException("Cached installer does not match the running installer.");
            VerifyInstallerSigner(cachedInstallerPath, signingCertificateSha256);
            VerifyStagedPayload(installDirectory, expectedVersion,
                trustedPublicKeySha256, signingCertificateSha256,
                signingRootCertificateSha256);
        }

        private static bool TryFindInstalledClickOnceAddIn(out InstalledClickOnceAddIn installed)
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                using (var uninstall = currentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    if (uninstall == null) continue;
                    foreach (var name in uninstall.GetSubKeyNames())
                    {
                        using (var entry = uninstall.OpenSubKey(name))
                        {
                            if (entry == null || !string.Equals(
                                Convert.ToString(entry.GetValue("DisplayName")),
                                "ChuanHoa.AddIn.Vsto", StringComparison.Ordinal))
                                continue;
                            installed = new InstalledClickOnceAddIn(
                                Convert.ToString(entry.GetValue("DisplayVersion")) ?? string.Empty,
                                Convert.ToString(entry.GetValue("UrlUpdateInfo")) ?? string.Empty);
                            return true;
                        }
                    }
                }
            }
            installed = new InstalledClickOnceAddIn(string.Empty, string.Empty);
            return false;
        }

        private sealed class InstalledClickOnceAddIn
        {
            public InstalledClickOnceAddIn(string version, string manifestUrl)
            {
                Version = version;
                ManifestUrl = manifestUrl;
            }

            public string Version { get; }
            public string ManifestUrl { get; }
        }

        private sealed class RegistryValueSnapshot
        {
            public RegistryValueSnapshot(bool keyExisted,
                IDictionary<string, Tuple<object, RegistryValueKind>> values)
            {
                KeyExisted = keyExisted;
                Values = values;
            }

            public bool KeyExisted { get; }
            public IDictionary<string, Tuple<object, RegistryValueKind>> Values { get; }
        }

        private sealed class FileValueSnapshot
        {
            public FileValueSnapshot(string path)
            {
                Path = path;
                Existed = File.Exists(path);
                Content = Existed ? File.ReadAllBytes(path) : null;
            }

            public string Path { get; }
            public bool Existed { get; }
            public byte[]? Content { get; }

            public void Restore()
            {
                if (!Existed)
                {
                    if (File.Exists(Path)) File.Delete(Path);
                    return;
                }
                var directory = System.IO.Path.GetDirectoryName(Path);
                if (string.IsNullOrEmpty(directory))
                    throw new InvalidOperationException("Cannot restore file snapshot: " + Path);
                Directory.CreateDirectory(directory);
                var temporaryPath = Path + ".rollback-" + Guid.NewGuid().ToString("N");
                File.WriteAllBytes(temporaryPath, Content!);
                try
                {
                    if (File.Exists(Path)) File.Replace(temporaryPath, Path, null);
                    else File.Move(temporaryPath, Path);
                }
                finally
                {
                    if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
                }
            }
        }

        private sealed class InstallTransaction : IDisposable
        {
            private const string AddInRegistryPath =
                @"Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto";
            private const string InstallerStateRegistryPath =
                @"Software\ChuanHoa\DevelopmentInstaller";
            private const string AppsFeaturesRegistryPath =
                @"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChuanHoa.DevelopmentTest";
            private readonly string _baseDirectory;
            private readonly string _stagingDirectory;
            private readonly string _installDirectory;
            private readonly string _previousDirectory;
            private readonly string _priorPreviousDirectory;
            private readonly string _trustedKeyPath;
            private readonly string _trustedKeyBackupPath;
            private readonly string _version;
            private readonly string _cacheRoot;
            private readonly string _cacheVersionDirectory;
            private readonly string _cacheTransactionDirectory;
            private readonly RegistryValueSnapshot _addInRegistrySnapshot;
            private readonly RegistryValueSnapshot _installerStateSnapshot;
            private readonly RegistryValueSnapshot _appsFeaturesRegistrySnapshot;
            private readonly FileValueSnapshot[] _accessCacheSnapshots;
            private readonly List<Tuple<StoreName, string, string>> _addedCertificates =
                new List<Tuple<StoreName, string, string>>();
            private bool _trustedKeyChanged;
            private bool _cacheChanged;
            private bool _activationStarted;
            private bool _activated;
            private bool _hadCurrent;
            private bool _committed;
            private bool _rolledBack;

            public InstallTransaction(string baseDirectory, string stagingDirectory,
                string installDirectory, string version)
            {
                _baseDirectory = baseDirectory;
                _stagingDirectory = stagingDirectory;
                _installDirectory = installDirectory;
                _version = version;
                _previousDirectory = Path.Combine(baseDirectory, "Previous");
                _priorPreviousDirectory = Path.Combine(
                    baseDirectory, "Transaction-" + Guid.NewGuid().ToString("N") +
                    "-prior-previous");
                _trustedKeyPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ChuanHoa", "Development", "trusted-key.xml");
                _trustedKeyBackupPath = Path.Combine(
                    baseDirectory, "Transaction-" + Guid.NewGuid().ToString("N") +
                    "-trusted-key.backup");
                _cacheRoot = GetInstallerCacheRoot();
                _cacheVersionDirectory = Path.Combine(_cacheRoot, version);
                _cacheTransactionDirectory = Path.Combine(_cacheRoot,
                    "Transaction-" + Guid.NewGuid().ToString("N") + "-" + version);
                _addInRegistrySnapshot = CaptureRegistryValues(AddInRegistryPath);
                _installerStateSnapshot = CaptureRegistryValues(InstallerStateRegistryPath);
                _appsFeaturesRegistrySnapshot = CaptureRegistryValues(AppsFeaturesRegistryPath);
                var cacheDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ChuanHoa", "Cache");
                _accessCacheSnapshots = new[]
                {
                    new FileValueSnapshot(Path.Combine(cacheDirectory, "lease.xml")),
                    new FileValueSnapshot(Path.Combine(cacheDirectory, "rules.xml")),
                    new FileValueSnapshot(Path.Combine(cacheDirectory, "server-time.txt"))
                };
            }

            public void TrustDevelopmentCertificate(string rootCertificatePath,
                string expectedRootSha256, string certificatePath,
                string expectedSha256)
            {
                using (var rootCertificate = LoadPinnedDevelopmentCertificate(
                    rootCertificatePath, expectedRootSha256,
                    ExpectedRootCertificateSubject))
                {
                    if (AddCertificate(StoreName.Root, rootCertificatePath,
                            rootCertificate, expectedRootSha256))
                        _addedCertificates.Add(Tuple.Create(
                            StoreName.Root, rootCertificate.Thumbprint,
                            expectedRootSha256));
                }
                using (var certificate = LoadPinnedDevelopmentCertificate(
                    certificatePath, expectedSha256,
                    ExpectedCertificateSubject))
                {
                    if (AddCertificate(StoreName.TrustedPublisher, certificatePath,
                            certificate, expectedSha256))
                        _addedCertificates.Add(Tuple.Create(
                            StoreName.TrustedPublisher, certificate.Thumbprint,
                            expectedSha256));
                }
            }

            public void InstallTrustedDevelopmentKey(string stagingDirectory,
                string expectedSha256)
            {
                var source = Path.Combine(
                    stagingDirectory, "DevelopmentSupport", "trusted-key.xml");
                if (!File.Exists(source))
                    throw new FileNotFoundException(
                        "The Development public trust key is missing.", source);
                if (!string.Equals(ComputeFileSha256(source), expectedSha256,
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "The Development public trust key does not match the pinned SHA-256.");

                var destinationDirectory = Path.GetDirectoryName(_trustedKeyPath);
                if (string.IsNullOrEmpty(destinationDirectory))
                    throw new InvalidOperationException(
                        "The Development trust destination is invalid.");
                Directory.CreateDirectory(destinationDirectory);
                if (File.Exists(_trustedKeyPath))
                    File.Copy(_trustedKeyPath, _trustedKeyBackupPath, false);

                var temporaryPath = _trustedKeyPath + ".new-" + Guid.NewGuid().ToString("N");
                File.Copy(source, temporaryPath, false);
                try
                {
                    if (File.Exists(_trustedKeyPath))
                        File.Replace(temporaryPath, _trustedKeyPath, null);
                    else
                        File.Move(temporaryPath, _trustedKeyPath);
                    _trustedKeyChanged = true;
                }
                finally
                {
                    if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
                }
            }

            public void Activate()
            {
                _activationStarted = true;
                try
                {
                    if (Directory.Exists(_previousDirectory))
                        Directory.Move(_previousDirectory, _priorPreviousDirectory);
                    if (Directory.Exists(_installDirectory))
                    {
                        _hadCurrent = true;
                        Directory.Move(_installDirectory, _previousDirectory);
                    }
                    Directory.Move(_stagingDirectory, _installDirectory);
                    _activated = true;
                }
                catch
                {
                    RestoreDirectoryState();
                    throw;
                }
            }

            public string CacheRunningInstaller(string expectedSignerSha256)
            {
                AssertDirectChildDirectory(_cacheRoot, _cacheVersionDirectory, _version);
                Directory.CreateDirectory(_cacheRoot);
                if (Directory.Exists(_cacheTransactionDirectory))
                    throw new InvalidOperationException("Installer cache transaction already exists.");
                if (Directory.Exists(_cacheVersionDirectory))
                    Directory.Move(_cacheVersionDirectory, _cacheTransactionDirectory);
                _cacheChanged = true;
                Directory.CreateDirectory(_cacheVersionDirectory);
                var destination = Path.Combine(
                    _cacheVersionDirectory, CachedInstallerFileName);
                var source = Assembly.GetExecutingAssembly().Location;
                File.Copy(source, destination, false);
                if (!string.Equals(ComputeFileSha256(source), ComputeFileSha256(destination),
                        StringComparison.Ordinal))
                    throw new InvalidOperationException("Cached installer hash mismatch.");
                VerifyInstallerSigner(destination, expectedSignerSha256);
                return destination;
            }

            public void Commit()
            {
                if (_activationStarted && !_hadCurrent &&
                    Directory.Exists(_priorPreviousDirectory))
                    Directory.Move(_priorPreviousDirectory, _previousDirectory);
                else if (Directory.Exists(_priorPreviousDirectory))
                    Directory.Delete(_priorPreviousDirectory, true);
                if (File.Exists(_trustedKeyBackupPath))
                    File.Delete(_trustedKeyBackupPath);
                if (Directory.Exists(_cacheTransactionDirectory))
                    Directory.Delete(_cacheTransactionDirectory, true);
                _committed = true;
                TryDeleteObsoleteInstallerCaches();
                // Retain Previous as the last known-good payload. A later successful
                // upgrade may rotate this recovery slot only after verifying that
                // Current is the registered, matching release.
            }

            public void Rollback()
            {
                if (_rolledBack || _committed) return;
                var failures = new List<Exception>();
                TryRollback(() => RestoreRegistryValues(
                    AddInRegistryPath, _addInRegistrySnapshot), failures);
                TryRollback(() => RestoreRegistryValues(
                    InstallerStateRegistryPath, _installerStateSnapshot), failures);
                TryRollback(() => RestoreRegistryValues(
                    AppsFeaturesRegistryPath, _appsFeaturesRegistrySnapshot), failures);

                if (_cacheChanged)
                    TryRollback(RestoreInstallerCache, failures);

                if (_activationStarted)
                {
                    TryRollback(RestoreDirectoryState, failures);
                }
                else if (Directory.Exists(_stagingDirectory))
                {
                    TryRollback(() => Directory.Delete(_stagingDirectory, true), failures);
                }

                if (_trustedKeyChanged)
                {
                    TryRollback(() =>
                    {
                        if (File.Exists(_trustedKeyBackupPath))
                        {
                            if (File.Exists(_trustedKeyPath))
                                File.Replace(_trustedKeyBackupPath, _trustedKeyPath, null);
                            else
                                File.Move(_trustedKeyBackupPath, _trustedKeyPath);
                        }
                        else if (File.Exists(_trustedKeyPath))
                        {
                            File.Delete(_trustedKeyPath);
                        }
                    }, failures);
                }
                else if (File.Exists(_trustedKeyBackupPath))
                {
                    TryRollback(() => File.Delete(_trustedKeyBackupPath), failures);
                }

                foreach (var cacheSnapshot in _accessCacheSnapshots)
                    TryRollback(cacheSnapshot.Restore, failures);

                foreach (var addedCertificate in _addedCertificates.AsEnumerable().Reverse())
                    TryRollback(() => RemoveCertificate(
                        addedCertificate.Item1, addedCertificate.Item2,
                        addedCertificate.Item3), failures);

                _rolledBack = true;
                TryClearWordRibbonValidationCache();
                if (failures.Count != 0)
                    throw new AggregateException("Installer rollback failed.", failures);
            }

            public void Dispose()
            {
                if (!_committed && !_rolledBack) Rollback();
            }

            private static RegistryValueSnapshot CaptureRegistryValues(string path)
            {
                using (var currentUser = RegistryKey.OpenBaseKey(
                    RegistryHive.CurrentUser, RegistryView.Default))
                using (var key = currentUser.OpenSubKey(path))
                {
                    if (key == null)
                        return new RegistryValueSnapshot(false,
                            new Dictionary<string, Tuple<object, RegistryValueKind>>(
                                StringComparer.OrdinalIgnoreCase));
                    var values = new Dictionary<string, Tuple<object, RegistryValueKind>>(
                        StringComparer.OrdinalIgnoreCase);
                    foreach (var name in key.GetValueNames())
                    {
                        values[name] = Tuple.Create(
                            key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                            key.GetValueKind(name));
                    }
                    return new RegistryValueSnapshot(true, values);
                }
            }

            private void RestoreDirectoryState()
            {
                if (_activated && Directory.Exists(_installDirectory))
                {
                    var failedDirectory = Path.Combine(
                        _baseDirectory, "Failed-" + Guid.NewGuid().ToString("N"));
                    Directory.Move(_installDirectory, failedDirectory);
                    try { Directory.Delete(failedDirectory, true); }
                    catch { }
                }
                if (_hadCurrent && Directory.Exists(_previousDirectory) &&
                    !Directory.Exists(_installDirectory))
                    Directory.Move(_previousDirectory, _installDirectory);
                if (Directory.Exists(_priorPreviousDirectory) &&
                    !Directory.Exists(_previousDirectory))
                    Directory.Move(_priorPreviousDirectory, _previousDirectory);
                _activated = false;
                _activationStarted = false;
            }

            private void RestoreInstallerCache()
            {
                AssertDirectChildDirectory(_cacheRoot, _cacheVersionDirectory, _version);
                if (Directory.Exists(_cacheVersionDirectory))
                    Directory.Delete(_cacheVersionDirectory, true);
                if (Directory.Exists(_cacheTransactionDirectory))
                    Directory.Move(_cacheTransactionDirectory, _cacheVersionDirectory);
                if (Directory.Exists(_cacheRoot) &&
                    Directory.GetFileSystemEntries(_cacheRoot).Length == 0)
                    Directory.Delete(_cacheRoot);
                _cacheChanged = false;
            }

            private void TryDeleteObsoleteInstallerCaches()
            {
                try
                {
                    if (!Directory.Exists(_cacheRoot)) return;
                    foreach (var directory in Directory.GetDirectories(_cacheRoot))
                    {
                        if (string.Equals(Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar),
                                Path.GetFullPath(_cacheVersionDirectory).TrimEnd(Path.DirectorySeparatorChar),
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        var name = Path.GetFileName(directory);
                        Version parsed;
                        if (!Version.TryParse(name, out parsed) || parsed.ToString(4) != name)
                            continue;
                        AssertDirectChildDirectory(_cacheRoot, directory, name);
                        Directory.Delete(directory, true);
                    }
                }
                catch
                {
                    // A stale cache is harmless and can be removed by a later upgrade.
                }
            }

            private static void RestoreRegistryValues(string path,
                RegistryValueSnapshot snapshot)
            {
                using (var currentUser = RegistryKey.OpenBaseKey(
                    RegistryHive.CurrentUser, RegistryView.Default))
                {
                    currentUser.DeleteSubKeyTree(path, false);
                    if (!snapshot.KeyExisted) return;
                    using (var key = currentUser.CreateSubKey(
                        path, RegistryKeyPermissionCheck.ReadWriteSubTree))
                    {
                        if (key == null)
                            throw new InvalidOperationException(
                                "Cannot restore registry key: " + path);
                        foreach (var value in snapshot.Values)
                            key.SetValue(value.Key, value.Value.Item1, value.Value.Item2);
                    }
                }
            }

            private static void TryRollback(Action action, ICollection<Exception> failures)
            {
                try { action(); }
                catch (Exception exception) { failures.Add(exception); }
            }
        }

        private static void ExtractPayload(string installDirectory)
        {
            using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(PayloadResource))
            {
                if (source == null) throw new InvalidOperationException("Installer payload is missing.");
                using (var archive = new ZipArchive(source, ZipArchiveMode.Read, false))
                {
                    var normalizedRoot = Path.GetFullPath(installDirectory)
                        .TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                    var extracted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var entry in archive.Entries)
                    {
                        var normalizedName = entry.FullName.Replace('\\', '/').TrimStart('/');
                        if (string.IsNullOrEmpty(entry.Name))
                        {
                            if (normalizedName.Contains("../") || normalizedName.Contains(":"))
                                throw new InvalidOperationException(
                                    "Payload chứa đường dẫn thư mục không an toàn.");
                            continue;
                        }
                        if (normalizedName.Contains("../") || normalizedName.Contains(":") ||
                            !extracted.Add(normalizedName))
                            throw new InvalidOperationException(
                                "Payload chứa đường dẫn không an toàn hoặc trùng lặp: " + normalizedName);

                        var targetPath = Path.GetFullPath(Path.Combine(
                            installDirectory, normalizedName.Replace('/', Path.DirectorySeparatorChar)));
                        if (!targetPath.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException(
                                "Payload cố ghi tệp ra ngoài thư mục staging.");
                        var targetDirectory = Path.GetDirectoryName(targetPath);
                        if (string.IsNullOrEmpty(targetDirectory))
                            throw new InvalidOperationException("Payload có đường dẫn đích không hợp lệ.");
                        Directory.CreateDirectory(targetDirectory);
                        using (var input = entry.Open())
                        using (var output = new FileStream(
                            targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                            input.CopyTo(output);
                    }
                }
            }
        }

        private static X509Certificate2 LoadPinnedDevelopmentCertificate(
            string certificatePath, string expectedSha256, string expectedSubject)
        {
            if (!File.Exists(certificatePath)) throw new FileNotFoundException("Development certificate is missing.", certificatePath);
            var certificate = new X509Certificate2(certificatePath);
            if (!string.Equals(certificate.Subject, expectedSubject, StringComparison.Ordinal))
            {
                certificate.Dispose();
                throw new InvalidOperationException("Unexpected Development certificate subject.");
            }
            if (!string.Equals(ComputeSha256(certificate.RawData), expectedSha256,
                    StringComparison.Ordinal))
            {
                certificate.Dispose();
                throw new InvalidOperationException(
                    "Development certificate does not match the pinned SHA-256.");
            }
            return certificate;
        }

        private static bool AddCertificate(StoreName storeName, string certificatePath,
            X509Certificate2 certificate, string expectedSha256)
        {
            if (HasPinnedCertificate(storeName, certificate.Thumbprint, expectedSha256)) return false;
            RunCertificateUtility("-f -user -addstore " + storeName + " \"" +
                certificatePath.Replace("\"", "\"\"") + "\"");
            if (!HasPinnedCertificate(storeName, certificate.Thumbprint, expectedSha256))
                throw new InvalidOperationException(
                    "KhÃ´ng cÃ i Ä‘Æ°á»£c certificate Development Ä‘Ã£ pin vÃ o " + storeName + ".");
            return true;
        }

        private static void EnsureDevelopmentAccess(string installDirectory, string version)
        {
            var verifierPath = Path.Combine(installDirectory, "ChuanHoa.DevelopmentAccessSmoke.exe");
            var startInfo = new ProcessStartInfo
            {
                FileName = verifierPath,
                WorkingDirectory = installDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            startInfo.EnvironmentVariables["CHUANHOA_DEVELOPMENT_TRUST_PATH"] =
                Path.Combine(installDirectory, "DevelopmentSupport", "trusted-key.xml");
            var result = RunCapturedProcess(
                startInfo, 15000, "Xác minh giấy phép Development");
            if (result.ExitCode != 0 ||
                result.StandardOutput.IndexOf(
                    "DEVELOPMENT_ACCESS_SMOKE_PASS", StringComparison.Ordinal) < 0)
                throw new InvalidOperationException(
                    "Không thể kích hoạt giấy phép và gói quy tắc cho đúng phiên bản " +
                    version + ". Hãy khởi động máy chủ Development rồi chạy lại bộ cài.\n\n" +
                    (string.IsNullOrWhiteSpace(result.StandardError)
                        ? result.StandardOutput
                        : result.StandardError));
        }

        private sealed class ProcessResult
        {
            public ProcessResult(int exitCode, string standardOutput,
                string standardError)
            {
                ExitCode = exitCode;
                StandardOutput = standardOutput;
                StandardError = standardError;
            }

            public int ExitCode { get; private set; }
            public string StandardOutput { get; private set; }
            public string StandardError { get; private set; }
        }
    }

    public sealed class RsaSha256SignatureDescription : SignatureDescription
    {
        public RsaSha256SignatureDescription()
        {
            KeyAlgorithm = typeof(RSA).AssemblyQualifiedName;
            DigestAlgorithm = typeof(SHA256CryptoServiceProvider).AssemblyQualifiedName;
            FormatterAlgorithm = typeof(RSAPKCS1SignatureFormatter).AssemblyQualifiedName;
            DeformatterAlgorithm = typeof(RSAPKCS1SignatureDeformatter).AssemblyQualifiedName;
        }

        public override AsymmetricSignatureDeformatter CreateDeformatter(
            AsymmetricAlgorithm key)
        {
            var deformatter = new RSAPKCS1SignatureDeformatter(key);
            deformatter.SetHashAlgorithm("SHA256");
            return deformatter;
        }

        public override AsymmetricSignatureFormatter CreateFormatter(
            AsymmetricAlgorithm key)
        {
            var formatter = new RSAPKCS1SignatureFormatter(key);
            formatter.SetHashAlgorithm("SHA256");
            return formatter;
        }

        public override HashAlgorithm CreateDigest()
        {
            return new SHA256CryptoServiceProvider();
        }
    }
}
