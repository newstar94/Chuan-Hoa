using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ChuanHoa.DevelopmentTestBootstrapper
{
    internal static class Program
    {
        private const string Title = "Chu\u1EA9n h\u00F3a - C\u00E0i \u0111\u1EB7t th\u1EED nghi\u1EC7m";
        private const string PayloadResource = "ChuanHoa.DevelopmentInstaller.Payload.zip";
        private const string VersionResource = "ChuanHoa.DevelopmentInstaller.Version.txt";
        private const string ExpectedCertificateSubject = "CN=Chuan Hoa Local Development";
        private const string WordRibbonValidationValueName =
            "ChuanHoa.AddIn.Vsto.Microsoft.Word.Document";

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

            try
            {
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

                var version = ReadTextResource(VersionResource).Trim();
                UninstallClickOnceDevelopmentAddIn(baseDirectory);
                var installDirectory = Path.Combine(baseDirectory, "Current");
                var stagingDirectory = Path.Combine(baseDirectory, "Staging-" + Guid.NewGuid().ToString("N"));
                PrepareStagingDirectory(baseDirectory, stagingDirectory);
                ExtractPayload(stagingDirectory);

                var certificatePath = Path.Combine(stagingDirectory, "ChuanHoa.LocalDevelopment.Public.cer");
                TrustDevelopmentCertificate(certificatePath);
                InstallTrustedDevelopmentKey(stagingDirectory);
                EnsureRequiredRuntimes();
                VerifyStagedPayload(stagingDirectory);
                EnsureDevelopmentAccess(stagingDirectory, version);

                ActivateStagingDirectory(baseDirectory, stagingDirectory, installDirectory);
                ClearWordRibbonValidationCache();
                ClearWordResiliencyAndVstoSolutionMetadata();
                RegisterInstalledDevelopmentManifest(installDirectory);
                RegisterInstallerState(version, installDirectory);
                VerifyDirectRegistration(version, installDirectory);

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
            var certificateThumbprints = ReadOwnedDevelopmentCertificateThumbprints(baseDirectory);
            RemoveDirectRegistrationIfOwned(baseDirectory);
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
                currentUser.DeleteSubKeyTree(@"Software\ChuanHoa\DevelopmentInstaller", false);

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
            foreach (var thumbprint in certificateThumbprints)
            {
                RemoveCertificate(StoreName.Root, thumbprint);
                RemoveCertificate(StoreName.TrustedPublisher, thumbprint);
            }
            ClearWordRibbonValidationCache();
        }

        private static string[] ReadOwnedDevelopmentCertificateThumbprints(string baseDirectory)
        {
            var result = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in new[] { "Current", "Previous" })
            {
                var path = Path.Combine(baseDirectory, name, "ChuanHoa.LocalDevelopment.Public.cer");
                if (!File.Exists(path)) continue;
                var certificate = new X509Certificate2(path);
                if (string.Equals(certificate.Subject, ExpectedCertificateSubject, StringComparison.Ordinal) &&
                    !string.IsNullOrWhiteSpace(certificate.Thumbprint))
                    result.Add(certificate.Thumbprint);
            }
            return new System.Collections.Generic.List<string>(result).ToArray();
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

        private static void RemoveCertificate(StoreName storeName, string thumbprint)
        {
            using (var store = new X509Store(storeName, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                var matches = store.Certificates.Find(
                    X509FindType.FindByThumbprint, thumbprint, false);
                foreach (X509Certificate2 certificate in matches) store.Remove(certificate);
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

        private static void PrepareStagingDirectory(string baseDirectory, string stagingDirectory)
        {
            Directory.CreateDirectory(baseDirectory);
            var normalizedBase = Path.GetFullPath(baseDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            var normalizedTarget = Path.GetFullPath(stagingDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!normalizedTarget.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Unsafe installer extraction directory.");
            Directory.CreateDirectory(stagingDirectory);
        }

        private static void VerifyStagedPayload(string stagingDirectory)
        {
            foreach (var required in new[]
            {
                "ChuanHoa.AddIn.Vsto.vsto",
                "ChuanHoa.AddIn.Vsto.dll.manifest",
                "ChuanHoa.AddIn.Vsto.dll",
                "ChuanHoa.Client.Core.dll",
                "ChuanHoa.DevelopmentAccessSmoke.exe",
                "ChuanHoa.LocalDevelopment.Public.cer"
            })
            {
                if (!File.Exists(Path.Combine(stagingDirectory, required)))
                    throw new InvalidOperationException("Bộ cài thiếu tệp bắt buộc: " + required);
            }
        }

        private static void ActivateStagingDirectory(string baseDirectory, string stagingDirectory,
            string installDirectory)
        {
            var previousDirectory = Path.Combine(baseDirectory, "Previous");
            if (Directory.Exists(previousDirectory)) Directory.Delete(previousDirectory, true);
            if (Directory.Exists(installDirectory)) Directory.Move(installDirectory, previousDirectory);
            try
            {
                Directory.Move(stagingDirectory, installDirectory);
            }
            catch
            {
                if (!Directory.Exists(installDirectory) && Directory.Exists(previousDirectory))
                    Directory.Move(previousDirectory, installDirectory);
                throw;
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
            process.WaitForExit();
            if (process.ExitCode != 0)
                throw new InvalidOperationException(
                    "Không gỡ được bản Chuẩn hóa Development " + installed.Version +
                    " (mã " + process.ExitCode + ").");
            InstalledClickOnceAddIn remaining;
            if (TryFindInstalledClickOnceAddIn(out remaining))
                throw new InvalidOperationException(
                    "Windows vẫn ghi nhận bản Chuẩn hóa Development " + remaining.Version + " sau khi gỡ.");
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

        private static void RegisterInstallerState(string version, string installDirectory)
        {
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var state = currentUser.CreateSubKey(
                @"Software\ChuanHoa\DevelopmentInstaller",
                RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                if (state == null) throw new InvalidOperationException("Không ghi được trạng thái bản cài Development.");
                state.SetValue("Version", version, RegistryValueKind.String);
                state.SetValue("InstallDirectory", installDirectory, RegistryValueKind.String);
            }
        }

        private static void VerifyDirectRegistration(string expectedVersion, string installDirectory)
        {
            var manifestPath = Path.Combine(installDirectory, "ChuanHoa.AddIn.Vsto.vsto");
            var expectedManifest = new Uri(manifestPath).AbsoluteUri + "|vstolocal";
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            using (var addIn = currentUser.OpenSubKey(@"Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto"))
            using (var state = currentUser.OpenSubKey(@"Software\ChuanHoa\DevelopmentInstaller"))
            {
                if (addIn == null || state == null ||
                    !string.Equals(Convert.ToString(addIn.GetValue("Manifest")), expectedManifest, StringComparison.OrdinalIgnoreCase) ||
                    Convert.ToInt32(addIn.GetValue("LoadBehavior", 0)) != 3 ||
                    !string.Equals(Convert.ToString(state.GetValue("Version")), expectedVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Không xác minh được đăng ký trực tiếp của add-in Development.");
            }
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

        private static void ExtractPayload(string installDirectory)
        {
            var zipPath = Path.Combine(installDirectory, "payload.zip");
            using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(PayloadResource))
            {
                if (source == null) throw new InvalidOperationException("Installer payload is missing.");
                using (var destination = File.Create(zipPath)) source.CopyTo(destination);
            }
            ZipFile.ExtractToDirectory(zipPath, installDirectory);
            File.Delete(zipPath);
        }

        private static void TrustDevelopmentCertificate(string certificatePath)
        {
            if (!File.Exists(certificatePath)) throw new FileNotFoundException("Development certificate is missing.", certificatePath);
            var certificate = new X509Certificate2(certificatePath);
            if (!string.Equals(certificate.Subject, ExpectedCertificateSubject, StringComparison.Ordinal))
                throw new InvalidOperationException("Unexpected Development certificate subject.");
            AddCertificate(StoreName.Root, certificate);
            AddCertificate(StoreName.TrustedPublisher, certificate);
        }

        private static void AddCertificate(StoreName storeName, X509Certificate2 certificate)
        {
            using (var store = new X509Store(storeName, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                var existing = store.Certificates.Find(X509FindType.FindByThumbprint, certificate.Thumbprint, false);
                if (existing.Count == 0) store.Add(certificate);
            }
        }

        private static void InstallTrustedDevelopmentKey(string installDirectory)
        {
            var source = Path.Combine(installDirectory, "DevelopmentSupport", "trusted-key.xml");
            if (!File.Exists(source)) throw new FileNotFoundException("The Development public trust key is missing.", source);
            var destinationDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ChuanHoa",
                "Development");
            Directory.CreateDirectory(destinationDirectory);
            File.Copy(source, Path.Combine(destinationDirectory, "trusted-key.xml"), true);

        }

        private static void EnsureDevelopmentAccess(string installDirectory, string version)
        {
            var verifierPath = Path.Combine(installDirectory, "ChuanHoa.DevelopmentAccessSmoke.exe");
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = verifierPath,
                WorkingDirectory = installDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            if (process == null)
                throw new InvalidOperationException("Không khởi động được bước xác minh giấy phép Development.");
            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            if (!process.WaitForExit(15000))
            {
                try { process.Kill(); } catch { }
                throw new TimeoutException("Xác minh giấy phép Development quá thời gian 15 giây.");
            }
            if (process.ExitCode != 0 ||
                output.IndexOf("DEVELOPMENT_ACCESS_SMOKE_PASS", StringComparison.Ordinal) < 0)
                throw new InvalidOperationException(
                    "Không thể kích hoạt giấy phép và gói quy tắc cho đúng phiên bản " +
                    version + ". Hãy khởi động máy chủ Development rồi chạy lại bộ cài.\n\n" +
                    (string.IsNullOrWhiteSpace(error) ? output : error));
        }
    }
}
