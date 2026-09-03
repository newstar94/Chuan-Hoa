using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using ChuanHoa.Client.Core.Licensing;
using ChuanHoa.Client.Core.Rules;
using ChuanHoa.Client.Core.Security;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    public sealed class LocalAccessManager : IDisposable
    {
        public const string FormatFeature = "FORMAT_SCAN";
        public const string SpellingFeature = "SPELLING_SCAN";
        public const string DocumentToolsFeature = "DOCUMENT_TOOLS";
        public const string AutoFixFeature = "AUTOFIX";
        private const string DevelopmentKeyId = "CHUANHOA-LOCAL-DEVELOPMENT-1";
        private readonly string _clientReleaseId;
        private readonly string _cacheDirectory;
        private readonly string _deviceThumbprint;
        private readonly HttpClient _httpClient;
        private readonly object _validatedCacheGate = new object();
        private ValidatedCache? _validatedCache;
        private Exception? _cacheLoadError;
        private int _refreshInProgress;
        private Exception? _lastRefreshError;

        public event EventHandler? CacheStateChanged;

        public LocalAccessManager(string clientReleaseId)
        {
            _clientReleaseId = clientReleaseId;
            _cacheDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChuanHoa", "Cache");
            Directory.CreateDirectory(_cacheDirectory);
            _deviceThumbprint = LoadOrCreateDeviceThumbprint();
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            LoadValidatedCacheFromDisk();
        }

        public LocalRulePack GetRulePack(string requiredFeature)
        {
            LocalRulePack pack;
            Exception cacheError;
            if (TryGetValidatedCache(requiredFeature, out pack, out cacheError)) return pack;
            // A Ribbon command always runs on Word's UI thread. Never perform network I/O
            // here: even a normal HTTP timeout makes Word appear as Not Responding. Ask the
            // existing worker refresh lane to recover the cache and fail this invocation
            // immediately. The command can be retried after CacheStateChanged re-enables it.
            WarmUp();
            throw new InvalidOperationException(
                "Chưa có giấy phép hoặc gói quy tắc cục bộ hợp lệ. " +
                "Ứng dụng đang làm mới ở chế độ nền; hãy thử lại sau ít giây.\n\n" +
                "Chi tiết: " + cacheError.Message,
                cacheError);
        }

        public string DescribeStatus()
        {
            LocalRulePack pack;
            Exception cacheError;
            if (TryGetValidatedCache(FormatFeature, out pack, out cacheError))
            {
                return "Gói quy tắc: " + pack.PackId + " " + pack.Version + "\nHiệu lực đến: " +
                    pack.ExpiresAtUtc.ToLocalTime().ToString("g", CultureInfo.CurrentCulture) + "\nThiết bị: " + _deviceThumbprint.Substring(0, 12) + "…";
            }

            var refreshError = _lastRefreshError;
            return "Chưa có giấy phép/gói quy tắc hợp lệ.\n" + cacheError.Message +
                (refreshError == null ? string.Empty : "\nLần làm mới gần nhất: " + refreshError.Message);
        }

        public bool HasCachedFeature(string requiredFeature)
        {
            LocalRulePack ignored;
            Exception ignoredError;
            return TryGetValidatedCache(requiredFeature, out ignored, out ignoredError);
        }

        public void WarmUp()
        {
            if (Interlocked.CompareExchange(ref _refreshInProgress, 1, 0) != 0) return;

            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    Refresh();
                    _lastRefreshError = null;
                }
                catch (Exception exception)
                {
                    // A valid signed cache remains usable while the background refresh is unavailable.
                    _lastRefreshError = exception;
                }
                finally
                {
                    Interlocked.Exchange(ref _refreshInProgress, 0);
                    var handler = CacheStateChanged;
                    if (handler != null) handler(this, EventArgs.Empty);
                }
            });
        }

        private bool TryGetValidatedCache(string requiredFeature, out LocalRulePack pack, out Exception error)
        {
            try
            {
                ValidatedCache cache;
                lock (_validatedCacheGate)
                {
                    cache = _validatedCache ?? throw _cacheLoadError ??
                        new InvalidOperationException("Chưa có cache giấy phép đã xác minh.");
                }

                var now = DateTimeOffset.UtcNow;
                new OfflineLeaseValidator().Validate(cache.Lease, _deviceThumbprint, _clientReleaseId,
                    requiredFeature, now, cache.TrustedServerTimeUtc);
                if (now < cache.RulePack.NotBeforeUtc)
                    throw new InvalidOperationException("RULE_PACK_NOT_ACTIVE");
                if (now >= cache.RulePack.ExpiresAtUtc)
                    throw new InvalidOperationException("RULE_PACK_EXPIRED");
                pack = cache.RulePack;
                error = new InvalidOperationException("No error.");
                return true;
            }
            catch (Exception exception)
            {
                pack = null!;
                error = exception;
                return false;
            }
        }

        private void LoadValidatedCacheFromDisk()
        {
            try
            {
                var verifier = CreateVerifier();
                var leasePayload = verifier.Verify(
                    File.ReadAllText(Path.Combine(_cacheDirectory, "lease.xml")), "offlineLease");
                var rulesPayload = verifier.Verify(
                    File.ReadAllText(Path.Combine(_cacheDirectory, "rules.xml")), "rulePack");
                var lease = OfflineLeaseParser.Parse(leasePayload);
                var now = DateTimeOffset.UtcNow;
                var trustedTime = ReadTrustedServerTime();
                new OfflineLeaseValidator().Validate(lease, _deviceThumbprint, _clientReleaseId,
                    FormatFeature, now, trustedTime);
                var pack = LocalRulePackParser.Parse(rulesPayload, now, _clientReleaseId);
                SetValidatedCache(lease, pack, trustedTime);
            }
            catch (Exception exception)
            {
                lock (_validatedCacheGate)
                {
                    _cacheLoadError = exception;
                }
            }
        }

        private DateTimeOffset? ReadTrustedServerTime()
        {
            var trustedPath = Path.Combine(_cacheDirectory, "server-time.txt");
            DateTimeOffset parsed;
            return File.Exists(trustedPath) && DateTimeOffset.TryParseExact(
                File.ReadAllText(trustedPath), "O", CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed)
                ? parsed
                : (DateTimeOffset?)null;
        }

        private void SetValidatedCache(OfflineLease lease, LocalRulePack pack,
            DateTimeOffset? trustedServerTimeUtc)
        {
            lock (_validatedCacheGate)
            {
                _validatedCache = new ValidatedCache(lease, pack, trustedServerTimeUtc);
                _cacheLoadError = null;
            }
        }

        private void Refresh()
        {
#if CHUANHOA_DEVELOPMENT
            var baseUrl = Environment.GetEnvironmentVariable("CHUANHOA_DEVELOPMENT_API_URL");
            if (string.IsNullOrWhiteSpace(baseUrl)) baseUrl = "http://127.0.0.1:5206";
            var body = "{\"deviceThumbprint\":\"" + JsonEscape(_deviceThumbprint) +
                "\",\"clientReleaseId\":\"" + JsonEscape(_clientReleaseId) + "\"}";
            using (var content = new StringContent(body, Encoding.UTF8, "application/json"))
            using (var request = new HttpRequestMessage(HttpMethod.Post, baseUrl.TrimEnd('/') + "/v1/development/bootstrap"))
            {
                request.Headers.Add("Idempotency-Key", "development-bootstrap-" + Guid.NewGuid().ToString("N"));
                request.Content = content;
                using (var response = _httpClient.SendAsync(request).GetAwaiter().GetResult())
            {
                var responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode) throw new InvalidOperationException("Development API trả về " + (int)response.StatusCode + ".");
                var data = Deserialize<BootstrapResponse>(responseBody);
                var verifier = CreateVerifier();
                var now = DateTimeOffset.UtcNow;
                var lease = OfflineLeaseParser.Parse(verifier.Verify(data.Lease, "offlineLease"));
                new OfflineLeaseValidator().Validate(lease, _deviceThumbprint, _clientReleaseId, FormatFeature, now);
                var pack = LocalRulePackParser.Parse(
                    verifier.Verify(data.RulePack, "rulePack"), now, _clientReleaseId);
                DateTimeOffset serverTime;
                if (!DateTimeOffset.TryParse(data.ServerTimeUtc, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out serverTime))
                    throw new FormatException("Development API trả serverTimeUtc không hợp lệ.");
                new OfflineLeaseValidator().Validate(lease, _deviceThumbprint, _clientReleaseId,
                    FormatFeature, now, serverTime);
                AtomicWrite("lease.xml", data.Lease);
                AtomicWrite("rules.xml", data.RulePack);
                AtomicWrite("server-time.txt", serverTime.ToString("O", CultureInfo.InvariantCulture));
                SetValidatedCache(lease, pack, serverTime);
            }
            }
#else
            throw new InvalidOperationException("Bản Release chưa được cấu hình endpoint và khóa ký production.");
#endif
        }

        private RsaSha256ArtifactVerifier CreateVerifier()
        {
#if CHUANHOA_DEVELOPMENT
            var path = Environment.GetEnvironmentVariable("CHUANHOA_DEVELOPMENT_TRUST_PATH");
            if (string.IsNullOrWhiteSpace(path))
                path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChuanHoa", "Development", "trusted-key.xml");
            var root = XElement.Load(path, LoadOptions.None);
            var keyId = (string)root.Attribute("keyId");
            if (!string.Equals(keyId, DevelopmentKeyId, StringComparison.Ordinal))
                throw new CryptographicException("Khóa Development không đúng định danh tin cậy.");
            var publicKey = root.Element("RSAKeyValue") ?? throw new FormatException("Khóa Development thiếu RSAKeyValue.");
            return new RsaSha256ArtifactVerifier(keyId, publicKey.ToString(SaveOptions.DisableFormatting));
#else
            throw new CryptographicException("Chưa cài khóa công khai production.");
#endif
        }

        private string LoadOrCreateDeviceThumbprint()
        {
            var path = Path.Combine(_cacheDirectory, "device-id.dat");
            string deviceId;
            if (File.Exists(path))
            {
                var protectedBytes = Convert.FromBase64String(File.ReadAllText(path).Trim());
                var clearBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    Encoding.UTF8.GetBytes("ChuanHoa.DeviceIdentity.v1"),
                    DataProtectionScope.CurrentUser);
                deviceId = Encoding.UTF8.GetString(clearBytes);
            }
            else
            {
                deviceId = Guid.NewGuid().ToString("D");
                var protectedBytes = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(deviceId),
                    Encoding.UTF8.GetBytes("ChuanHoa.DeviceIdentity.v1"),
                    DataProtectionScope.CurrentUser);
                AtomicWrite("device-id.dat", Convert.ToBase64String(protectedBytes));
            }
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(deviceId))).Replace("-", string.Empty);
        }

        private void AtomicWrite(string name, string value)
        {
            var target = Path.Combine(_cacheDirectory, name);
            var temporary = target + ".new";
            File.WriteAllText(temporary, value, new UTF8Encoding(false));
            if (File.Exists(target)) File.Replace(temporary, target, null);
            else File.Move(temporary, target);
        }

        private static string JsonEscape(string value) => value.Replace("\\", "\\\\").Replace("\"", "\\\"");
        private static T Deserialize<T>(string json)
        {
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(stream);
        }

        public void Dispose() { _httpClient.Dispose(); }

        [DataContract]
        private sealed class BootstrapResponse
        {
            [DataMember(Name = "serverTimeUtc")] public string ServerTimeUtc { get; set; } = string.Empty;
            [DataMember(Name = "lease")] public string Lease { get; set; } = string.Empty;
            [DataMember(Name = "rulePack")] public string RulePack { get; set; } = string.Empty;
        }

        private sealed class ValidatedCache
        {
            public ValidatedCache(OfflineLease lease, LocalRulePack rulePack,
                DateTimeOffset? trustedServerTimeUtc)
            {
                Lease = lease;
                RulePack = rulePack;
                TrustedServerTimeUtc = trustedServerTimeUtc;
            }

            public OfflineLease Lease { get; }
            public LocalRulePack RulePack { get; }
            public DateTimeOffset? TrustedServerTimeUtc { get; }
        }
    }
}
