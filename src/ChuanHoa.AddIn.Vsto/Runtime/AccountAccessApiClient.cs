using System;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    internal sealed class AccountAccessApiClient : IDisposable
    {
        // Production uses the Bidding domain as the single public ingress. An
        // environment override remains available for staging and local QA.
        private const string DefaultBaseUrl = "https://hosodauthau.online/chuan-hoa";
        private readonly HttpClient _http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        private readonly string _baseUrl;

        public AccountAccessApiClient()
        {
            var configuredUrl = Environment.GetEnvironmentVariable("CHUANHOA_API_URL");
            _baseUrl = (string.IsNullOrWhiteSpace(configuredUrl) ? DefaultBaseUrl : configuredUrl).TrimEnd('/');
        }

        public bool IsConfigured { get { return Uri.TryCreate(_baseUrl, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps; } }

        public Task<AccountSessionDto> RegisterAsync(string email, string password, string displayName, string device)
            => PostAsync<AccountSessionDto>("/v1/access/register", new RegisterDto { Email = email, Password = password, DisplayName = displayName, DeviceThumbprint = device });

        public Task<AccountSessionDto> LoginAsync(string email, string password, string device)
            => PostAsync<AccountSessionDto>("/v1/access/login", new LoginDto { Email = email, Password = password, DeviceThumbprint = device });

        public Task<VipActivationDto> ActivateAsync(string key, string device)
            => PostAsync<VipActivationDto>("/v1/access/activation", new ActivationDto { ActivationKey = key, DeviceThumbprint = device });

        public async Task LogoutAsync(string token)
        {
            if (!IsConfigured || string.IsNullOrWhiteSpace(token)) return;
            using (var request = new HttpRequestMessage(HttpMethod.Post, _baseUrl + "/v1/access/logout"))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + token);
                using (var response = await _http.SendAsync(request).ConfigureAwait(false))
                    if (!response.IsSuccessStatusCode && (int)response.StatusCode != 401)
                        throw new InvalidOperationException("Không thể đăng xuất khỏi tài khoản.");
            }
        }

        private async Task<T> PostAsync<T>(string path, object request)
        {
            if (!IsConfigured) throw new InvalidOperationException("Dịch vụ Chuẩn Hóa chưa được cấu hình HTTPS.");
            string json;
            var serializer = new DataContractJsonSerializer(request.GetType());
            using (var stream = new MemoryStream()) { serializer.WriteObject(stream, request); json = Encoding.UTF8.GetString(stream.ToArray()); }
            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            using (var response = await _http.PostAsync(_baseUrl + path, content).ConfigureAwait(false))
            {
                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) throw new InvalidOperationException(FriendlyError(body, (int)response.StatusCode));
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(body)))
                    return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(stream);
            }
        }

        private static string FriendlyError(string body, int status)
        {
            if (body.IndexOf("ACCOUNT_DEVICE_IN_USE", StringComparison.Ordinal) >= 0) return "Tài khoản này đang được sử dụng trên một thiết bị khác.";
            if (body.IndexOf("ACTIVATION_KEY_DEVICE_LIMIT", StringComparison.Ordinal) >= 0) return "Mã kích hoạt đã đạt số lượng thiết bị tối đa.";
            if (body.IndexOf("ACTIVATION_KEY_EXPIRED", StringComparison.Ordinal) >= 0) return "Mã kích hoạt đã hết hạn.";
            if (body.IndexOf("ACTIVATION_KEY_REVOKED", StringComparison.Ordinal) >= 0) return "Mã kích hoạt đã bị thu hồi.";
            if (status == 401) return "Thông tin đăng nhập hoặc mã kích hoạt không hợp lệ.";
            return "Dịch vụ Chuẩn Hóa không thể xử lý yêu cầu.";
        }

        public void Dispose() { _http.Dispose(); }

        [DataContract] private sealed class RegisterDto { [DataMember(Name="email")] public string Email = ""; [DataMember(Name="password")] public string Password = ""; [DataMember(Name="displayName")] public string DisplayName = ""; [DataMember(Name="deviceThumbprint")] public string DeviceThumbprint = ""; }
        [DataContract] private sealed class LoginDto { [DataMember(Name="email")] public string Email = ""; [DataMember(Name="password")] public string Password = ""; [DataMember(Name="deviceThumbprint")] public string DeviceThumbprint = ""; }
        [DataContract] private sealed class ActivationDto { [DataMember(Name="activationKey")] public string ActivationKey = ""; [DataMember(Name="deviceThumbprint")] public string DeviceThumbprint = ""; }
    }

    [DataContract] internal sealed class AccountSessionDto
    {
        [DataMember(Name="email")] public string Email = "";
        [DataMember(Name="displayName")] public string DisplayName = "";
        [DataMember(Name="issuedAtUtc")] public DateTimeOffset IssuedAtUtc;
        [DataMember(Name="expiresAtUtc")] public DateTimeOffset ExpiresAtUtc;
        [DataMember(Name="deviceThumbprint")] public string DeviceThumbprint = "";
        [DataMember(Name="sessionToken")] public string SessionToken = "";
    }

    [DataContract] internal sealed class VipActivationDto
    {
        [DataMember(Name="keyPrefix")] public string KeyPrefix = "";
        [DataMember(Name="expiresAtUtc")] public DateTimeOffset? ExpiresAtUtc;
        [DataMember(Name="remainingDevices")] public int RemainingDevices;
    }
}
