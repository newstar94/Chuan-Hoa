namespace ChuanHoa.Contracts;

public sealed record RegisterAccountRequest(string Email, string Password, string DisplayName, string DeviceThumbprint);
public sealed record LoginAccountRequest(string Email, string Password, string DeviceThumbprint);
public sealed record AccountSessionResponse(Guid UserId, string Email, string DisplayName, DateTimeOffset IssuedAtUtc, DateTimeOffset ExpiresAtUtc, string DeviceThumbprint, string SessionToken);
public sealed record AccountAccessResponse(string AccessMode, AccountSessionResponse? Session, IReadOnlySet<string> Features, DateTimeOffset? ExpiresAtUtc);
public sealed record ActivateVipKeyRequest(string ActivationKey, string DeviceThumbprint);
public sealed record VipActivationResponse(Guid ActivationKeyId, string KeyPrefix, IReadOnlySet<string> Features, int MaxDevices, int UsedDevices, int RemainingDevices, DateTimeOffset? ExpiresAtUtc);
public sealed record CreateVipKeyRequest(Guid ProductId, string[] FeatureCodes, int MaxDevices, DateTimeOffset? ExpiresAtUtc, string Note, string ActorId, Guid CorrelationId);
public sealed record CreatedVipKeyResponse(Guid Id, string ActivationKey, string KeyPrefix, int MaxDevices, DateTimeOffset? ExpiresAtUtc);
public sealed record AdminDeviceMutationRequest(Guid TargetId, string DeviceThumbprint, string ActorId, Guid CorrelationId);
public sealed record AdminKeyMutationRequest(Guid KeyId, string ActorId, Guid CorrelationId);
public sealed record ActivationKeyAdminView(Guid Id, string KeyPrefix, string Status, Guid ProductId, int MaxDevices, int UsedDevices, int RemainingDevices, DateTimeOffset CreatedAtUtc, DateTimeOffset? ExpiresAtUtc, string? CreatedByExternalActor, string? Note);
public sealed record ActivationKeyAdminPage(IReadOnlyList<ActivationKeyAdminView> Items, int Page, int PageSize, long Total);

public interface IAccountAccessStore
{
    Task<AccountSessionResponse> RegisterAsync(RegisterAccountRequest request, CancellationToken cancellationToken);
    Task<AccountSessionResponse> LoginAsync(LoginAccountRequest request, CancellationToken cancellationToken);
    Task LogoutAsync(string token, CancellationToken cancellationToken);
    Task<AccountSessionResponse?> GetSessionAsync(string token, CancellationToken cancellationToken);
    Task<VipActivationResponse> ActivateVipKeyAsync(ActivateVipKeyRequest request, CancellationToken cancellationToken);
}

public interface IActivationKeyAdminStore
{
    Task<CreatedVipKeyResponse> CreateVipKeyAsync(CreateVipKeyRequest request, CancellationToken cancellationToken);
    Task ReleaseKeyDeviceAsync(AdminDeviceMutationRequest request, CancellationToken cancellationToken);
    Task RevokeKeyAsync(AdminKeyMutationRequest request, CancellationToken cancellationToken);
    Task ResetAccountDeviceAsync(AdminDeviceMutationRequest request, CancellationToken cancellationToken);
    Task<ActivationKeyAdminPage> ListAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
}
