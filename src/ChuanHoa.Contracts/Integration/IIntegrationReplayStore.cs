namespace ChuanHoa.Contracts.Integration;

public interface IIntegrationReplayStore
{
    Task<bool> TryClaimAsync(string clientId, string nonce, DateTimeOffset expiresAtUtc, CancellationToken cancellationToken);
}
