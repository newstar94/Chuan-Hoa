using System.Text.Json;
using System.Text.Json.Serialization;

namespace ChuanHoa.Api.Development;

public sealed record DevelopmentAdminUser(
    Guid Id,
    string Email,
    string DisplayName,
    string Status,
    DateTimeOffset CreatedAtUtc,
    DateTimeOffset? AccessEndsAtUtc);

public sealed record DevelopmentAdminOffer(
    Guid Id,
    string Name,
    long AmountVnd,
    DateTimeOffset EffectiveFromUtc,
    DateTimeOffset? EffectiveUntilUtc,
    string Status);

public sealed record DevelopmentTrialSettings(
    DateTimeOffset LaunchStartsAtUtc,
    DateTimeOffset LaunchEndsAtUtc,
    bool LaunchTrialEnabled,
    bool PersonalTrialEnabled,
    int PersonalTrialDays);

public sealed record DevelopmentAdminState(
    IReadOnlyList<DevelopmentAdminUser> Users,
    IReadOnlyList<DevelopmentAdminOffer> Offers,
    DevelopmentTrialSettings Trial,
    DateTimeOffset UpdatedAtUtc);

public sealed record CreateDevelopmentUserRequest(string Email, string DisplayName, DateTimeOffset? AccessEndsAtUtc);
public sealed record SetDevelopmentUserStatusRequest(string Status);
public sealed record CreateDevelopmentOfferRequest(
    string Name, long AmountVnd, DateTimeOffset EffectiveFromUtc, DateTimeOffset? EffectiveUntilUtc, string Status);

public sealed class DevelopmentAdminStore
{
    private readonly object _gate = new();
    private readonly string _path;
    private DevelopmentAdminState _state;

    public DevelopmentAdminStore(IWebHostEnvironment environment, IConfiguration configuration, TimeProvider timeProvider)
    {
        _path = configuration["ChuanHoa:DevelopmentAdminStatePath"] ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ChuanHoa", "Development", "development-admin-state.json");
        _state = Load(timeProvider.GetUtcNow());
    }

    public DevelopmentAdminState Read()
    {
        lock (_gate) return _state;
    }

    public DevelopmentAdminState CreateUser(CreateDevelopmentUserRequest request, DateTimeOffset now)
    {
        if (request is null || string.IsNullOrWhiteSpace(request.Email) || !request.Email.Contains('@'))
            throw new ArgumentException("Email hợp lệ là bắt buộc.", nameof(request));
        lock (_gate)
        {
            if (_state.Users.Any(item => string.Equals(item.Email, request.Email.Trim(), StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("DEVELOPMENT_USER_EMAIL_EXISTS");
            var users = _state.Users.Append(new DevelopmentAdminUser(Guid.NewGuid(), request.Email.Trim(),
                string.IsNullOrWhiteSpace(request.DisplayName) ? request.Email.Trim() : request.DisplayName.Trim(),
                "ACTIVE", now, request.AccessEndsAtUtc)).OrderBy(item => item.Email).ToArray();
            return Save(_state with { Users = users, UpdatedAtUtc = now });
        }
    }

    public DevelopmentAdminState SetUserStatus(Guid id, SetDevelopmentUserStatusRequest request, DateTimeOffset now)
    {
        var status = (request?.Status ?? string.Empty).Trim().ToUpperInvariant();
        if (status is not ("ACTIVE" or "DISABLED" or "REVOKED"))
            throw new ArgumentException("Trạng thái người dùng không hợp lệ.", nameof(request));
        lock (_gate)
        {
            if (!_state.Users.Any(item => item.Id == id)) throw new KeyNotFoundException("DEVELOPMENT_USER_NOT_FOUND");
            var users = _state.Users.Select(item => item.Id == id ? item with { Status = status } : item).ToArray();
            return Save(_state with { Users = users, UpdatedAtUtc = now });
        }
    }

    public DevelopmentAdminState CreateOffer(CreateDevelopmentOfferRequest request, DateTimeOffset now)
    {
        if (request is null || string.IsNullOrWhiteSpace(request.Name) || request.AmountVnd < 0)
            throw new ArgumentException("Tên và giá bán hợp lệ là bắt buộc.", nameof(request));
        var status = request.Status.Trim().ToUpperInvariant();
        if (status is not ("DRAFT" or "SCHEDULED" or "PUBLISHED" or "RETIRED"))
            throw new ArgumentException("Trạng thái bảng giá không hợp lệ.", nameof(request));
        if (request.EffectiveUntilUtc.HasValue && request.EffectiveUntilUtc <= request.EffectiveFromUtc)
            throw new ArgumentException("Thời điểm kết thúc phải sau thời điểm bắt đầu.", nameof(request));
        lock (_gate)
        {
            var offers = _state.Offers.Append(new DevelopmentAdminOffer(Guid.NewGuid(), request.Name.Trim(),
                request.AmountVnd, request.EffectiveFromUtc, request.EffectiveUntilUtc, status))
                .OrderByDescending(item => item.EffectiveFromUtc).ToArray();
            return Save(_state with { Offers = offers, UpdatedAtUtc = now });
        }
    }

    public DevelopmentAdminState SetTrial(DevelopmentTrialSettings request, DateTimeOffset now)
    {
        if (request is null || request.LaunchEndsAtUtc <= request.LaunchStartsAtUtc ||
            request.PersonalTrialDays is < 1 or > 90)
            throw new ArgumentException("Cấu hình trial không hợp lệ.", nameof(request));
        lock (_gate) return Save(_state with { Trial = request, UpdatedAtUtc = now });
    }

    private DevelopmentAdminState Load(DateTimeOffset now)
    {
        try
        {
            if (File.Exists(_path))
            {
                var loaded = JsonSerializer.Deserialize<DevelopmentAdminState>(File.ReadAllText(_path), JsonOptions());
                if (loaded is not null) return loaded;
            }
        }
        catch (JsonException) { /* Development state starts clean when the local file is malformed. */ }
        return new DevelopmentAdminState(Array.Empty<DevelopmentAdminUser>(), Array.Empty<DevelopmentAdminOffer>(),
            new DevelopmentTrialSettings(now, now.AddMonths(1), true, false, 30), now);
    }

    private DevelopmentAdminState Save(DevelopmentAdminState state)
    {
        var directory = Path.GetDirectoryName(_path) ?? throw new InvalidOperationException("Development state path is invalid.");
        Directory.CreateDirectory(directory);
        var temporary = _path + ".new";
        File.WriteAllText(temporary, JsonSerializer.Serialize(state, JsonOptions()));
        if (File.Exists(_path)) File.Move(temporary, _path, true);
        else File.Move(temporary, _path);
        _state = state;
        return state;
    }

    private static JsonSerializerOptions JsonOptions() => new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };
}
