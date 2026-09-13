using ChuanHoa.Contracts;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;

namespace ChuanHoa.Api.Controllers;

[ApiController]
[Route("v1/access")]
[EnableRateLimiting("access")]
public sealed class AccountAccessController(IAccountAccessStore store) : ControllerBase
{
    [HttpPost("register")]
    public async Task<ActionResult<AccountSessionResponse>> Register(RegisterAccountRequest request, CancellationToken cancellationToken)
    {
        try { return Ok(await store.RegisterAsync(request, cancellationToken)); }
        catch (ArgumentException ex) { return BadRequest(new { code = ex.Message }); }
        catch (Npgsql.PostgresException ex) when (ex.SqlState == "23505") { return Conflict(new { code = "ACCOUNT_ALREADY_EXISTS" }); }
    }

    [HttpPost("login")]
    public async Task<ActionResult<AccountSessionResponse>> Login(LoginAccountRequest request, CancellationToken cancellationToken)
    {
        try { return Ok(await store.LoginAsync(request, cancellationToken)); }
        catch (UnauthorizedAccessException) { return Unauthorized(new { code = "EMAIL_OR_PASSWORD_INVALID", message = "Email hoặc mật khẩu không đúng." }); }
        catch (InvalidOperationException ex) when (ex.Message == "ACCOUNT_DEVICE_IN_USE") { return Conflict(new { code = ex.Message, message = "Tài khoản này đang được sử dụng trên một thiết bị khác." }); }
    }

    [HttpPost("logout")]
    public async Task<IActionResult> Logout(CancellationToken cancellationToken)
    {
        var token = BearerToken();
        if (token is not null) await store.LogoutAsync(token, cancellationToken);
        return NoContent();
    }

    [HttpGet("session")]
    public async Task<ActionResult<AccountSessionResponse>> Session(CancellationToken cancellationToken)
    {
        var token = BearerToken();
        var session = token is null ? null : await store.GetSessionAsync(token, cancellationToken);
        return session is null ? Unauthorized(new { code = "SESSION_REQUIRED", message = "Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại." }) : Ok(session);
    }

    [HttpPost("activation")]
    public async Task<ActionResult<VipActivationResponse>> Activate(ActivateVipKeyRequest request, CancellationToken cancellationToken)
    {
        try { return Ok(await store.ActivateVipKeyAsync(request, cancellationToken)); }
        catch (UnauthorizedAccessException ex) { return Unauthorized(new { code = ex.Message }); }
        catch (InvalidOperationException ex) when (ex.Message == "ACTIVATION_KEY_DEVICE_LIMIT") { return Conflict(new { code = ex.Message, message = "Mã kích hoạt đã đạt số lượng thiết bị tối đa." }); }
        catch (ArgumentException ex) { return BadRequest(new { code = ex.Message }); }
    }

    private string? BearerToken()
    {
        var value = Request.Headers.Authorization.ToString();
        return value.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? value[7..].Trim() : null;
    }
}
