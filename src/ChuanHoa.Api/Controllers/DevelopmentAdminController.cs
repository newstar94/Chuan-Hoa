using ChuanHoa.Api.Development;
using Microsoft.AspNetCore.Mvc;

namespace ChuanHoa.Api.Controllers;

[ApiController]
[Route("v1/development/admin")]
public sealed class DevelopmentAdminController : ControllerBase
{
    private readonly DevelopmentArtifactIssuer _issuer;
    private readonly DevelopmentAdminStore _store;
    private readonly TimeProvider _timeProvider;

    public DevelopmentAdminController(DevelopmentArtifactIssuer issuer, DevelopmentAdminStore store, TimeProvider timeProvider)
    {
        _issuer = issuer;
        _store = store;
        _timeProvider = timeProvider;
    }

    [HttpGet("state")]
    public ActionResult<DevelopmentAdminState> State()
    {
        var denied = Denied();
        if (denied is not null) return denied;
        return Ok(_store.Read());
    }

    [HttpPost("users")]
    public ActionResult<DevelopmentAdminState> CreateUser([FromBody] CreateDevelopmentUserRequest request)
    {
        var denied = Denied();
        if (denied is not null) return denied;
        return Ok(_store.CreateUser(request, _timeProvider.GetUtcNow()));
    }

    [HttpPatch("users/{id:guid}/status")]
    public ActionResult<DevelopmentAdminState> SetUserStatus(Guid id, [FromBody] SetDevelopmentUserStatusRequest request)
    {
        var denied = Denied();
        if (denied is not null) return denied;
        return Ok(_store.SetUserStatus(id, request, _timeProvider.GetUtcNow()));
    }

    [HttpPost("offers")]
    public ActionResult<DevelopmentAdminState> CreateOffer([FromBody] CreateDevelopmentOfferRequest request)
    {
        var denied = Denied();
        if (denied is not null) return denied;
        return Ok(_store.CreateOffer(request, _timeProvider.GetUtcNow()));
    }

    [HttpPut("trial")]
    public ActionResult<DevelopmentAdminState> SetTrial([FromBody] DevelopmentTrialSettings request)
    {
        var denied = Denied();
        if (denied is not null) return denied;
        return Ok(_store.SetTrial(request, _timeProvider.GetUtcNow()));
    }

    private ActionResult<DevelopmentAdminState>? Denied()
    {
        if (!_issuer.IsEnabled) return NotFound();
        var remote = HttpContext.Connection.RemoteIpAddress;
        if (remote is not null && !System.Net.IPAddress.IsLoopback(remote)) return Forbid();
        return null;
    }
}
