using ChuanHoa.Api.Development;
using Microsoft.AspNetCore.Mvc;

namespace ChuanHoa.Api.Controllers;

[ApiController]
[Route("v1/development")]
public sealed class DevelopmentBootstrapController : ControllerBase
{
    private readonly DevelopmentArtifactIssuer _issuer;
    private readonly IWebHostEnvironment _environment;

    public DevelopmentBootstrapController(DevelopmentArtifactIssuer issuer, IWebHostEnvironment environment)
    {
        _issuer = issuer;
        _environment = environment;
    }

    [HttpPost("bootstrap")]
    [ProducesResponseType<DevelopmentBootstrapResponse>(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status403Forbidden)]
    public ActionResult<DevelopmentBootstrapResponse> Bootstrap([FromBody] DevelopmentBootstrapRequest request)
    {
        if (!_issuer.IsEnabled) return NotFound();
        var remote = HttpContext.Connection.RemoteIpAddress;
        if (remote is not null && !System.Net.IPAddress.IsLoopback(remote)) return Forbid();
        return Ok(_issuer.Issue(request));
    }
}
