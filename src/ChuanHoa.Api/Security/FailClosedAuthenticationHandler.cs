using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;

namespace ChuanHoa.Api.Security;

public static class FailClosedAuthenticationDefaults
{
    public const string Scheme = "ChuanHoaFailClosed";
}

public sealed class FailClosedAuthenticationHandler : AuthenticationHandler<AuthenticationSchemeOptions>
{
    public FailClosedAuthenticationHandler(
        IOptionsMonitor<AuthenticationSchemeOptions> options,
        ILoggerFactory logger,
        UrlEncoder encoder)
        : base(options, logger, encoder)
    {
    }

    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
        return Task.FromResult(AuthenticateResult.NoResult());
    }
}
