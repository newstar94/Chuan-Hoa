using ChuanHoa.Contracts;
using Microsoft.Extensions.Primitives;

namespace ChuanHoa.Api.Middleware;

public sealed class CorrelationIdMiddleware(RequestDelegate next)
{
    public const string ItemKey = "ChuanHoa.CorrelationId";

    public async Task InvokeAsync(HttpContext context)
    {
        var correlationId = ResolveCorrelationId(context.Request.Headers[ApiHeaders.CorrelationId]);
        context.TraceIdentifier = correlationId;
        context.Items[ItemKey] = correlationId;
        context.Response.Headers[ApiHeaders.CorrelationId] = correlationId;
        await next(context);
    }

    public static string GetCorrelationId(HttpContext context)
    {
        if (context.Items.TryGetValue(ItemKey, out var stored) && stored is string correlationId)
        {
            return correlationId;
        }

        return context.TraceIdentifier;
    }

    private static string ResolveCorrelationId(StringValues requestedValues)
    {
        if (requestedValues.Count == 1 &&
            Guid.TryParseExact(requestedValues[0], "D", out var requestedCorrelationId) &&
            requestedCorrelationId != Guid.Empty)
        {
            return requestedCorrelationId.ToString("D");
        }

        return Guid.NewGuid().ToString("D");
    }
}
