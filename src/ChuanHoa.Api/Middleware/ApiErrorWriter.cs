using ChuanHoa.Contracts;

namespace ChuanHoa.Api.Middleware;

public static class ApiErrorWriter
{
    public static Task WriteAsync(
        HttpContext context,
        int statusCode,
        string code,
        string title,
        string detail,
        string recovery,
        IReadOnlyDictionary<string, IReadOnlyList<string>>? validationErrors = null)
    {
        if (context.Response.HasStarted)
        {
            return Task.CompletedTask;
        }

        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/problem+json";
        var response = new ApiErrorResponse(
            ApiContractVersions.ErrorV1,
            code,
            title,
            detail,
            recovery,
            statusCode,
            CorrelationIdMiddleware.GetCorrelationId(context),
            context.Request.Path.Value ?? "/",
            validationErrors);
        return context.Response.WriteAsJsonAsync(response);
    }
}
