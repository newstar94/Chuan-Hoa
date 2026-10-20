using ChuanHoa.Api.Middleware;
using ChuanHoa.Domain.Common;
using Microsoft.AspNetCore.Diagnostics;

namespace ChuanHoa.Api.Errors;

public sealed class ApiExceptionHandler(ILogger<ApiExceptionHandler> logger) : IExceptionHandler
{
    public async ValueTask<bool> TryHandleAsync(
        HttpContext httpContext,
        Exception exception,
        CancellationToken cancellationToken)
    {
        var mapping = MapException(exception);
        var correlationId = CorrelationIdMiddleware.GetCorrelationId(httpContext);

        if (mapping.StatusCode >= StatusCodes.Status500InternalServerError)
        {
            logger.LogError(exception, "Unhandled API exception. CorrelationId={CorrelationId}", correlationId);
        }
        else
        {
            logger.LogWarning(
                "Handled API exception. Code={Code} CorrelationId={CorrelationId}",
                mapping.Code,
                correlationId);
        }

        await ApiErrorWriter.WriteAsync(
            httpContext,
            mapping.StatusCode,
            mapping.Code,
            mapping.Title,
            mapping.Detail,
            mapping.Recovery);
        return true;
    }

    private static ErrorMapping MapException(Exception exception)
    {
        if (exception is not DomainException domainException)
        {
            return new ErrorMapping(
                StatusCodes.Status500InternalServerError,
                "INTERNAL_ERROR",
                "The request could not be completed",
                "An unexpected server error occurred.",
                "Retry once using the same idempotency key. If the error persists, contact support with the correlation ID.");
        }

        return domainException.Code switch
        {
            _ => new ErrorMapping(
                StatusCodes.Status422UnprocessableEntity,
                domainException.Code,
                "The business rule rejected the request",
                domainException.Message,
                "Refresh authoritative state, correct the request, and retry with a new idempotency key.")
        };
    }

    private sealed record ErrorMapping(
        int StatusCode,
        string Code,
        string Title,
        string Detail,
        string Recovery);
}
