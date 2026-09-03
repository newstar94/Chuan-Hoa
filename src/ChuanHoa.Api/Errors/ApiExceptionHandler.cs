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
            "OFFER_NOT_AVAILABLE" => new ErrorMapping(
                StatusCodes.Status404NotFound,
                domainException.Code,
                "No offer is available",
                domainException.Message,
                "Refresh the offer catalog or contact support if an offer should be active."),
            "OFFER_EFFECTIVE_RANGE_OVERLAP" => new ErrorMapping(
                StatusCodes.Status409Conflict,
                domainException.Code,
                "Offer configuration conflicts",
                domainException.Message,
                "Do not continue the commercial mutation. Resolve the overlapping published offers first."),
            "PRODUCT_CODE_REQUIRED" or "PERSONAL_TRIAL_DURATION_INVALID" => new ErrorMapping(
                StatusCodes.Status400BadRequest,
                domainException.Code,
                "The request is invalid",
                domainException.Message,
                "Correct the request using the published API contract and retry."),
            var code when code.StartsWith("GRANT_", StringComparison.Ordinal) => new ErrorMapping(
                StatusCodes.Status403Forbidden,
                domainException.Code,
                "Execution grant was rejected",
                domainException.Message,
                "Request a new execution grant for the current user, device, release, command, and document."),
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
