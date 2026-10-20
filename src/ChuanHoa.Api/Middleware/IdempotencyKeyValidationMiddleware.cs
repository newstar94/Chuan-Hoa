using System.Text.RegularExpressions;
using ChuanHoa.Contracts;

namespace ChuanHoa.Api.Middleware;

public sealed partial class IdempotencyKeyValidationMiddleware(RequestDelegate next)
{
    public const string ItemKey = "ChuanHoa.IdempotencyKey";

    public async Task InvokeAsync(HttpContext context)
    {
        if (!RequiresIdempotencyKey(context.Request))
        {
            await next(context);
            return;
        }

        var values = context.Request.Headers[ApiHeaders.IdempotencyKey];
        if (values.Count != 1 || string.IsNullOrWhiteSpace(values[0]))
        {
            await ApiErrorWriter.WriteAsync(
                context,
                StatusCodes.Status400BadRequest,
                "IDEMPOTENCY_KEY_REQUIRED",
                "Idempotency key is required",
                "This mutation request did not contain exactly one Idempotency-Key header.",
                "Create one stable key for this logical mutation and reuse it only when retrying the same request.");
            return;
        }

        var idempotencyKey = values[0]!;
        if (!IdempotencyKeyPattern().IsMatch(idempotencyKey))
        {
            await ApiErrorWriter.WriteAsync(
                context,
                StatusCodes.Status400BadRequest,
                "IDEMPOTENCY_KEY_INVALID",
                "Idempotency key is invalid",
                "Idempotency-Key must contain 16 to 128 ASCII letters, digits, dot, underscore, colon, or hyphen characters.",
                "Generate a new opaque key that matches the documented format.");
            return;
        }

        context.Items[ItemKey] = idempotencyKey;
        await next(context);
    }

    private static bool RequiresIdempotencyKey(HttpRequest request)
    {
        if (!request.Path.StartsWithSegments("/v1"))
        {
            return false;
        }

        return HttpMethods.IsPost(request.Method) ||
               HttpMethods.IsPut(request.Method) ||
               HttpMethods.IsPatch(request.Method) ||
               HttpMethods.IsDelete(request.Method);
    }

    [GeneratedRegex("^[A-Za-z0-9._:-]{16,128}$", RegexOptions.CultureInvariant)]
    private static partial Regex IdempotencyKeyPattern();
}
