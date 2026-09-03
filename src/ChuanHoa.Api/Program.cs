using ChuanHoa.Api.Errors;
using ChuanHoa.Api.Middleware;
using ChuanHoa.Api.Security;
using ChuanHoa.Api.Development;
using ChuanHoa.Application;
using ChuanHoa.Contracts;
using ChuanHoa.Infrastructure;
using Npgsql;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddProblemDetails();
builder.Services.AddExceptionHandler<ApiExceptionHandler>();
builder.Services.AddControllers();
builder.Services.AddSingleton(TimeProvider.System);
builder.Services.AddSingleton<DevelopmentArtifactIssuer>();
builder.Services.AddSingleton<DevelopmentAdminStore>();
builder.Services
    .AddAuthentication(FailClosedAuthenticationDefaults.Scheme)
    .AddScheme<Microsoft.AspNetCore.Authentication.AuthenticationSchemeOptions, FailClosedAuthenticationHandler>(
        FailClosedAuthenticationDefaults.Scheme,
        _ => { });
builder.Services.AddAuthorization(options =>
{
    options.AddPolicy("DocumentScan", policy =>
    {
        policy.RequireAuthenticatedUser();
        policy.RequireClaim("scope", "document.scan");
    });
});
builder.Services.AddOpenApi("v1", options =>
{
    options.AddDocumentTransformer((document, _, _) =>
    {
        document.Info.Title = "Chuẩn Hóa API";
        document.Info.Version = "v1";
        document.Servers?.Clear();
        return Task.CompletedTask;
    });
});
builder.Services.AddChuanHoaApplication();
var chuanHoaConnectionString = builder.Configuration.GetConnectionString("ChuanHoa");
if (!string.IsNullOrWhiteSpace(chuanHoaConnectionString))
{
    builder.Services.AddChuanHoaInfrastructure(builder.Configuration);
}
else if (!builder.Environment.IsDevelopment())
{
    throw new InvalidOperationException(
        "ConnectionStrings:ChuanHoa is required outside the Development bootstrap environment.");
}

var app = builder.Build();
app.UseMiddleware<CorrelationIdMiddleware>();
app.UseExceptionHandler();
app.UseAuthentication();
app.UseAuthorization();
app.UseStatusCodePages(async statusCodeContext =>
{
    var response = statusCodeContext.HttpContext.Response;
    var mapping = response.StatusCode switch
    {
        StatusCodes.Status404NotFound => (
            "RESOURCE_NOT_FOUND",
            "The requested resource was not found",
            "No API resource matches this request.",
            "Verify the API version, path, and resource identifier before retrying."),
        StatusCodes.Status405MethodNotAllowed => (
            "METHOD_NOT_ALLOWED",
            "The HTTP method is not allowed",
            "This resource does not support the requested HTTP method.",
            "Use the method published in the OpenAPI contract."),
        _ => (
            "HTTP_ERROR",
            "The request could not be completed",
            "The server rejected the request before producing a response body.",
            "Correct the request using the OpenAPI contract and retry.")
    };
    await ApiErrorWriter.WriteAsync(
        statusCodeContext.HttpContext,
        response.StatusCode,
        mapping.Item1,
        mapping.Item2,
        mapping.Item3,
        mapping.Item4);
});
app.UseMiddleware<IdempotencyKeyValidationMiddleware>();
app.MapControllers();
app.MapOpenApi("/openapi/{documentName}.json");
app.MapGet("/development/admin", (HttpContext context, DevelopmentArtifactIssuer issuer) =>
{
    if (!issuer.IsEnabled) return Results.NotFound();
    var remote = context.Connection.RemoteIpAddress;
    if (remote is not null && !System.Net.IPAddress.IsLoopback(remote)) return Results.Forbid();
    using var stream = typeof(DevelopmentArtifactIssuer).Assembly.GetManifestResourceStream(
        "ChuanHoa.Api.DevelopmentAdmin.index.html");
    if (stream is null) return Results.NotFound();
    using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
    return Results.Content(reader.ReadToEnd(), "text/html; charset=utf-8", System.Text.Encoding.UTF8);
}).ExcludeFromDescription();

static HealthResponse CreateHealthResponse() => new(
    ApiContractVersions.HealthV1,
    "ok",
    "ChuanHoa.Api",
    DateTimeOffset.UtcNow);

app.MapGet("/health", CreateHealthResponse)
    .Produces<HealthResponse>()
    .WithName("GetLivenessHealth");

app.MapGet("/health/ready", async (IServiceProvider services, CancellationToken cancellationToken) =>
{
    var dataSource = services.GetService<NpgsqlDataSource>();
    if (dataSource is null)
    {
        return Results.Json(new { schema = "chuanhoa.health.v1", status = "development_no_database" }, statusCode: 503);
    }
    await using var command = dataSource.CreateCommand("SELECT 1;");
    await command.ExecuteScalarAsync(cancellationToken);
    return Results.Ok(CreateHealthResponse());
})
    .Produces<HealthResponse>()
    .Produces<ApiErrorResponse>(StatusCodes.Status500InternalServerError)
    .WithName("GetReadinessHealth");

app.Run();

public partial class Program;
