using ChuanHoa.Application.Scanning;
using ChuanHoa.Contracts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ChuanHoa.Api.Controllers;

[ApiController]
[Route("v1/document-jobs")]
[Authorize(Policy = "DocumentScan")]
public sealed class DocumentScanController : ControllerBase
{
    private readonly TechnicalDocumentScanner _scanner;

    public DocumentScanController(TechnicalDocumentScanner scanner)
    {
        _scanner = scanner ?? throw new ArgumentNullException(nameof(scanner));
    }

    [HttpPost("scan")]
    [ProducesResponseType<ScanResult>(StatusCodes.Status200OK)]
    [ProducesResponseType<ApiErrorResponse>(StatusCodes.Status400BadRequest)]
    [ProducesResponseType<ApiErrorResponse>(StatusCodes.Status401Unauthorized)]
    [ProducesResponseType<ApiErrorResponse>(StatusCodes.Status403Forbidden)]
    public ActionResult<ScanResult> Scan([FromBody] ScanRequest request)
    {
        return Ok(_scanner.Scan(request));
    }
}
