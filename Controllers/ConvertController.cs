using Docx2PDFService.Models;
using Docx2PDFService.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace Docx2PDFService.Controllers;

[ApiController]
[Route("[controller]")]
public class ConvertController : ControllerBase
{
    private readonly IDocxProcessingService _docxService;
    private readonly IPdfConversionService  _pdfService;
    private readonly ILogger<ConvertController> _logger;

    public ConvertController(
        IDocxProcessingService docxService,
        IPdfConversionService  pdfService,
        ILogger<ConvertController> logger)
    {
        _docxService = docxService;
        _pdfService  = pdfService;
        _logger      = logger;

    }

    // Consistent error helper
    private ObjectResult Problem(int status, string message, string? detail = null)
    {
        _logger.LogWarning("Returning {Status}: {Message} — {Detail}", status, message, detail);
        return StatusCode(status, new
        {
            error   = message,
            detail  = detail,
            timestamp = DateTime.UtcNow
        });
    }

    [HttpPost]
    [Consumes("multipart/form-data")]
    [Produces("application/pdf")]
    public async Task<IActionResult> Convert(CancellationToken ct)
    {
        if (!Request.HasFormContentType)
            return Problem(400, "Only multipart/form-data with 'file' and 'fields' is accepted.");

        IFormCollection form;
        try
        {
            form = await Request.ReadFormAsync(ct);
        }
        catch (Exception ex)
        {
            return Problem(400, "Failed to read multipart form.", ex.Message);
        }

        var uploadedFile = form.Files.GetFile("file");
        if (uploadedFile is null)
            return Problem(400, "Form file 'file' is required.");

        var fieldsJson = form["fields"].FirstOrDefault();
        if (string.IsNullOrWhiteSpace(fieldsJson))
            return Problem(400, "Form field 'fields' (JSON object) is required.");

        Dictionary<string, string> fields;
        try
        {
            fields = JsonSerializer.Deserialize<Dictionary<string, string>>(fieldsJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? new Dictionary<string, string>();
        }
        catch (JsonException ex)
        {
            return Problem(400, "Invalid JSON in 'fields'.", ex.Message);
        }

        // Buffer uploaded file
        await using var uploadStream = new MemoryStream();
        await uploadedFile.CopyToAsync(uploadStream, ct);
        uploadStream.Position = 0;

        var tempDir = Path.Combine(Path.GetTempPath(), "docx2pdf_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        _logger.LogInformation("Working directory: {TempDir}", tempDir);

        try
        {
            var tempDocx = Path.Combine(tempDir, "document.docx");

            _logger.LogInformation("Replacing fields: {Fields}", string.Join(", ", fields.Keys));

            // Create the modified DOCX
            await using (var destStream = new FileStream(tempDocx, FileMode.Create, FileAccess.ReadWrite))
            {
                // Reset position just in case
                uploadStream.Position = 0;
                _docxService.ReplaceFieldsAndSave(uploadStream, destStream, fields);
            }

            _logger.LogInformation("Converting to PDF...");
            var pdfPath = await _pdfService.ConvertToPdfAsync(tempDocx, ct);

            if (!System.IO.File.Exists(pdfPath))
                return Problem(500, "PDF conversion produced no output file.", pdfPath);

            var pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath, ct);
            _logger.LogInformation("Conversion successful. PDF size: {Size} bytes", pdfBytes.Length);

            var origName = uploadedFile.FileName ?? "document";
            var baseName = Path.GetFileNameWithoutExtension(origName);
            if (string.IsNullOrWhiteSpace(baseName)) baseName = "document";
            var fileName = baseName + ".pdf";

            return File(pdfBytes, "application/pdf", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Conversion failed");
            return Problem(500, "Conversion failed.", ex.Message);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); }
            catch (Exception ex) { _logger.LogWarning("Failed to clean up {TempDir}: {Error}", tempDir, ex.Message); }
        }
    }
}
