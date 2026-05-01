using Docx2PDFService.Services;

var builder = WebApplication.CreateBuilder(args);

// ── Increase the request body size limit (large DOCX files) ──────────
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 50 * 1024 * 1024; // 50 MB
});

// ── Services ──────────────────────────────────────────────────────────
builder.Services.AddControllers();

builder.Services.AddSingleton<IDocxProcessingService, DocxProcessingService>();
builder.Services.AddSingleton<IPdfConversionService,  LibreOfficePdfConversionService>();

// ── Azure Blob Storage (optional) ─────────────────────────────────────
// Register only when AzureBlob:ConnectionString is configured.
if (!string.IsNullOrWhiteSpace(builder.Configuration["AzureBlob:ConnectionString"]))
{
    builder.Services.AddSingleton<IBlobStorageService, AzureBlobStorageService>();
}

// ── Kestrel port ──────────────────────────────────────────────────────
builder.WebHost.UseUrls("http://0.0.0.0:80");

// ── Pipeline ──────────────────────────────────────────────────────────
var app = builder.Build();

app.UseAuthorization();
app.MapControllers();
app.MapGet("/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));

app.Logger.LogInformation("Docx2PDFService listening on http://0.0.0.0:80");

app.Run();
