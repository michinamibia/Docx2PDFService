using System.Text.Json.Serialization;

namespace Docx2PDFService.Models;

public class ConversionRequest
{
    [JsonPropertyName("docx_path")]
    public string? DocxPath { get; set; }

    [JsonPropertyName("pdf_path")]
    public string? PdfPath { get; set; }

    [JsonPropertyName("fields")]
    public Dictionary<string, string> Fields { get; set; } = new();

    [JsonPropertyName("created_at")]
    public string? CreatedAt { get; set; }
}
