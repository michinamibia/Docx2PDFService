using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Docx2PDFService.Services;

public interface IPdfConversionService
{
    /// <summary>
    /// Converts the DOCX at <paramref name="docxPath"/> to a PDF.
    /// Returns the path to the generated PDF file.
    /// </summary>
    Task<string> ConvertToPdfAsync(string docxPath, CancellationToken ct = default);
}

public class LibreOfficePdfConversionService : IPdfConversionService
{
    private readonly ILogger<LibreOfficePdfConversionService> _logger;
    private readonly string _libreOfficePath;

    public LibreOfficePdfConversionService(
        IConfiguration configuration,
        ILogger<LibreOfficePdfConversionService> logger)
    {
        _logger = logger;
        _libreOfficePath = ResolveLibreOfficePath(configuration);
    }

    public async Task<string> ConvertToPdfAsync(string docxPath, CancellationToken ct = default)
    {
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("DOCX file not found.", docxPath);

        var outDir = Path.GetDirectoryName(docxPath)
                     ?? Path.GetTempPath();

        var expectedPdf = Path.Combine(
            outDir,
            Path.GetFileNameWithoutExtension(docxPath) + ".pdf");

        // Delete any stale PDF so we can detect a fresh conversion
        if (File.Exists(expectedPdf))
            File.Delete(expectedPdf);

        var args = $"--headless --convert-to pdf --outdir \"{outDir}\" \"{docxPath}\"";

        _logger.LogInformation(
            "Running LibreOffice conversion: {Executable} {Args}",
            _libreOfficePath, args);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName               = _libreOfficePath,
                Arguments              = args,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                UseShellExecute        = false,
                CreateNoWindow         = true
            }
        };

        process.Start();

        var stdOut = await process.StandardOutput.ReadToEndAsync(ct);
        var stdErr = await process.StandardError.ReadToEndAsync(ct);

        await process.WaitForExitAsync(ct);

        if (process.ExitCode != 0)
        {
            _logger.LogError(
                "LibreOffice exited with code {Code}.\nstdout: {Out}\nstderr: {Err}",
                process.ExitCode, stdOut, stdErr);
            throw new InvalidOperationException(
                $"LibreOffice conversion failed (exit code {process.ExitCode}): {stdErr}");
        }

        if (!File.Exists(expectedPdf))
            throw new InvalidOperationException(
                $"LibreOffice finished successfully but the expected PDF was not found at: {expectedPdf}");

        _logger.LogInformation("PDF created: {Path}", expectedPdf);
        return expectedPdf;
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    private static string ResolveLibreOfficePath(IConfiguration cfg)
    {
        // Allow override via appsettings or environment variable
        var configured = cfg["LibreOffice:ExecutablePath"];
        if (!string.IsNullOrWhiteSpace(configured))
            return configured;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            // Common Windows installation paths
            string[] candidates =
            [
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ];

            foreach (var c in candidates)
                if (File.Exists(c)) return c;

            // Fall back to PATH lookup
            return "soffice.exe";
        }

        // Linux / macOS
        return "libreoffice";
    }
}
