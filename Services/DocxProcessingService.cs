using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace Docx2PDFService.Services;

public interface IDocxProcessingService
{
    /// <summary>
    /// Loads a DOCX from <paramref name="sourceStream"/>, replaces all
    /// {{key}} placeholders using <paramref name="fields"/>, and writes
    /// the resulting DOCX to <paramref name="destinationStream"/>.
    /// </summary>
    void ReplaceFieldsAndSave(
        Stream sourceStream,
        Stream destinationStream,
        Dictionary<string, string> fields);
}

public class DocxProcessingService : IDocxProcessingService
{
    private static readonly Regex PlaceholderPattern =
        new(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

    public void ReplaceFieldsAndSave(
        Stream sourceStream,
        Stream destinationStream,
        Dictionary<string, string> fields)
    {
        // Copy source into destination so we can edit it in-place
        sourceStream.CopyTo(destinationStream);
        destinationStream.Position = 0;

        using var wordDoc = WordprocessingDocument.Open(destinationStream, isEditable: true);

        // Process main document body
        if (wordDoc.MainDocumentPart?.Document.Body is not null)
            ProcessContainer(wordDoc.MainDocumentPart.Document.Body, fields);

        // Process headers
        foreach (var headerPart in wordDoc.MainDocumentPart?.HeaderParts ?? [])
            ProcessContainer(headerPart.Header, fields);

        // Process footers
        foreach (var footerPart in wordDoc.MainDocumentPart?.FooterParts ?? [])
            ProcessContainer(footerPart.Footer, fields);

        wordDoc.Save();
    }

    // ---------------------------------------------------------------
    // Helpers
    // ---------------------------------------------------------------

    private static void ProcessContainer(OpenXmlElement container, Dictionary<string, string> fields)
    {
        // Collect every paragraph inside this container (body, header, footer, table cells…)
        foreach (var paragraph in container.Descendants<Paragraph>())
            ReplaceInParagraph(paragraph, fields);
    }

    /// <summary>
    /// Merges all Run texts in a paragraph, applies replacements, then
    /// rewrites the runs so the result is a single run whose text is the
    /// replaced value. This handles the common case where Word splits a
    /// placeholder like {{first_name}} across multiple runs.
    /// </summary>
    private static void ReplaceInParagraph(Paragraph paragraph, Dictionary<string, string> fields)
    {
        var runs = paragraph.Descendants<Run>().ToList();
        if (runs.Count == 0) return;

        // Build combined plain text for this paragraph
        var combined = string.Concat(runs.Select(r => r.InnerText));

        // Fast-exit if no placeholder present at all
        if (!combined.Contains("{{")) return;

        // Check whether any of our fields actually appear
        bool hasMatch = false;
        foreach (var key in fields.Keys)
        {
            if (combined.Contains($"{{{{{key}}}}}"))
            {
                hasMatch = true;
                break;
            }
        }
        if (!hasMatch) return;

        // Perform all replacements on the combined string
        foreach (var (key, value) in fields)
            combined = combined.Replace($"{{{{{key}}}}}", value ?? string.Empty);

        // Keep the formatting (RunProperties) of the first run that
        // contains actual characters, fall back to the very first run.
        var templateRun = runs.FirstOrDefault(r => !string.IsNullOrEmpty(r.InnerText)) ?? runs[0];
        var runProps = templateRun.RunProperties?.CloneNode(true) as RunProperties;

        // Remove all existing runs
        foreach (var run in runs)
            run.Remove();

        // Insert a single new run with the replaced text.
        // Preserve explicit newlines (line-breaks in combined text caused
        // by <w:br/> elements are lost at this stage, so we re-add them).
        var newRun = new Run();
        if (runProps is not null)
            newRun.AppendChild(runProps);

        // Split on '\n' to preserve any soft returns that were represented
        // as newline characters when InnerText was collected.
        var segments = combined.Split('\n');
        for (int i = 0; i < segments.Length; i++)
        {
            if (i > 0)
                newRun.AppendChild(new Break());

            var text = new Text(segments[i])
            {
                Space = SpaceProcessingModeValues.Preserve
            };
            newRun.AppendChild(text);
        }

        // Append after the last ParagraphProperties element (if any),
        // otherwise just at the end of the paragraph.
        var pPr = paragraph.GetFirstChild<ParagraphProperties>();
        if (pPr is not null)
            pPr.InsertAfterSelf(newRun);
        else
            paragraph.AppendChild(newRun);
    }
}
