# Docx2PDFService

Lightweight .NET web service that receives a DOCX file and a JSON object of fields, replaces placeholders in the DOCX, converts the result to PDF (via LibreOffice), and returns the PDF binary.

## What it does

- Exposes HTTP POST /convert
- Accepts a multipart/form-data request containing:
  - `file` — the DOCX file to process (.docx)
  - `fields` — a JSON object mapping placeholder names to replacement strings (example below)
- Replaces occurrences of `{{key}}` in the DOCX with the provided value (handles placeholders split across Word runs)
- Converts the modified DOCX to PDF using LibreOffice headless CLI
- Returns the PDF binary (`application/pdf`)

## Endpoint

POST /convert
- Content-Type: `multipart/form-data`
- Form fields:
  - `file` — the .docx file upload
  - `fields` — JSON object (e.g. `{"first_name":"John","last_name":"Smith"}`)

Success: 200 with PDF bytes
Errors: JSON object `{ "error": "short message", "detail": "longer detail", "timestamp": "..." }`

## Example: curl (Linux/macOS)

Create a `fields.json`:

```json
{
  "first_name": "John",
  "last_name": "Smith",
  "email": "john.smith@example.com"
}
```

Upload and convert:

```bash
curl -X POST http://localhost:5525/convert \
  -F "file=@/path/to/tpl.docx;type=application/vnd.openxmlformats-officedocument.wordprocessingml.document" \
  -F "fields=@fields.json;type=application/json" \
  -o output.pdf
```

## Example: PowerShell

```powershell
curl -X POST http://localhost:5525/convert `
  -F "file=@C:\path\to\tpl.docx;type=application/vnd.openxmlformats-officedocument.wordprocessingml.document" `
  -F "fields=@C:\path\to\fields.json;type=application/json" `
  -o output.pdf
```

## Health

GET /health
- Returns 200 JSON `{ "status": "healthy", "timestamp": "..." }`

## Docker

The included `Dockerfile` installs LibreOffice and publishes the app.

Build and run with Docker:

```bash
docker build -t docx2pdf .
docker run -d --restart unless-stopped -p 5525:5525 --name docx2pdf docx2pdf
```

Or with docker-compose (provided `docker-compose.yml`):

```bash
docker compose build --no-cache
docker compose up -d --force-recreate
docker compose ps
docker compose logs -f
```

## Troubleshooting

- HTTP 400 responses will include `detail` explaining the problem (e.g. invalid JSON in `fields`, missing form parts).
- Make sure placeholder tokens in the DOCX use double-curly braces exactly: `{{first_name}}`.
- LibreOffice must be available in the container or host. Use `appsettings.json` or environment variable `LibreOffice__ExecutablePath` to override the executable path.
- If conversion fails, check container logs for LibreOffice stderr output.

## Notes

- The service no longer accepts server file paths — it requires the DOCX upload and `fields` JSON.
- The `DocxProcessingService` handles the common Word behavior where placeholders are split across runs.

## License

Unlicensed — adapt as needed for your use-case.
