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
curl -X POST http://localhost/convert \
  -F "file=@/path/to/tpl.docx;type=application/vnd.openxmlformats-officedocument.wordprocessingml.document" \
  -F "fields=@fields.json;type=application/json" \
  -o output.pdf
```

## Example: PowerShell

```powershell
curl -X POST http://localhost/convert `
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
docker run -d --restart unless-stopped -p 80:80 --name docx2pdf docx2pdf
```

Or with docker-compose (provided `docker-compose.yml`):

```bash
docker compose build --no-cache
docker compose up -d --force-recreate
docker compose ps
docker compose logs -f
```

## Azure Blob Storage (optional)

By default, the service returns the PDF bytes directly in the HTTP response.

If you configure Azure Blob Storage credentials, the service uploads the PDF to a private blob container and returns a JSON object containing a time-limited, non-guessable SAS URL instead:

```json
{ "url": "https://<account>.blob.core.windows.net/pdfs/<guid>/document.pdf?..." }
```

### Environment variables

| Variable | Description | Default |
|---|---|---|
| `AzureBlob__ConnectionString` | Azure Storage connection string. Leave blank to disable. | `""` (disabled) |
| `AzureBlob__ContainerName` | Container name to store PDFs in. Created automatically if missing. | `pdfs` |
| `AzureBlob__SasExpiryHours` | How many hours the SAS URL stays valid. | `24` |

### docker-compose example

```yaml
environment:
  - AzureBlob__ConnectionString=DefaultEndpointsProtocol=https;AccountName=<account>;AccountKey=<key>;EndpointSuffix=core.windows.net
  - AzureBlob__ContainerName=pdfs
  - AzureBlob__SasExpiryHours=48
```

### Notes

- The blob is stored with a GUID-prefixed name (`<guid>/<filename>.pdf`), making the URL non-guessable even in a private container.
- The container is created with **private** access. The SAS URL grants read-only access for the configured expiry window.
- When Azure Blob Storage is disabled (no connection string), the API continues to return the PDF binary directly — behaviour is unchanged.

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
