using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage;

namespace Docx2PDFService.Services;

public interface IBlobStorageService
{
    /// <summary>
    /// Uploads a file to blob storage and returns a non-guessable URL.
    /// </summary>
    Task<string> UploadAsync(Stream content, string fileName, CancellationToken ct = default);
}

public class AzureBlobStorageService : IBlobStorageService
{
    private readonly BlobContainerClient _container;
    private readonly ILogger<AzureBlobStorageService> _logger;
    private readonly TimeSpan _sasExpiry;

    public AzureBlobStorageService(
        IConfiguration configuration,
        ILogger<AzureBlobStorageService> logger)
    {
        _logger = logger;

        var connectionString = configuration["AzureBlob:ConnectionString"]
            ?? throw new InvalidOperationException("AzureBlob:ConnectionString is not configured.");

        var containerName = configuration["AzureBlob:ContainerName"] ?? "pdfs";

        var expiryHours = configuration.GetValue<int?>("AzureBlob:SasExpiryHours") ?? 24;
        _sasExpiry = TimeSpan.FromHours(expiryHours);

        _container = new BlobContainerClient(connectionString, containerName);
    }

    public async Task<string> UploadAsync(Stream content, string fileName, CancellationToken ct = default)
    {
        // Ensure the container exists
        await _container.CreateIfNotExistsAsync(PublicAccessType.None, cancellationToken: ct);

        // Build a non-guessable blob name: GUID prefix + original name
        var blobName = $"{Guid.NewGuid():N}/{fileName}";

        _logger.LogInformation("Uploading blob: {BlobName}", blobName);

        var blob = _container.GetBlobClient(blobName);
        content.Position = 0;
        await blob.UploadAsync(content, new BlobHttpHeaders { ContentType = "application/pdf" }, cancellationToken: ct);

        // Generate a SAS URL valid for the configured expiry period
        var sasUri = blob.GenerateSasUri(
            Azure.Storage.Sas.BlobSasPermissions.Read,
            DateTimeOffset.UtcNow.Add(_sasExpiry));

        _logger.LogInformation("Blob uploaded. SAS URL expires at: {Expiry}", DateTimeOffset.UtcNow.Add(_sasExpiry));

        return sasUri.ToString();
    }
}
