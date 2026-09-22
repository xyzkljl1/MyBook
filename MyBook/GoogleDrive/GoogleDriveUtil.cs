using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using DriveFile = Google.Apis.Drive.v3.Data.File;

namespace MyBook;

// Setup: create a Google Cloud service account and a JSON key, but assign the
// service account no Google Cloud/IAM roles. In Google Drive, share only the
// Reports folder with the key's client_email as Viewer (read-only).
//
// This read-only transport deliberately has no arbitrary file-id operation:
// callers can only use files returned from a direct child of the single
// top-level Reports folder resolved by this instance.
sealed partial class GoogleDriveUtil : IDisposable
{
    private const string CredentialSectionName = "GoogleCloudServeAccountKey";
    private const string ReportsFolderName = "Reports";
    private const string GoogleFolderMimeType = "application/vnd.google-apps.folder";
    private const string GoogleNativeMimePrefix = "application/vnd.google-apps.";
    private const string GoogleTokenUri = "https://oauth2.googleapis.com/token";
    private const string GoogleUniverseDomain = "googleapis.com";
    private const long DefaultMaximumDownloadBytes = 64L * 1024 * 1024;
    private const long AbsoluteMaximumDownloadBytes = 256L * 1024 * 1024;
    private const int MaximumListPages = 100;
    private static readonly TimeSpan RequestTimeout = TimeSpan.FromMinutes(2);
    private static readonly Regex ServiceAccountEmailRegex = new(
        @"\A[^\s@]+@[^\s@]+\.iam\.gserviceaccount\.com\z",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly DriveService service;
    private readonly DatabaseUtil? database;
    private readonly SemaphoreSlim reportsFolderLock = new(1, 1);
    private readonly object scopeToken = new();
    private string? reportsFolderId;
    private bool disposed;

    public GoogleDriveUtil(IConfiguration config, DatabaseUtil? database = null)
    {
        ArgumentNullException.ThrowIfNull(config);
        this.database = database;
        var credential = CreateCredential(config.GetSection(CredentialSectionName));
        service = new DriveService(new BaseClientService.Initializer
        {
            ApplicationName = "MyBook",
            HttpClientInitializer = credential
        });
        service.HttpClient.Timeout = RequestTimeout;
    }

    public static bool IsConfigured(IConfiguration config)
    {
        ArgumentNullException.ThrowIfNull(config);
        var section = config.GetSection(CredentialSectionName);
        return new[] { "type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "token_uri" }
            .All(field => !String.IsNullOrWhiteSpace(section[field]));
    }

    public async Task<IReadOnlyList<ReportFile>> ListReportFilesAsync(
        string subfolderName,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        ValidateSubfolderName(subfolderName);

        var reportsId = await GetReportsFolderIdAsync(cancellationToken).ConfigureAwait(false);
        var subfolders = await ListFilesCoreAsync(
            "resolve direct report subfolder",
            $"'{EscapeQueryLiteral(reportsId)}' in parents and name = '{EscapeQueryLiteral(subfolderName)}' "
                + $"and mimeType = '{GoogleFolderMimeType}' and trashed = false",
            "id,name,mimeType,parents,trashed",
            cancellationToken).ConfigureAwait(false);
        if (subfolders.Count != 1)
        {
            throw new GoogleDriveAccessException(
                "resolve direct report subfolder",
                subfolders.Count == 0 ? "folder_not_found" : "folder_ambiguous",
                "Google Drive GET /drive/v3/files could not resolve exactly one direct report subfolder.");
        }

        var subfolder = subfolders[0];
        if (subfolder.Parents is null || !subfolder.Parents.Contains(reportsId, StringComparer.Ordinal))
            throw new GoogleDriveAccessException(
                "validate direct report subfolder",
                "scope_violation",
                "Google Drive returned a report subfolder outside the Reports scope.");

        var files = await ListFilesCoreAsync(
            "list report files",
            $"'{EscapeQueryLiteral(subfolder.Id)}' in parents and trashed = false and mimeType != '{GoogleFolderMimeType}'",
            ReportFileFields,
            cancellationToken).ConfigureAwait(false);
        return files
            .OrderBy(file => file.Name, StringComparer.Ordinal)
            .ThenBy(file => file.Id, StringComparer.Ordinal)
            .Select(file => new ReportFile(scopeToken, subfolder.Id, file))
            .ToList();
    }

    public async Task<DownloadedReportFile> DownloadReportFileAsync(
        ReportFile file,
        long maximumBytes = DefaultMaximumDownloadBytes,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(file);
        if (!ReferenceEquals(file.ScopeToken, scopeToken))
            throw new ArgumentException("The report file was not listed by this Google Drive instance.", nameof(file));
        if (maximumBytes <= 0 || maximumBytes > AbsoluteMaximumDownloadBytes)
            throw new ArgumentOutOfRangeException(nameof(maximumBytes));

        var before = await GetReportFileMetadataAsync(file.Id, "validate report file before download", cancellationToken)
            .ConfigureAwait(false);
        ValidateDownloadMetadata(file, before, maximumBytes, requireListedRevision: true);

        using var output = new MemoryStream(checked((int)before.Size!.Value));
        var request = service.Files.Get(file.Id);
        request.SupportsAllDrives = true;
        IDownloadProgress progress;
        try
        {
            progress = await request.DownloadAsync(output, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw CreateRequestException("download report file", exception);
        }

        if (progress.Status != DownloadStatus.Completed)
            throw CreateRequestException(
                "download report file",
                progress.Exception ?? new IOException("Google Drive media download did not complete."));
        if (output.Length != before.Size.Value)
            throw new GoogleDriveAccessException(
                "validate downloaded report file",
                "size_mismatch",
                "Google Drive GET /drive/v3/files/{fileId}?alt=media returned an unexpected byte count.");

        var after = await GetReportFileMetadataAsync(file.Id, "validate report file after download", cancellationToken)
            .ConfigureAwait(false);
        ValidateDownloadMetadata(file, after, maximumBytes, requireListedRevision: false);
        if (!SameRevision(before, after))
            throw new GoogleDriveAccessException(
                "validate report file after download",
                "revision_changed",
                "The Google Drive report file changed while it was being downloaded.");

        var content = output.ToArray();
        var sha256 = Convert.ToHexString(SHA256.HashData(content)).ToLowerInvariant();
        return new DownloadedReportFile(file, content, sha256);
    }

    private async Task<string> GetReportsFolderIdAsync(CancellationToken cancellationToken)
    {
        if (reportsFolderId is not null)
            return reportsFolderId;

        await reportsFolderLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (reportsFolderId is not null)
                return reportsFolderId;

            var rootRequest = service.Files.Get("root");
            rootRequest.Fields = "id";
            rootRequest.SupportsAllDrives = true;
            var driveRoot = await ExecuteRequestAsync(
                "resolve Drive root",
                token => rootRequest.ExecuteAsync(token),
                cancellationToken).ConfigureAwait(false);

            var candidates = await ListFilesCoreAsync(
                "resolve Reports root",
                $"name = '{ReportsFolderName}' and mimeType = '{GoogleFolderMimeType}' and trashed = false",
                "id,name,mimeType,parents,sharedWithMeTime,trashed",
                cancellationToken).ConfigureAwait(false);
            var topLevel = candidates
                .Where(folder => folder.Parents?.Contains(driveRoot.Id, StringComparer.Ordinal) == true
                    || folder.SharedWithMeTimeDateTimeOffset.HasValue)
                .ToList();
            if (topLevel.Count != 1)
            {
                throw new GoogleDriveAccessException(
                    "resolve Reports root",
                    topLevel.Count == 0 ? "folder_not_found" : "folder_ambiguous",
                    "Google Drive GET /drive/v3/files could not resolve exactly one top-level Reports folder.");
            }

            reportsFolderId = topLevel[0].Id;
            return reportsFolderId;
        }
        finally
        {
            reportsFolderLock.Release();
        }
    }

    private async Task<List<DriveFile>> ListFilesCoreAsync(
        string stage,
        string query,
        string fileFields,
        CancellationToken cancellationToken)
    {
        var files = new List<DriveFile>();
        string? pageToken = null;
        var pageCount = 0;
        do
        {
            if (++pageCount > MaximumListPages)
                throw new GoogleDriveAccessException(
                    stage,
                    "excessive_pagination",
                    "Google Drive GET /drive/v3/files exceeded the bounded page count.");

            var request = service.Files.List();
            request.Q = query;
            request.Spaces = "drive";
            request.PageSize = 1000;
            request.PageToken = pageToken;
            request.OrderBy = "name";
            request.IncludeItemsFromAllDrives = true;
            request.SupportsAllDrives = true;
            request.Fields = $"nextPageToken,incompleteSearch,files({fileFields})";
            var page = await ExecuteRequestAsync(
                stage,
                token => request.ExecuteAsync(token),
                cancellationToken).ConfigureAwait(false);
            if (page.IncompleteSearch == true)
                throw new GoogleDriveAccessException(
                    stage,
                    "incomplete_search",
                    "Google Drive GET /drive/v3/files reported an incomplete search.");

            if (page.Files is not null)
                files.AddRange(page.Files);
            pageToken = String.IsNullOrWhiteSpace(page.NextPageToken) ? null : page.NextPageToken;
        }
        while (pageToken is not null);

        if (files.Select(file => file.Id).Distinct(StringComparer.Ordinal).Count() != files.Count)
            throw new GoogleDriveAccessException(
                stage,
                "duplicate_file_id",
                "Google Drive GET /drive/v3/files returned a duplicate file identifier.");
        return files;
    }

    private async Task<DriveFile> GetReportFileMetadataAsync(
        string fileId,
        string stage,
        CancellationToken cancellationToken)
    {
        var request = service.Files.Get(fileId);
        request.Fields = ReportFileFields;
        request.SupportsAllDrives = true;
        return await ExecuteRequestAsync(stage, token => request.ExecuteAsync(token), cancellationToken)
            .ConfigureAwait(false);
    }

    private static void ValidateDownloadMetadata(
        ReportFile listed,
        DriveFile current,
        long maximumBytes,
        bool requireListedRevision)
    {
        if (current.Trashed == true
            || current.Parents is null
            || !current.Parents.Contains(listed.ParentFolderId, StringComparer.Ordinal))
        {
            throw new GoogleDriveAccessException(
                "validate report file scope",
                "scope_violation",
                "The Google Drive report file is no longer a direct child of its allowed report subfolder.");
        }
        if (current.Capabilities?.CanDownload != true)
            throw new GoogleDriveAccessException(
                "validate report file download capability",
                "download_forbidden",
                "The Google Drive report file is not downloadable by this service account.");
        if (String.IsNullOrWhiteSpace(current.MimeType)
            || current.MimeType.StartsWith(GoogleNativeMimePrefix, StringComparison.Ordinal))
        {
            throw new GoogleDriveAccessException(
                "validate report file type",
                "unsupported_google_workspace_file",
                "Google Workspace native files are not supported by the report downloader.");
        }
        if (!current.Size.HasValue || current.Size.Value < 0 || current.Size.Value > maximumBytes)
            throw new GoogleDriveAccessException(
                "validate report file size",
                "invalid_or_excessive_size",
                "The Google Drive report file has no bounded downloadable size.");
        if (current.Size.Value > Int32.MaxValue)
            throw new GoogleDriveAccessException(
                "validate report file size",
                "unsupported_size",
                "The Google Drive report file exceeds the in-memory downloader limit.");
        if (requireListedRevision && !SameRevision(listed, current))
            throw new GoogleDriveAccessException(
                "validate report file before download",
                "revision_changed",
                "The Google Drive report file changed after it was listed.");
    }

    private static bool SameRevision(ReportFile listed, DriveFile current)
    {
        return listed.Size == current.Size
            && listed.Version == current.Version
            && String.Equals(listed.Md5Checksum, current.Md5Checksum, StringComparison.Ordinal)
            && listed.ModifiedTime == current.ModifiedTimeDateTimeOffset;
    }

    private static bool SameRevision(DriveFile left, DriveFile right)
    {
        return left.Size == right.Size
            && left.Version == right.Version
            && String.Equals(left.Md5Checksum, right.Md5Checksum, StringComparison.Ordinal)
            && left.ModifiedTimeDateTimeOffset == right.ModifiedTimeDateTimeOffset;
    }

    private static ServiceAccountCredential CreateCredential(IConfigurationSection section)
    {
        var type = RequireCredentialValue(section, "type");
        var projectId = RequireCredentialValue(section, "project_id");
        var privateKeyId = RequireCredentialValue(section, "private_key_id");
        var privateKey = RequireCredentialValue(section, "private_key");
        var clientEmail = RequireCredentialValue(section, "client_email");
        _ = RequireCredentialValue(section, "client_id");
        var tokenUri = RequireCredentialValue(section, "token_uri");
        var universeDomain = section["universe_domain"]?.Trim();

        if (!String.Equals(type, "service_account", StringComparison.Ordinal)
            || !String.Equals(tokenUri, GoogleTokenUri, StringComparison.Ordinal)
            || !String.IsNullOrWhiteSpace(universeDomain)
                && !String.Equals(universeDomain, GoogleUniverseDomain, StringComparison.OrdinalIgnoreCase)
            || !ServiceAccountEmailRegex.IsMatch(clientEmail)
            || !privateKey.Contains("-----BEGIN PRIVATE KEY-----", StringComparison.Ordinal)
            || !privateKey.Contains("-----END PRIVATE KEY-----", StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"{CredentialSectionName} is not a supported Google service-account credential.");
        }

        try
        {
            var initializer = new ServiceAccountCredential.Initializer(clientEmail)
            {
                ProjectId = projectId,
                KeyId = privateKeyId,
                Scopes = [DriveService.Scope.DriveReadonly],
                UniverseDomain = String.IsNullOrWhiteSpace(universeDomain) ? GoogleUniverseDomain : universeDomain
            }.FromPrivateKey(privateKey);
            return new ServiceAccountCredential(initializer);
        }
        catch (Exception exception) when (exception is ArgumentException or CryptographicException or FormatException)
        {
            throw new InvalidOperationException(
                $"{CredentialSectionName} contains an invalid private key.",
                exception);
        }
    }

    private static string RequireCredentialValue(IConfigurationSection section, string name)
    {
        var value = section[name];
        if (String.IsNullOrWhiteSpace(value))
            throw new InvalidOperationException($"Missing {CredentialSectionName}.{name} in config.json.");
        return value.Trim();
    }

    private static void ValidateSubfolderName(string value)
    {
        if (String.IsNullOrWhiteSpace(value)
            || value.Length > 128
            || value is "." or ".."
            || value.Contains('/')
            || value.Contains('\\')
            || value.Any(Char.IsControl))
        {
            throw new ArgumentException(
                "A Google Drive report subfolder must be one non-empty direct-child name.",
                nameof(value));
        }
    }

    private static string EscapeQueryLiteral(string value)
    {
        return value.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("'", "\\'", StringComparison.Ordinal);
    }

    private static async Task<T> ExecuteRequestAsync<T>(
        string stage,
        Func<CancellationToken, Task<T>> request,
        CancellationToken cancellationToken)
    {
        try
        {
            return await request(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw CreateRequestException(stage, exception);
        }
    }

    private static GoogleDriveAccessException CreateRequestException(string stage, Exception exception)
    {
        if (exception is GoogleDriveAccessException known)
            return known;

        HttpStatusCode? status = exception switch
        {
            GoogleApiException google => google.HttpStatusCode,
            HttpRequestException request => request.StatusCode,
            _ => null
        };
        var category = status switch
        {
            HttpStatusCode.Unauthorized => "authentication_failed",
            HttpStatusCode.Forbidden => "access_denied_or_api_disabled",
            HttpStatusCode.NotFound => "resource_not_found",
            HttpStatusCode.TooManyRequests => "rate_limited",
            >= HttpStatusCode.InternalServerError => "google_server_error",
            not null => "http_error",
            null when exception is HttpRequestException => "network_error",
            _ => "download_error"
        };
        var protocolResult = status.HasValue ? $"HTTP {(int)status.Value}" : exception.GetType().Name;
        return new GoogleDriveAccessException(
            stage,
            category,
            $"Google Drive request failed at {stage}: GET www.googleapis.com/drive/v3/files, {protocolResult}, category={category}.",
            exception);
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(disposed, this);
    }

    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;
        service.Dispose();
        reportsFolderLock.Dispose();
    }

    private const string ReportFileFields =
        "id,name,mimeType,size,version,md5Checksum,modifiedTime,parents,capabilities(canDownload),trashed";

    internal sealed class ReportFile
    {
        private readonly object scopeToken;

        internal ReportFile(object scopeToken, string parentFolderId, DriveFile file)
        {
            this.scopeToken = scopeToken;
            ParentFolderId = parentFolderId;
            Id = file.Id;
            Name = file.Name;
            MimeType = file.MimeType;
            Size = file.Size;
            Version = file.Version;
            Md5Checksum = file.Md5Checksum;
            ModifiedTime = file.ModifiedTimeDateTimeOffset;
            CanDownload = file.Capabilities?.CanDownload == true;
        }

        internal object ScopeToken => scopeToken;
        internal string ParentFolderId { get; }
        internal string Id { get; }
        public string Name { get; }
        public string MimeType { get; }
        public long? Size { get; }
        public long? Version { get; }
        public string? Md5Checksum { get; }
        public DateTimeOffset? ModifiedTime { get; }
        public bool CanDownload { get; }
        public bool IsGoogleWorkspaceFile => MimeType.StartsWith(GoogleNativeMimePrefix, StringComparison.Ordinal);
    }

    internal sealed record DownloadedReportFile(ReportFile File, byte[] Content, string Sha256);
}

sealed class GoogleDriveAccessException : InvalidOperationException
{
    public GoogleDriveAccessException(string stage, string category, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        Stage = stage;
        Category = category;
    }

    public string Stage { get; }
    public string Category { get; }
}
