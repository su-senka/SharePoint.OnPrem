// A small command-line tool built on SharePoint.OnPrem. Useful as a first smoke test against a farm:
//
//   dotnet run --project samples/SharePoint.OnPrem.Sample -- https://sharepoint.contoso.local/sites/team whoami
//   ... ls "/sites/team/Shared Documents"
//   ... upload ./big.iso "/sites/team/Shared Documents"
//   ... download "/sites/team/Shared Documents/report.pdf" ./report.pdf
//   ... versions "/sites/team/Shared Documents/report.pdf"
//
// Authenticates as the current Windows user, or with SP_USERNAME / SP_PASSWORD / SP_DOMAIN when set.

using System.Net;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Diagnostics.HealthChecks;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.DependencyInjection;
using SharePoint.OnPrem.HealthChecks;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: sp-onprem <siteUrl> <whoami|health|ls|upload|download|versions> [args]");
    return 2;
}

var siteUrl = args[0];
var command = args[1];

var services = new ServiceCollection();
services.AddSharePointOnPrem(
    configureCore: o =>
    {
        o.SiteBaseUrl = siteUrl;
        o.HttpTimeoutMinutes = 30;
    },
    configureHttpClient: http => http.ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler { Credentials = GetCredentials() }));
services.AddHealthChecks().AddSharePointOnPrem();

await using var provider = services.BuildServiceProvider();

try
{
    switch (command)
    {
        case "whoami":
        {
            var site = provider.GetRequiredService<ISharePointSiteClient>();
            var user = await site.GetCurrentUserAsync();
            var web = await site.GetWebInfoAsync();
            Console.WriteLine($"User:    {user.LoginName} ({user.Title}){(user.IsSiteAdmin ? ", site admin" : "")}");
            Console.WriteLine($"Web:     {web.Title} <{web.Url}>");
            Console.WriteLine($"Server:  {web.ProductName ?? "unknown"} ({web.ServerVersion ?? "no version header"})");
            break;
        }
        case "health":
        {
            var report = await provider.GetRequiredService<HealthCheckService>().CheckHealthAsync();
            foreach (var (name, entry) in report.Entries)
                Console.WriteLine($"{name}: {entry.Status} - {entry.Description}");
            return report.Status == HealthStatus.Healthy ? 0 : 1;
        }
        case "ls":
        {
            var folders = provider.GetRequiredService<ISharePointFolderClient>();
            var folder = Arg(2, "folder");
            foreach (var sub in await folders.GetFoldersAsync(folder))
                Console.WriteLine($"{"<dir>",12}  {"",16}  {sub.Name}/");
            foreach (var file in await folders.GetFilesAsync(folder))
                Console.WriteLine($"{file.Length,12:N0}  {file.TimeLastModified:yyyy-MM-dd HH:mm}  {file.Name}");
            break;
        }
        case "upload":
        {
            var localPath = Arg(2, "local file");
            var folder = Arg(3, "target folder");
            await using var stream = File.OpenRead(localPath);
            var total = stream.Length;
            var progress = new SyncProgress(sent => Console.Write($"\r{sent * 100 / Math.Max(total, 1),3}%  {sent:N0} / {total:N0} bytes"));

            await provider.GetRequiredService<ISharePointFolderClient>().EnsureFolderPathAsync(folder);
            var stored = await provider.GetRequiredService<ISharePointFileClient>()
                .UploadChunkedAsync(new UploadFileRequest(folder, Path.GetFileName(localPath), stream), progress: progress);
            Console.WriteLine($"\nUploaded to {stored.ServerRelativeUrl}");
            break;
        }
        case "download":
        {
            var url = Arg(2, "server-relative file URL");
            var localPath = args.Length > 3 ? args[3] : Path.GetFileName(url);
            await using var download = await provider.GetRequiredService<ISharePointFileClient>().DownloadAsync(url);
            await using var target = File.Create(localPath);
            await download.Content.CopyToAsync(target);
            Console.WriteLine($"Saved {target.Length:N0} bytes to {localPath}");
            break;
        }
        case "versions":
        {
            var url = Arg(2, "server-relative file URL");
            foreach (var version in await provider.GetRequiredService<ISharePointFileClient>().GetVersionsAsync(url))
                Console.WriteLine($"{version.VersionLabel,6}  {version.Created:yyyy-MM-dd HH:mm}  {version.Size,12:N0}  {version.CheckInComment}");
            break;
        }
        default:
            Console.Error.WriteLine($"Unknown command '{command}'.");
            return 2;
    }

    return 0;
}
catch (SharePointException ex)
{
    Console.Error.WriteLine($"{ex.GetType().Name}: {ex.Message}");
    if (ex.CorrelationId is not null)
        Console.Error.WriteLine($"Look up SPRequestGuid {ex.CorrelationId} in the ULS logs for details.");
    return 1;
}
catch (HttpRequestException ex)
{
    Console.Error.WriteLine($"Cannot reach {siteUrl}: {ex.Message}");
    return 1;
}

string Arg(int index, string name)
    => args.Length > index ? args[index] : throw new ArgumentException($"Missing argument: {name}.");

static ICredentials GetCredentials()
{
    var user = Environment.GetEnvironmentVariable("SP_USERNAME");
    return string.IsNullOrEmpty(user)
        ? CredentialCache.DefaultCredentials
        : new NetworkCredential(user, Environment.GetEnvironmentVariable("SP_PASSWORD"), Environment.GetEnvironmentVariable("SP_DOMAIN"));
}

// Progress<T> posts to the thread pool, which can print out of order; report synchronously instead.
sealed class SyncProgress(Action<long> report) : IProgress<long>
{
    public void Report(long value) => report(value);
}
