using System.Text.Json;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

internal sealed class SharePointSiteClient(HttpClient httpClient, ISharePointRequestExecutor requestExecutor) : ISharePointSiteClient
{
    private readonly HttpClient _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
    private readonly ISharePointRequestExecutor _requestExecutor = requestExecutor ?? throw new ArgumentNullException(nameof(requestExecutor));

    public async Task<SharePointWebInfo> GetWebInfoAsync(CancellationToken ct = default)
    {
        using var message = new HttpRequestMessage(HttpMethod.Get, "_api/web?$select=Title,Url,ServerRelativeUrl,Language");
        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);

        var serverVersion = response.Headers.TryGetValues("MicrosoftSharePointTeamServices", out var values)
            ? values.FirstOrDefault()
            : null;

        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct));
        var root = document.RootElement;

        return new SharePointWebInfo(
            ReadString(root, "Title"),
            ReadString(root, "Url"),
            ReadString(root, "ServerRelativeUrl"),
            root.TryGetProperty("Language", out var language) && language.TryGetInt32(out var lcid) ? lcid : 0,
            serverVersion);
    }

    public async Task<SharePointUserInfo> GetCurrentUserAsync(CancellationToken ct = default)
    {
        using var message = new HttpRequestMessage(HttpMethod.Get, "_api/web/currentuser?$select=Id,LoginName,Title,Email,IsSiteAdmin");
        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);

        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct));
        var root = document.RootElement;

        var email = ReadString(root, "Email");
        return new SharePointUserInfo(
            root.TryGetProperty("Id", out var id) && id.TryGetInt32(out var userId) ? userId : 0,
            ReadString(root, "LoginName"),
            ReadString(root, "Title"),
            email.Length == 0 ? null : email,
            root.TryGetProperty("IsSiteAdmin", out var isAdmin) && isAdmin.ValueKind == JsonValueKind.True);
    }

    private static string ReadString(JsonElement element, string propertyName)
        => element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;
}
