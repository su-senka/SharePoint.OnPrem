using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Diagnostics.HealthChecks;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.HealthChecks;

/// <summary>
/// Registers a health check that calls SharePoint as the configured account.
/// </summary>
public static class SharePointHealthChecksBuilderExtensions
{
    /// <summary>
    /// Adds a check that reads the current user and web through <see cref="ISharePointSiteClient"/>. It is healthy when
    /// SharePoint answers and accepts the credentials; the result data contains the login, web URL, and server build.
    /// Requires <c>AddSharePointOnPrem(...)</c>.
    /// </summary>
    public static IHealthChecksBuilder AddSharePointOnPrem(
        this IHealthChecksBuilder builder,
        string name = "sharepoint",
        HealthStatus? failureStatus = null,
        IEnumerable<string>? tags = null,
        TimeSpan? timeout = null)
    {
        ArgumentNullException.ThrowIfNull(builder);

        return builder.Add(new HealthCheckRegistration(
            name,
            sp => new SharePointHealthCheck(sp.GetRequiredService<ISharePointSiteClient>()),
            failureStatus,
            tags,
            timeout));
    }
}

internal sealed class SharePointHealthCheck(ISharePointSiteClient siteClient) : IHealthCheck
{
    public async Task<HealthCheckResult> CheckHealthAsync(HealthCheckContext context, CancellationToken cancellationToken = default)
    {
        var failureStatus = context.Registration.FailureStatus;

        try
        {
            var user = await siteClient.GetCurrentUserAsync(cancellationToken);
            var web = await siteClient.GetWebInfoAsync(cancellationToken);

            var data = new Dictionary<string, object>
            {
                ["login"] = user.LoginName,
                ["webUrl"] = web.Url,
                ["serverVersion"] = web.ServerVersion ?? "unknown",
                ["product"] = web.ProductName ?? "unknown"
            };

            return HealthCheckResult.Healthy($"Connected to '{web.Title}' as {user.LoginName}.", data);
        }
        catch (SharePointUnauthorizedException ex)
        {
            return new HealthCheckResult(
                failureStatus,
                "SharePoint rejected the credentials (401/403). Check the handler configured through configureHttpClient.",
                ex);
        }
        catch (SharePointException ex)
        {
            return new HealthCheckResult(failureStatus, $"SharePoint returned an error: {ex.Message}", ex);
        }
        catch (HttpRequestException ex)
        {
            return new HealthCheckResult(failureStatus, $"SharePoint is unreachable: {ex.Message}", ex);
        }
        catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            return new HealthCheckResult(failureStatus, "SharePoint did not answer within the HTTP timeout.", ex);
        }
    }
}
