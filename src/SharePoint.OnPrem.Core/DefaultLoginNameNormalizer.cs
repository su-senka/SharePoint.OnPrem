using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

/// <summary>
/// Default login-name normalizer for SharePoint on-prem claims-based environments.
/// </summary>
/// <remarks>
/// <list type="bullet">
/// <item>Claims-encoded logins (containing '|') and UPNs (containing '@') are returned unchanged.</item>
/// <item><c>DOMAIN\user</c> becomes <c>{ClaimsPrefix}DOMAIN\user</c>.</item>
/// <item>A bare <c>user</c> becomes <c>{ClaimsPrefix}{Domain}\user</c>.</item>
/// </list>
/// Producing the same form SharePoint reports for group members is what lets
/// <see cref="ISharePointSecurityClient.SyncGroupMembershipAsync"/> compare desired and current members.
/// </remarks>
internal sealed class DefaultLoginNameNormalizer(SharePointIdentityOptions options) : ILoginNameNormalizer
{
    private readonly SharePointIdentityOptions _options = options ?? throw new ArgumentNullException(nameof(options));

    public string Normalize(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            throw new ArgumentException("Login name must be provided.", nameof(input));

        var value = input.Trim();
        var prefix = _options.ClaimsPrefix ?? string.Empty;

        if ((prefix.Length > 0 && value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            || value.Contains('|')
            || value.Contains('@'))
        {
            return value;
        }

        if (value.Contains('\\'))
            return prefix + value;

        if (string.IsNullOrWhiteSpace(_options.Domain))
            throw new InvalidOperationException(
                $"Cannot normalize bare login '{value}': SharePointIdentityOptions.Domain is not configured. Pass DOMAIN\\user or configure a domain.");

        return $"{prefix}{_options.Domain}\\{value}";
    }
}
