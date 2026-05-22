using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

/// <summary>
/// Default login-name normalizer for SharePoint on-prem claims-based environments.
/// </summary>
internal sealed class DefaultLoginNameNormalizer(SharePointIdentityOptions options) : ILoginNameNormalizer
{
    private readonly SharePointIdentityOptions _options = options ?? throw new ArgumentNullException(nameof(options));

    public string Normalize(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            throw new ArgumentException("Login name must be provided.", nameof(input));

        var value = input.Trim();

        if (value.StartsWith(_options.ClaimsPrefix, StringComparison.OrdinalIgnoreCase)
            || value.Contains('\\')
            || value.Contains('@'))
        {
            return value;
        }

        return $"{_options.ClaimsPrefix}{_options.Domain}\\{value}";
    }
}

