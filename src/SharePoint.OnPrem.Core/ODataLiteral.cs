namespace SharePoint.OnPrem.Core;

/// <summary>
/// Encodes values placed inside single-quoted OData string literals in request URLs, e.g. <c>GetFolderByServerRelativeUrl('...')</c>.
/// </summary>
internal static class ODataLiteral
{
    // OData escapes a quote inside a string literal by doubling it. SharePoint percent-decodes the URL before it parses
    // the literal, so the doubling must happen before encoding; otherwise names like "O'Brien.docx" break the request.
    public static string Encode(string value) => Uri.EscapeDataString(value.Replace("'", "''"));
}
