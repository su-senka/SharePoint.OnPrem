using FluentAssertions;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class SharePointOptionsValidationTests
{
    [Fact]
    public void Validate_WhenCoreOptionsAreValid_DoesNotThrow()
    {
        var options = new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.local",
            HttpTimeoutMinutes = 5
        };

        var act = options.Validate;

        act.Should().NotThrow();
    }

    [Fact]
    public void Validate_WhenIdentityOptionsMissingClaimsPrefix_Throws()
    {
        var options = new SharePointIdentityOptions
        {
            Domain = "ACR",
            ClaimsPrefix = string.Empty
        };

        var act = options.Validate;

        act.Should().Throw<InvalidOperationException>()
            .WithMessage("*ClaimsPrefix*");
    }
}

