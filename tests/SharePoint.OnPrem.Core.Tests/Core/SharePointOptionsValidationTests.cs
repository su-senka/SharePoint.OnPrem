using AwesomeAssertions;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class SharePointOptionsValidationTests
{
    [Fact]
    public void Validate_WhenCoreOptionsAreValid_DoesNotThrow()
    {
        var options = new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.contoso.local",
            HttpTimeoutMinutes = 5
        };

        var act = options.Validate;

        act.Should().NotThrow();
    }

    [Fact]
    public void Validate_WhenIdentityOptionsUseDefaults_DoesNotThrow()
    {
        var options = new SharePointIdentityOptions();

        var act = options.Validate;

        act.Should().NotThrow();
        options.ClaimsPrefix.Should().Be("i:0#.w|");
    }

    [Fact]
    public void Validate_WhenIdentityDomainContainsBackslash_Throws()
    {
        var options = new SharePointIdentityOptions { Domain = "CONTOSO\\jdoe" };

        var act = options.Validate;

        act.Should().Throw<InvalidOperationException>().WithMessage("*Domain*");
    }

    [Fact]
    public void Validate_WhenSiteBaseUrlIsNotHttp_Throws()
    {
        var options = new SharePointOnPremOptions { SiteBaseUrl = "file:///c:/sharepoint" };

        var act = options.Validate;

        act.Should().Throw<InvalidOperationException>().WithMessage("*http(s)*");
    }
}
