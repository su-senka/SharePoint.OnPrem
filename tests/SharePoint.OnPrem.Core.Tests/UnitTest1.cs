using FluentAssertions;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core.Tests;

public class DefaultLoginNameNormalizerTests
{
    [Fact]
    public void Normalize_WhenBareLogin_IsExpandedToClaimsLogin()
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions
        {
            Domain = "ACR",
            ClaimsPrefix = "i:0#.w|"
        });

        var result = sut.Normalize("jnovak");

        result.Should().Be("i:0#.w|ACR\\jnovak");
    }

    [Theory]
    [InlineData("i:0#.w|ACR\\jnovak")]
    [InlineData("ACR\\jnovak")]
    [InlineData("jnovak@example.local")]
    public void Normalize_WhenInputAlreadyQualified_ReturnsInput(string login)
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions
        {
            Domain = "ACR",
            ClaimsPrefix = "i:0#.w|"
        });

        var result = sut.Normalize(login);

        result.Should().Be(login);
    }
}

public class ServerRelativeUrlTests
{
    [Fact]
    public void Combine_WhenBaseAndRelativeAreValid_ReturnsCombinedServerRelativePath()
    {
        var result = ServerRelativeUrl.Combine("/sites/pp/Attachments", "3255/26/001");

        result.Should().Be("/sites/pp/Attachments/3255/26/001");
    }

    [Fact]
    public void Validate_WhenPathDoesNotStartWithSlash_ThrowsValidationException()
    {
        var act = () => ServerRelativeUrl.Validate("sites/pp/Attachments");

        act.Should().Throw<SharePointValidationException>()
            .WithMessage("*must start with '/'*");
    }

    [Fact]
    public void Validate_WhenPathLooksEncoded_ThrowsValidationException()
    {
        var act = () => ServerRelativeUrl.Validate("/sites/pp/Shared%20Documents");

        act.Should().Throw<SharePointValidationException>()
            .WithMessage("*must not already be URL encoded*");
    }
}

public class BaseFolderPathScopeTests
{
    [Fact]
    public void ToServerRelativePath_WhenBaseFolderConfigured_ScopesThePath()
    {
        var sut = new BaseFolderPathScope(new SharePointStorageScopeOptions
        {
            BaseFolderServerRelativeUrl = "/sites/pp/Attachments"
        });

        var result = sut.ToServerRelativePath("3255/26/001/003");

        result.Should().Be("/sites/pp/Attachments/3255/26/001/003");
    }
}

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
