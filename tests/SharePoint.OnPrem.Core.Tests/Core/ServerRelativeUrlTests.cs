using FluentAssertions;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core.Tests.Core;

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

