using AwesomeAssertions;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class ServerRelativeUrlTests
{
    [Fact]
    public void Combine_WhenBaseAndRelativeAreValid_ReturnsCombinedServerRelativePath()
    {
        var result = ServerRelativeUrl.Combine("/sites/team/Attachments", "customers/1042/2026");

        result.Should().Be("/sites/team/Attachments/customers/1042/2026");
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
        var act = () => ServerRelativeUrl.Validate("/sites/team/Shared%20Documents");

        act.Should().Throw<SharePointValidationException>()
            .WithMessage("*must not already be URL encoded*");
    }

    [Theory]
    [InlineData("/sites/team/Attachments/../Private")]
    [InlineData("/sites/team/./Attachments")]
    public void Validate_WhenPathContainsDotSegments_ThrowsValidationException(string path)
    {
        var act = () => ServerRelativeUrl.Validate(path);

        act.Should().Throw<SharePointValidationException>().WithMessage("*'.' or '..'*");
    }

    [Theory]
    [InlineData("../Private")]
    [InlineData("customers/../../Private")]
    [InlineData("customers\\..\\Private")]
    public void Combine_WhenRelativePathEscapesBase_ThrowsValidationException(string relativePath)
    {
        var act = () => ServerRelativeUrl.Combine("/sites/team/Attachments", relativePath);

        act.Should().Throw<SharePointValidationException>();
    }

    [Fact]
    public void Validate_WhenNameContainsDots_IsAllowed()
    {
        ServerRelativeUrl.Validate("/sites/team/Docs/report..v2.docx").Should().Be("/sites/team/Docs/report..v2.docx");
    }
}
