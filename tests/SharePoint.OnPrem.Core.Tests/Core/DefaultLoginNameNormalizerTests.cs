using FluentAssertions;
using Xunit;

namespace SharePoint.OnPrem.Core.Tests.Core;

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

