using AwesomeAssertions;
using Xunit;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class DefaultLoginNameNormalizerTests
{
    [Fact]
    public void Normalize_WhenBareLogin_IsExpandedToClaimsLogin()
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions
        {
            Domain = "CONTOSO",
            ClaimsPrefix = "i:0#.w|"
        });

        var result = sut.Normalize("jdoe");

        result.Should().Be("i:0#.w|CONTOSO\\jdoe");
    }

    [Theory]
    [InlineData("i:0#.w|CONTOSO\\jdoe")]
    [InlineData("c:0+.w|s-1-5-21-1004")]
    [InlineData("jdoe@example.local")]
    public void Normalize_WhenInputAlreadyQualified_ReturnsInput(string login)
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions
        {
            Domain = "CONTOSO",
            ClaimsPrefix = "i:0#.w|"
        });

        var result = sut.Normalize(login);

        result.Should().Be(login);
    }

    [Fact]
    public void Normalize_WhenDomainQualifiedLogin_AddsClaimsPrefix()
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions
        {
            Domain = "CONTOSO",
            ClaimsPrefix = "i:0#.w|"
        });

        var result = sut.Normalize("OTHER\\jdoe");

        result.Should().Be("i:0#.w|OTHER\\jdoe");
    }

    [Fact]
    public void Normalize_WhenClaimsPrefixEmpty_LeavesDomainQualifiedLoginUnchanged()
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions
        {
            Domain = "CONTOSO",
            ClaimsPrefix = string.Empty
        });

        sut.Normalize("CONTOSO\\jdoe").Should().Be("CONTOSO\\jdoe");
        sut.Normalize("jdoe").Should().Be("CONTOSO\\jdoe");
    }

    [Fact]
    public void Normalize_WhenBareLoginAndNoDomainConfigured_Throws()
    {
        var sut = new DefaultLoginNameNormalizer(new SharePointIdentityOptions());

        var act = () => sut.Normalize("jdoe");

        act.Should().Throw<InvalidOperationException>().WithMessage("*Domain*");
    }
}
