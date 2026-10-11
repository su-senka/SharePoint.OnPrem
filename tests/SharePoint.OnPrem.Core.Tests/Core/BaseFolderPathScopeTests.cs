using AwesomeAssertions;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class BaseFolderPathScopeTests
{
    [Fact]
    public void ToServerRelativePath_WhenBaseFolderConfigured_ScopesThePath()
    {
        var sut = new BaseFolderPathScope(new SharePointStorageScopeOptions
        {
            BaseFolderServerRelativeUrl = "/sites/team/Attachments"
        });

        var result = sut.ToServerRelativePath("customers/1042/2026/q1");

        result.Should().Be("/sites/team/Attachments/customers/1042/2026/q1");
    }
}

