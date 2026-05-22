using FluentAssertions;

namespace SharePoint.OnPrem.Core.Tests.Core;

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

