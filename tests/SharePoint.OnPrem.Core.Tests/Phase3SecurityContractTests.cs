using System.Net;
using FluentAssertions;
using SharePoint.OnPrem.Security;

namespace SharePoint.OnPrem.Core.Tests;

public class SharePointSecurityClientTests
{
    [Fact]
    public async Task EnsureGroupAsync_WhenGroupAlreadyExists_ReturnsExistingIdWithoutCreateCall()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":42}"));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        var result = await sut.EnsureGroupAsync("PP Readers", "ignored");

        result.Should().Be(42);
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Uri.Should().Contain("sitegroups/getbyname('PP%20Readers')");
    }

    [Fact]
    public async Task EnsureGroupAsync_WhenCreateConflicts_ResolvesExistingGroupId()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Conflict),
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":77}"));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        var result = await sut.EnsureGroupAsync("PP Writers", "writers");

        result.Should().Be(77);
        handler.Requests.Should().HaveCount(4);
        handler.Requests[2].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task EnsureUserAsync_NormalizesBareLoginBeforeCallingEnsureUser()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":15}"));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        var result = await sut.EnsureUserAsync("jnovak");

        result.Should().Be(15);
        handler.Requests[1].Body.Should().Contain("i:0#.w|ACR\\\\jnovak");
    }

    [Fact]
    public async Task GetGroupMembersAsync_ReturnsLoginNamesFromValueArray()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"value\":[{\"LoginName\":\"i:0#.w|ACR\\\\jnovak\"},{\"LoginName\":\"i:0#.w|ACR\\\\svoboda\"}]}"));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        var result = await sut.GetGroupMembersAsync("PP Readers");

        result.Should().BeEquivalentTo(["i:0#.w|ACR\\jnovak", "i:0#.w|ACR\\svoboda"]);
    }

    [Fact]
    public async Task AddUsersToGroupAsync_NormalizesAndAddsEachMissingUser()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.Conflict));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.AddUsersToGroupAsync("PP Readers", ["jnovak", "jnovak", "ACR\\svoboda"]);

        var addRequests = handler.Requests.Where(r => r.Uri.Contains("/users")).ToList();
        addRequests.Should().HaveCount(2);
        addRequests[0].Body.Should().Contain("i:0#.w|ACR\\\\jnovak");
        addRequests[1].Body.Should().Contain("ACR\\\\svoboda");
    }

    [Fact]
    public async Task RemoveUsersFromGroupAsync_EncodesNormalizedLoginNames()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.RemoveUsersFromGroupAsync("PP Readers", ["jnovak"]);

        handler.Requests[1].Uri.Should().Contain("removebyloginname(@v)");
        handler.Requests[1].Uri.Should().Contain("i%3A0%23.w%7CACR%5Cjnovak");
    }

    [Fact]
    public async Task SyncGroupMembershipAsync_AddsMissingAndRemovesExtraMembers()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"value\":[{\"LoginName\":\"i:0#.w|ACR\\\\existing\"},{\"LoginName\":\"i:0#.w|ACR\\\\obsolete\"}]}"),
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.SyncGroupMembershipAsync("PP Readers", ["existing", "new-user"]);

        handler.Requests.Should().HaveCount(4);
        handler.Requests[2].Body.Should().Contain("i:0#.w|ACR\\\\new-user");
        handler.Requests[3].Uri.Should().Contain("obsolete");
    }

    [Fact]
    public async Task BreakInheritanceAsync_UsesFolderListItemEndpointWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.BreakInheritanceAsync("/sites/pp/Attachments/2026", copyRoleAssignments: true);

        handler.Requests[1].Uri.Should().Contain("breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task ResetInheritanceAsync_UsesResetEndpoint()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.ResetInheritanceAsync("/sites/pp/Attachments/2026");

        handler.Requests[1].Uri.Should().Contain("resetroleinheritance");
    }

    [Fact]
    public async Task BindRoleToFolderAsync_EnsuresGroupFetchesRoleIdAndBindsAssignment()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":55}"),
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":1073741826}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.BindRoleToFolderAsync("/sites/pp/Attachments/2026", "PP Readers", "Read");

        handler.Requests.Should().HaveCount(6);
        handler.Requests[4].Uri.Should().Contain("roledefinitions/getbytype(2)");
        handler.Requests[5].Uri.Should().Contain("addroleassignment(principalid=55,roledefid=1073741826)");
    }

    [Fact]
    public async Task RemoveRoleFromFolderAsync_WhenGroupDoesNotExist_DoesNothing()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.RemoveRoleFromFolderAsync("/sites/pp/Attachments/2026", "Missing Group");

        handler.Requests.Should().HaveCount(1);
    }

    [Fact]
    public async Task RemoveRoleFromFolderAsync_WhenGroupExists_SendsDeleteRoleAssignmentRequest()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":55}"),
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.RemoveRoleFromFolderAsync("/sites/pp/Attachments/2026", "PP Readers");

        handler.Requests[2].Uri.Should().Contain("roleassignments/getbyprincipalid(55)");
        handler.Requests[2].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("DELETE");
        handler.Requests[2].Headers["IF-MATCH"].Should().ContainSingle().Which.Should().Be("*");
    }

    [Fact]
    public async Task DeleteGroupAsync_WhenGroupExists_UsesRemoveById()
    {
        var handler = new QueueMessageHandler(
            _ => Phase2TestHelpers.JsonResponse("{\"Id\":88}"),
            _ => Phase2TestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = Phase2TestHelpers.CreateHttpClient(handler);
        var sut = Phase2TestHelpers.CreateSecurityClient(client);

        await sut.DeleteGroupAsync("PP Readers");

        handler.Requests[2].Uri.Should().Contain("sitegroups/removebyid(88)");
    }
}

