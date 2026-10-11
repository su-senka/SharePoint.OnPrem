using SharePoint.OnPrem.Abstractions;
using System.Net;
using AwesomeAssertions;
using SharePoint.OnPrem.Core.Tests.Contracts;
using SharePoint.OnPrem.Security;

namespace SharePoint.OnPrem.Core.Tests.Security;

public class SharePointSecurityClientTests
{
    [Fact]
    public async Task EnsureGroupAsync_WhenGroupAlreadyExists_ReturnsExistingIdWithoutCreateCall()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":42}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var result = await sut.EnsureGroupAsync("Team Readers", "ignored");

        result.Should().Be(42);
        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Uri.Should().Contain("sitegroups/getbyname('Team%20Readers')");
    }

    [Fact]
    public async Task EnsureGroupAsync_WhenCreateConflicts_ResolvesExistingGroupId()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Conflict),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":77}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var result = await sut.EnsureGroupAsync("Team Writers", "writers");

        result.Should().Be(77);
        handler.Requests.Should().HaveCount(4);
        handler.Requests[2].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task EnsureUserAsync_NormalizesBareLoginBeforeCallingEnsureUser()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":15}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var result = await sut.EnsureUserAsync("jdoe");

        result.Should().Be(15);
        handler.Requests[1].Body.Should().Contain("i:0#.w|CONTOSO\\\\jdoe");
    }

    [Fact]
    public async Task GetGroupMembersAsync_ReturnsLoginNamesFromValueArray()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"value\":[{\"LoginName\":\"i:0#.w|CONTOSO\\\\jdoe\"},{\"LoginName\":\"i:0#.w|CONTOSO\\\\asmith\"}]}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var result = await sut.GetGroupMembersAsync("Team Readers");

        result.Should().BeEquivalentTo(["i:0#.w|CONTOSO\\jdoe", "i:0#.w|CONTOSO\\asmith"]);
    }

    [Fact]
    public async Task AddUsersToGroupAsync_NormalizesAndAddsEachMissingUser()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.Conflict));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.AddUsersToGroupAsync("Team Readers", ["jdoe", "jdoe", "CONTOSO\\asmith"]);

        var addRequests = handler.Requests.Where(r => r.Uri.Contains("/users")).ToList();
        addRequests.Should().HaveCount(2);
        addRequests[0].Body.Should().Contain("i:0#.w|CONTOSO\\\\jdoe");
        addRequests[1].Body.Should().Contain("CONTOSO\\\\asmith");
    }

    [Fact]
    public async Task RemoveUsersFromGroupAsync_EncodesNormalizedLoginNames()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.RemoveUsersFromGroupAsync("Team Readers", ["jdoe"]);

        handler.Requests[1].Uri.Should().Contain("removebyloginname(@v)");
        handler.Requests[1].Uri.Should().Contain("i%3A0%23.w%7CCONTOSO%5Cjdoe");
    }

    [Fact]
    public async Task SyncGroupMembershipAsync_AddsMissingAndRemovesExtraMembers()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"value\":[{\"LoginName\":\"i:0#.w|CONTOSO\\\\existing\"},{\"LoginName\":\"i:0#.w|CONTOSO\\\\obsolete\"}]}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.SyncGroupMembershipAsync("Team Readers", ["existing", "new-user"]);

        handler.Requests.Should().HaveCount(4);
        handler.Requests[2].Body.Should().Contain("i:0#.w|CONTOSO\\\\new-user");
        handler.Requests[3].Uri.Should().Contain("obsolete");
    }

    [Fact]
    public async Task SyncGroupMembershipAsync_WhenDesiredLoginIsDomainQualified_KeepsExistingClaimsMember()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"value\":[{\"LoginName\":\"i:0#.w|contoso\\\\asmith\"}]}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.SyncGroupMembershipAsync("Team Readers", ["CONTOSO\\asmith"]);

        handler.Requests.Should().ContainSingle("the member is already in the group, so nothing is added or removed");
    }

    [Fact]
    public async Task BindRoleToFolderAsync_WhenRoleIsNotBuiltIn_ResolvesRoleDefinitionByName()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":55}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":1073741930}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.BindRoleToFolderAsync("/sites/team/Attachments/2026", "Team Readers", "Approver's Read");

        handler.Requests[1].Uri.Should().Contain("roledefinitions/getbyname('Approver%27%27s%20Read')");
        handler.Requests[3].Uri.Should().Contain("addroleassignment(principalid=55,roledefid=1073741930)");
    }

    [Fact]
    public async Task BindRoleToFolderAsync_WhenRoleDefinitionMissing_ThrowsValidationException()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":55}"),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var act = () => sut.BindRoleToFolderAsync("/sites/team/Attachments/2026", "Team Readers", "No Such Role");

        await act.Should().ThrowAsync<SharePointValidationException>().WithMessage("*No Such Role*");
    }

    [Fact]
    public async Task BreakInheritanceAsync_UsesFolderListItemEndpointWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.BreakInheritanceAsync("/sites/team/Attachments/2026", copyRoleAssignments: true);

        handler.Requests[1].Uri.Should().Contain("breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task ResetInheritanceAsync_UsesResetEndpoint()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.ResetInheritanceAsync("/sites/team/Attachments/2026");

        handler.Requests[1].Uri.Should().Contain("resetroleinheritance");
    }

    [Fact]
    public async Task BindRoleToFolderAsync_EnsuresGroupFetchesRoleIdAndBindsAssignment()
    {
        // Each SharePointSecurityClient instance has its own role definition cache, so no cross-test cleanup needed.

        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":55}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":1073741826}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.BindRoleToFolderAsync("/sites/team/Attachments/2026", "Team Readers", "Read");

        handler.Requests.Should().HaveCount(6);
        handler.Requests[4].Uri.Should().Contain("roledefinitions/getbytype(2)");
        handler.Requests[5].Uri.Should().Contain("addroleassignment(principalid=55,roledefid=1073741826)");
    }

    [Fact]
    public async Task GetFolderRoleAssignmentsAsync_ReturnsPrincipalRoleMappings()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("""
                {
                  "value": [
                    {
                      "PrincipalId": 55,
                      "Member": {
                        "Title": "Team Readers",
                        "LoginName": "Team Readers",
                        "PrincipalType": 8
                      },
                      "RoleDefinitionBindings": [
                        {
                          "Id": 1073741826,
                          "Name": "Read"
                        }
                      ]
                    }
                  ]
                }
                """));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var result = await sut.GetFolderRoleAssignmentsAsync("/sites/team/Attachments/2026");

        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Uri.Should().Contain("/RoleAssignments?");
        handler.Requests[0].Uri.Should().Contain("$expand=Member,RoleDefinitionBindings");

        result.Should().HaveCount(1);
        result[0].PrincipalId.Should().Be(55);
        result[0].PrincipalTitle.Should().Be("Team Readers");
        result[0].Roles.Should().ContainSingle();
        result[0].Roles[0].Name.Should().Be("Read");
    }

    [Fact]
    public async Task BindRoleToFileAsync_ResolvesPrincipalRoleAndBindsAssignment()
    {
        // Each SharePointSecurityClient instance has its own role definition cache, so no cross-test cleanup needed.

        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":55}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":1073741826}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.BindRoleToFileAsync("/sites/team/Attachments/2026/file.txt", "Team Readers", "Read");

        handler.Requests.Should().HaveCount(4);
        handler.Requests[3].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2F2026%2Ffile.txt')");
        handler.Requests[3].Uri.Should().Contain("addroleassignment(principalid=55,roledefid=1073741826)");
    }

    [Fact]
    public async Task RemoveRoleFromFileAsync_WhenPrincipalExists_SendsDeleteRoleAssignmentRequest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":55}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.RemoveRoleFromFileAsync("/sites/team/Attachments/2026/file.txt", "Team Readers");

        handler.Requests[2].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2F2026%2Ffile.txt')");
        handler.Requests[2].Uri.Should().Contain("roleassignments/getbyprincipalid(55)");
        handler.Requests[2].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("DELETE");
        handler.Requests[2].Headers["IF-MATCH"].Should().ContainSingle().Which.Should().Be("*");
    }

    [Fact]
    public async Task GetFileRoleAssignmentsAsync_ReturnsPrincipalRoleMappings()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("""
                {
                  "value": [
                    {
                      "PrincipalId": 61,
                      "Member": {
                        "Title": "Team Owners",
                        "LoginName": "Team Owners",
                        "PrincipalType": 8
                      },
                      "RoleDefinitionBindings": [
                        {
                          "Id": 1073741827,
                          "Name": "Edit"
                        }
                      ]
                    }
                  ]
                }
                """));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        var result = await sut.GetFileRoleAssignmentsAsync("/sites/team/Attachments/2026/file.txt");

        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2F2026%2Ffile.txt')");
        handler.Requests[0].Uri.Should().Contain("/RoleAssignments?");
        handler.Requests[0].Uri.Should().Contain("$expand=Member,RoleDefinitionBindings");

        result.Should().HaveCount(1);
        result[0].PrincipalId.Should().Be(61);
        result[0].PrincipalTitle.Should().Be("Team Owners");
        result[0].Roles.Should().ContainSingle();
        result[0].Roles[0].Name.Should().Be("Edit");
    }

    [Fact]
    public async Task BreakFileInheritanceAsync_UsesFileListItemEndpointWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.BreakFileInheritanceAsync("/sites/team/Attachments/2026/report.xlsx", copyRoleAssignments: false);

        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2F2026%2Freport.xlsx')");
        handler.Requests[1].Uri.Should().Contain("breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task ResetFileInheritanceAsync_UsesFileResetEndpoint()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.ResetFileInheritanceAsync("/sites/team/Attachments/2026/report.xlsx");

        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2F2026%2Freport.xlsx')");
        handler.Requests[1].Uri.Should().Contain("resetroleinheritance");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task RemoveRoleFromFolderAsync_WhenGroupDoesNotExist_DoesNothing()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.RemoveRoleFromFolderAsync("/sites/team/Attachments/2026", "Missing Group");

        handler.Requests.Should().HaveCount(1);
    }

    [Fact]
    public async Task RemoveRoleFromFolderAsync_WhenGroupExists_SendsDeleteRoleAssignmentRequest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":55}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.RemoveRoleFromFolderAsync("/sites/team/Attachments/2026", "Team Readers");

        handler.Requests[2].Uri.Should().Contain("roleassignments/getbyprincipalid(55)");
        handler.Requests[2].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("DELETE");
        handler.Requests[2].Headers["IF-MATCH"].Should().ContainSingle().Which.Should().Be("*");
    }

    [Fact]
    public async Task DeleteGroupAsync_WhenGroupExists_UsesRemoveById()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Id\":88}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateSecurityClient(client);

        await sut.DeleteGroupAsync("Team Readers");

        handler.Requests[2].Uri.Should().Contain("sitegroups/removebyid(88)");
    }
}


