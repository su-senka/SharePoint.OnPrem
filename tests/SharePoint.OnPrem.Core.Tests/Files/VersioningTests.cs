using AwesomeAssertions;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core.Tests.Contracts;

namespace SharePoint.OnPrem.Core.Tests.Files;

public class VersioningTests
{
    private const string Digest = "{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}";

    [Fact]
    public async Task CheckOutAsync_PostsCheckOutWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => new HttpResponseMessage(System.Net.HttpStatusCode.OK));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await sut.CheckOutAsync("/sites/team/Docs/contract.docx");

        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().EndWith("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FDocs%2Fcontract.docx')/CheckOut()");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task CheckInAsync_EncodesCommentAndCheckInType()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => new HttpResponseMessage(System.Net.HttpStatusCode.OK));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await sut.CheckInAsync("/sites/team/Docs/contract.docx", "Legal's review", SharePointCheckInType.MajorCheckIn);

        handler.Requests[1].Uri.Should().EndWith("/CheckIn(comment='Legal%27%27s%20review',checkintype=1)");
    }

    [Fact]
    public async Task UndoCheckOutAsync_PostsUndoCheckOut()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => new HttpResponseMessage(System.Net.HttpStatusCode.OK));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await sut.UndoCheckOutAsync("/sites/team/Docs/contract.docx");

        handler.Requests[1].Uri.Should().EndWith("/UndoCheckOut()");
    }

    [Fact]
    public async Task GetVersionsAsync_ReturnsVersions()
    {
        var handler = new QueueMessageHandler(_ => SharePointContractTestHelpers.JsonResponse(
            "{\"value\":[{\"ID\":512,\"VersionLabel\":\"1.0\",\"Created\":\"2026-03-01T10:00:00Z\",\"Size\":2048,\"CheckInComment\":\"Initial\",\"IsCurrentVersion\":false}," +
            "{\"ID\":513,\"VersionLabel\":\"1.1\",\"Created\":\"2026-03-02T10:00:00Z\",\"Size\":\"4096\",\"CheckInComment\":\"\",\"IsCurrentVersion\":false}]}"));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        var versions = await sut.GetVersionsAsync("/sites/team/Docs/contract.docx");

        handler.Requests[0].Uri.Should().Contain("/Versions?$select=");
        versions.Should().Equal(
            new SharePointFileVersion(512, "1.0", new DateTimeOffset(2026, 3, 1, 10, 0, 0, TimeSpan.Zero), 2048, "Initial", false),
            new SharePointFileVersion(513, "1.1", new DateTimeOffset(2026, 3, 2, 10, 0, 0, TimeSpan.Zero), 4096, null, false));
    }
}
