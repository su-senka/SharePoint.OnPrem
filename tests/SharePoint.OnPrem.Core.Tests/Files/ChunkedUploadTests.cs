using System.Net;
using AwesomeAssertions;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core.Tests.Contracts;

namespace SharePoint.OnPrem.Core.Tests.Files;

public class ChunkedUploadTests
{
    private const string Digest = "{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}";
    private const string StoredFile = "{\"ServerRelativeUrl\":\"/sites/team/Docs/big.bin\",\"Name\":\"big.bin\"}";

    [Fact]
    public async Task UploadChunkedAsync_WhenContentFitsInOneChunk_UsesPlainUpload()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => SharePointContractTestHelpers.JsonResponse(StoredFile));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));
        var progress = new RecordingProgress();

        await using var content = new MemoryStream(new byte[8]);
        var result = await sut.UploadChunkedAsync(new UploadFileRequest("/sites/team/Docs", "big.bin", content), chunkSizeBytes: 10, progress);

        result.ServerRelativeUrl.Should().Be("/sites/team/Docs/big.bin");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("/Files/add(url='big.bin',overwrite=true)");
        progress.Values.Should().Equal(8);
    }

    [Fact]
    public async Task UploadChunkedAsync_SendsStartContinueFinishWithOffsets()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => SharePointContractTestHelpers.JsonResponse(StoredFile),
            _ => SharePointContractTestHelpers.JsonResponse("{\"value\":\"10\"}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"value\":\"20\"}"),
            _ => SharePointContractTestHelpers.JsonResponse(StoredFile));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));
        var progress = new RecordingProgress();

        await using var content = new NonSeekableStream(Enumerable.Range(0, 25).Select(i => (byte)'a').ToArray());
        var result = await sut.UploadChunkedAsync(new UploadFileRequest("/sites/team/Docs", "big.bin", content), chunkSizeBytes: 10, progress);

        result.Should().Be(new SharePointStoredFile("/sites/team/Docs/big.bin", "big.bin"));
        handler.Requests.Should().HaveCount(5);
        handler.Requests[1].Body.Should().BeEmpty("the session needs an existing, empty file");
        handler.Requests[2].Uri.Should().MatchRegex(@"GetFileByServerRelativeUrl\('%2Fsites%2Fteam%2FDocs%2Fbig\.bin'\)/StartUpload\(uploadId=guid'[0-9a-f-]{36}'\)$");
        handler.Requests[2].Body.Should().HaveLength(10);
        handler.Requests[3].Uri.Should().Contain(",fileOffset=10)");
        handler.Requests[3].Uri.Should().Contain("/ContinueUpload(");
        handler.Requests[3].Body.Should().HaveLength(10);
        handler.Requests[4].Uri.Should().Contain("/FinishUpload(").And.EndWith(",fileOffset=20)");
        handler.Requests[4].Body.Should().HaveLength(5);
        progress.Values.Should().Equal(10, 20, 25);
    }

    [Fact]
    public async Task UploadChunkedAsync_WhenLengthIsExactMultipleOfChunk_FinishesWithLastFullChunk()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => SharePointContractTestHelpers.JsonResponse(StoredFile),
            _ => SharePointContractTestHelpers.JsonResponse("{}"),
            _ => SharePointContractTestHelpers.JsonResponse(StoredFile));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await using var content = new MemoryStream(new byte[20]);
        await sut.UploadChunkedAsync(new UploadFileRequest("/sites/team/Docs", "big.bin", content), chunkSizeBytes: 10);

        handler.Requests.Should().HaveCount(4);
        handler.Requests[3].Uri.Should().Contain("/FinishUpload(").And.EndWith(",fileOffset=10)");
        handler.Requests[3].Body.Should().HaveLength(10);
    }

    [Fact]
    public async Task UploadChunkedAsync_WhenChunkFails_CancelsSessionAndRethrows()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest),
            _ => SharePointContractTestHelpers.JsonResponse(StoredFile),
            _ => SharePointContractTestHelpers.JsonResponse("{}"),
            _ => new HttpResponseMessage(HttpStatusCode.ServiceUnavailable),
            _ => new HttpResponseMessage(HttpStatusCode.OK));
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await using var content = new MemoryStream(new byte[25]);
        var act = () => sut.UploadChunkedAsync(new UploadFileRequest("/sites/team/Docs", "big.bin", content), chunkSizeBytes: 10);

        await act.Should().ThrowAsync<SharePointTransientException>();
        handler.Requests.Should().HaveCount(5);
        handler.Requests[4].Uri.Should().Contain("/CancelUpload(uploadId=guid'");
        handler.Requests[4].Headers["X-RequestDigest"].Should().ContainSingle();
    }

    [Fact]
    public async Task UploadChunkedAsync_WhenChunkSizeInvalid_Throws()
    {
        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(new QueueMessageHandler()));

        var act = () => sut.UploadChunkedAsync(new UploadFileRequest("/sites/team/Docs", "a.bin", Stream.Null), chunkSizeBytes: 0);

        await act.Should().ThrowAsync<ArgumentOutOfRangeException>();
    }

    private sealed class RecordingProgress : IProgress<long>
    {
        public List<long> Values { get; } = [];
        public void Report(long value) => Values.Add(value);
    }

    private sealed class NonSeekableStream(byte[] data) : MemoryStream(data)
    {
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    }
}
