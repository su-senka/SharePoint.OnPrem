using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();
services.AddSharePointOnPrem(
    configureCore: options =>
    {
        options.SiteBaseUrl = "https://sharepoint.local";
        options.HttpTimeoutMinutes = 5;
    },
    configureIdentity: options =>
    {
        options.Domain = "ACR";
        options.ClaimsPrefix = "i:0#.w|";
    },
    configureStorageScope: options =>
    {
        options.BaseFolderServerRelativeUrl = "/sites/pp/Attachments";
    });

using var provider = services.BuildServiceProvider();

var normalizer = provider.GetRequiredService<ILoginNameNormalizer>();
var scope = provider.GetRequiredService<ISharePointPathScope>();

Console.WriteLine(normalizer.Normalize("jnovak"));
Console.WriteLine(scope.ToServerRelativePath("3255/26/001/003"));

