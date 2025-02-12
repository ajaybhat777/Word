using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Identity.Client;

public class SharePointDocParser
{
    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly string _tenantId;
    private readonly string _sharePointSiteUrl;

    public SharePointDocParser(string clientId, string clientSecret, string tenantId, string sharePointSiteUrl)
    {
        _clientId = clientId;
        _clientSecret = clientSecret;
        _tenantId = tenantId;
        _sharePointSiteUrl = sharePointSiteUrl;
    }

    // Authenticate with Azure AD and get an access token
    private async Task<string> GetAccessTokenAsync()
    {
        var app = ConfidentialClientApplicationBuilder
            .Create(_clientId)
            .WithClientSecret(_clientSecret)
            .WithAuthority($"https://login.microsoftonline.com/{_tenantId}")
            .Build();

        string[] scopes = new[] { "https://graph.microsoft.com/.default" };
        var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
        return result.AccessToken;
    }

    // Download the Word document from SharePoint
    private async Task<Stream> DownloadDocumentAsync(string documentLibrary, string filePath)
    {
        var accessToken = await GetAccessTokenAsync();
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

        var encodedPath = Uri.EscapeDataString(filePath);
        var apiUrl = $"{_sharePointSiteUrl}/_api/web/GetFolderByServerRelativeUrl('{documentLibrary}')/Files('{encodedPath}')/$value";
        var response = await httpClient.GetAsync(apiUrl);
        return await response.Content.ReadAsStreamAsync();
    }

    // Extract content controls from the Word document
    public List<ContentControlInfo> ExtractContentControls(Stream docStream)
    {
        var contentControls = new List<ContentControlInfo>();
        using (var document = WordprocessingDocument.Open(docStream, false))
        {
            var body = document.MainDocumentPart.Document.Body;
            foreach (var sdt in body.Descendants<SdtElement>())
            {
                var properties = sdt.SdtProperties;
                var title = properties?.GetFirstChild<Alias>()?.Val?.Value ?? "";
                var tag = properties?.GetFirstChild<Tag>()?.Val?.Value ?? "";
                var text = sdt.InnerText;

                contentControls.Add(new ContentControlInfo
                {
                    Title = title,
                    Tag = tag,
                    Text = text
                });
            }
        }
        return contentControls;
    }

    public class ContentControlInfo
    {
        public string Title { get; set; }
        public string Tag { get; set; }
        public string Text { get; set; }
    }
}
