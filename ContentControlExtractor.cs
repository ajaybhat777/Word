using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class HyperlinkContentControlProcessor
{
    public Dictionary<string, string> GetHyperlinks(byte[] wordBytes)
    {
        var hyperlinks = new Dictionary<string, string>();

        using (var stream = new MemoryStream(wordBytes))
        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = doc.MainDocumentPart;
            var document = mainPart.Document;

            foreach (var sdt in document.Descendants<SdtElement>())
            {
                // Get content control name
                var props = sdt.SdtProperties;
                var name = props?.GetFirstChild<Tag>()?.Val?.Value 
                         ?? props?.GetFirstChild<SdtAlias>()?.Val?.Value;

                if (string.IsNullOrEmpty(name)) continue;

                // Find the first hyperlink in the content control
                var hyperlink = sdt.Descendants<Hyperlink>().FirstOrDefault();
                if (hyperlink?.Id?.Value == null) continue;

                // Resolve the hyperlink URL from relationships
                var rel = mainPart.HyperlinkRelationships
                    .FirstOrDefault(r => r.Id == hyperlink.Id.Value);

                if (rel != null)
                    hyperlinks[name] = rel.Uri.ToString();
            }
        }

        return hyperlinks;
    }
}
----------

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

public class WordContentControlProcessor
{
    public string ProcessDocument(byte[] wordBytes, string htmlTemplate)
    {
        var placeholderValues = new Dictionary<string, string>();

        using (var stream = new MemoryStream(wordBytes))
        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = doc.MainDocumentPart;
            var document = mainPart.Document;

            foreach (var sdt in document.Descendants<SdtElement>())
            {
                var props = sdt.SdtProperties;
                var name = props?.GetFirstChild<Tag>()?.Val?.Value 
                         ?? props?.GetFirstChild<SdtAlias>()?.Val?.Value;

                if (string.IsNullOrEmpty(name)) continue;

                var content = sdt.SdtContentBlock;
                if (content == null) continue;

                var tempBody = new Body(content.Elements().Select(e => e.CloneNode(true)));
                var settings = new HtmlConverterSettings
                {
                    ImageHandler = ConvertImageToUrl
                };

                string html = HtmlConverter.ConvertToHtml(tempBody, mainPart, settings);
                placeholderValues[name] = html;
            }
        }

        foreach (var entry in placeholderValues)
            htmlTemplate = htmlTemplate.Replace("{{" + entry.Key + "}}", entry.Value);

        return htmlTemplate;
    }

    private string ConvertImageToUrl(ImageInfo imageInfo)
    {
        string extension = imageInfo.ContentType.Split('/')[1];
        string filename = $"{Guid.NewGuid()}.{extension}";
        string path = Path.Combine("wwwroot/images", filename);
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        File.WriteAllBytes(path, imageInfo.ImageByteArray);
        return $"/images/{filename}";
    }
}
