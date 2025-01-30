using Aspose.Words;
using System.IO;

public void CheckImagesInDocx(byte[] docxBytes)
{
    using (MemoryStream stream = new MemoryStream(docxBytes))
    {
        Document doc = new Document(stream);
        
        // Extract all shapes (including images)
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Save the image to a file or check stream length
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    if (imageStream.Length == 0)
                    {
                        throw new InvalidDataException("Image stream is empty!");
                    }
                    File.WriteAllBytes($"image_{Guid.NewGuid()}.png", imageStream.ToArray());
                }
            }
        }
    }
}
using Aspose.Words;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using Newtonsoft.Json;

public string ConvertDocxToHtml(byte[] docxBytes)
{
    using (var docStream = new MemoryStream(docxBytes))
    {
        Document doc = new Document(docStream);
        var images = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().Where(s => s.HasImage);

        // Process each image
        foreach (var image in images)
        {
            using (var imageStream = new MemoryStream())
            {
                image.ImageData.Save(imageStream);
                var metadata = ExtractAlphaChannelMetadata(imageStream.ToArray());

                if (metadata != null)
                {
                    // Replace image source with SharePoint URL
                    string newSrc = $"https://your-sharepoint-site/_api/images/{metadata.SiteId}/{metadata.ListId}/{metadata.ItemId}";
                    ReplaceImageSourceInHtml(doc, image, newSrc);
                }
            }
        }

        // Save as HTML
        HtmlSaveOptions options = new HtmlSaveOptions { ExportImagesAsBase64 = true };
        using (var htmlStream = new MemoryStream())
        {
            doc.Save(htmlStream, options);
            return Encoding.UTF8.GetString(htmlStream.ToArray());
        }
    }
}

private SharePointImageData ExtractAlphaChannelMetadata(byte[] imageBytes)
{
    try
    {
        using (var image = Image.Load<Rgba32>(imageBytes))
        {
            var alphaBytes = new List<byte>();

            // Extract alpha channel data
            image.ProcessPixelRows(accessor =>
            {
                for (int y = 0; y < accessor.Height; y++)
                {
                    Span<Rgba32> row = accessor.GetRowSpan(y);
                    foreach (var pixel in row)
                    {
                        alphaBytes.Add(pixel.A); // Get alpha value
                    }
                }
            });

            // Trim null bytes and decode JSON
            string json = Encoding.UTF8.GetString(alphaBytes.ToArray()).TrimEnd('\0');
            return JsonConvert.DeserializeObject<SharePointImageData>(json);
        }
    }
    catch
    {
        return null;
    }
}

private void ReplaceImageSourceInHtml(Document doc, Shape image, string newSrc)
{
    // Find the image's HTML <img> tag and replace src
    // (Requires custom logic based on Aspose.Words HTML output)
    // Alternative: Use regex replacement on the final HTML string
}
