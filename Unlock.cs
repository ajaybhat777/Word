public byte[] RemoveContentControlByTagOrAlias(byte[] fileBytes, string identifier)
{
    using (var memStream = new MemoryStream())
    {
        memStream.Write(fileBytes, 0, fileBytes.Length);
        memStream.Position = 0;

        using (var wordDoc = WordprocessingDocument.Open(memStream, true))
        {
            var body = wordDoc.MainDocumentPart.Document.Body;

            // Find the specific SdtBlock by tag or alias
            var sdt = body.Elements<SdtBlock>()
                .FirstOrDefault(s =>
                {
                    var props = s.GetFirstChild<SdtProperties>();
                    var tag = props?.Elements<Tag>().FirstOrDefault();
                    var alias = props?.Elements<SdtAlias>().FirstOrDefault();

                    return (tag != null && tag.Val == identifier) ||
                           (alias != null && alias.Val == identifier);
                });

            if (sdt != null)
            {
                var contentBlock = sdt.GetFirstChild<SdtContentBlock>();
                if (contentBlock != null)
                {
                    var content = contentBlock.Elements().ToList();
                    foreach (var element in content)
                    {
                        body.AppendChild(element.CloneNode(true));
                    }
                }

                sdt.Remove(); // Remove the original SdtBlock
            }

            wordDoc.MainDocumentPart.Document.Save();
        }

        return memStream.ToArray();
    }
}
