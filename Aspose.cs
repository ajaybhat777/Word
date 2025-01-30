Document doc = new Document("input.docx");
NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
foreach (Shape shape in shapes)
{
    if (shape.HasImage)
    {
        using (MemoryStream stream = new MemoryStream())
        {
            shape.ImageData.Save(stream);
            // Check if stream has data
        }
    }
}
