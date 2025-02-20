using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Text;

public string ConvertContentControlToHtml(string filePath, string contentControlAlias)
{
    StringBuilder html = new StringBuilder();

    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
    {
        MainDocumentPart mainPart = doc.MainDocumentPart;
        var contentControl = FindContentControl(mainPart.Document.Body, contentControlAlias);

        if (contentControl != null)
        {
            html.Append(ConvertContentControlContentToHtml(contentControl));
        }
    }

    return html.ToString();
}

private SdtElement FindContentControl(Body body, string contentControlAlias)
{
    // Traverse all content controls in the document
    foreach (var sdt in body.Descendants<SdtElement>())
    {
        // Check by Alias
        var alias = sdt.SdtProperties?.Elements<SdtAlias>().FirstOrDefault();
        if (alias != null && alias.Val.Value == contentControlAlias)
        {
            return sdt;
        }

        // Check by Tag (if Alias is not used)
        var tag = sdt.SdtProperties?.Elements<Tag>().FirstOrDefault();
        if (tag != null && tag.Val.Value == contentControlAlias)
        {
            return sdt;
        }
    }
    return null;
}

private string ConvertContentControlContentToHtml(SdtElement contentControl)
{
    StringBuilder html = new StringBuilder();

    // Traverse the content inside the content control
    foreach (var element in contentControl.Descendants())
    {
        if (element is Paragraph paragraph)
        {
            if (IsListItem(paragraph))
            {
                html.Append(ConvertListItemToHtml(paragraph));
            }
            else
            {
                html.Append(ConvertParagraphToHtml(paragraph));
            }
        }
        else if (element is Run run)
        {
            html.Append(ConvertRunToHtml(run));
        }
    }

    return html.ToString();
}

private bool IsListItem(Paragraph paragraph)
{
    // Check if the paragraph has numbering properties (indicating it's part of a list)
    return paragraph.ParagraphProperties?.NumberingProperties != null;
}

private string ConvertListItemToHtml(Paragraph paragraph)
{
    StringBuilder listItemHtml = new StringBuilder();

    // Start a new unordered list if this is the first list item
    if (!IsPreviousElementListItem(paragraph))
    {
        listItemHtml.Append("<ul>");
    }

    // Add the list item
    listItemHtml.Append("<li>");

    // Traverse the runs and text within the paragraph
    foreach (var run in paragraph.Elements<Run>())
    {
        listItemHtml.Append(ConvertRunToHtml(run));
    }

    listItemHtml.Append("</li>");

    // Close the unordered list if this is the last list item
    if (!IsNextElementListItem(paragraph))
    {
        listItemHtml.Append("</ul>");
    }

    return listItemHtml.ToString();
}

private bool IsPreviousElementListItem(Paragraph paragraph)
{
    // Check if the previous element is also a list item
    var previousElement = paragraph.PreviousSibling<Paragraph>();
    return previousElement != null && IsListItem(previousElement);
}

private bool IsNextElementListItem(Paragraph paragraph)
{
    // Check if the next element is also a list item
    var nextElement = paragraph.NextSibling<Paragraph>();
    return nextElement != null && IsListItem(nextElement);
}

private string ConvertParagraphToHtml(Paragraph paragraph)
{
    StringBuilder paragraphHtml = new StringBuilder();

    // Check if the paragraph is a heading
    var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
    if (style != null && style.Value.StartsWith("Heading"))
    {
        int headingLevel = int.Parse(style.Value.Replace("Heading", ""));
        paragraphHtml.Append($"<h{headingLevel}>");
    }
    else
    {
        paragraphHtml.Append("<p>");
    }

    // Traverse the runs and text within the paragraph
    foreach (var run in paragraph.Elements<Run>())
    {
        paragraphHtml.Append(ConvertRunToHtml(run));
    }

    if (style != null && style.Value.StartsWith("Heading"))
    {
        int headingLevel = int.Parse(style.Value.Replace("Heading", ""));
        paragraphHtml.Append($"</h{headingLevel}>");
    }
    else
    {
        paragraphHtml.Append("</p>");
    }

    return paragraphHtml.ToString();
}

private string ConvertRunToHtml(Run run)
{
    StringBuilder runHtml = new StringBuilder();

    // Check for formatting (bold, italic, underline, etc.)
    bool isBold = run.RunProperties?.Elements<Bold>().Any() ?? false;
    bool isItalic = run.RunProperties?.Elements<Italic>().Any() ?? false;
    bool isUnderline = run.RunProperties?.Elements<Underline>().Any() ?? false;

    if (isBold) runHtml.Append("<strong>");
    if (isItalic) runHtml.Append("<em>");
    if (isUnderline) runHtml.Append("<u>");

    // Append the text content
    foreach (var text in run.Elements<Text>())
    {
        runHtml.Append(text.Text);
    }

    if (isUnderline) runHtml.Append("</u>");
    if (isItalic) runHtml.Append("</em>");
    if (isBold) runHtml.Append("</strong>");

    return runHtml.ToString();
}
