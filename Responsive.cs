HtmlSaveOptions saveOptions = new HtmlSaveOptions
{
    // Use inline CSS (critical for email clients)
    CssStyleSheetType = CssStyleSheetType.Inline,
    
    // Use relative font sizes (em/% instead of fixed pt/px)
    ExportRelativeFontSize = true,
    
    // Embed images as base64 to avoid broken links
    ExportImagesAsBase64 = true,
    
    // Avoid external font dependencies
    ExportFontResources = false,
    
    // Add viewport meta tag (optional but useful for mobile)
    HtmlHead = "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">",
    
    // Force tables to use percentage-based widths
    TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly
};

saveOptions.HtmlHead = @"
    <meta name='viewport' content='width=device-width, initial-scale=1'>
    <style type='text/css'>
      /* Basic email reset */
      body { margin: 0; padding: 0; }
      img { max-width: 100%; height: auto; }
      table { border-collapse: collapse; width: 100% !important; }
      .wrapper { width: 100%; max-width: 600px; margin: 0 auto; }
      /* Hybrid coding for Outlook */
      .outer-table { width: 100%; max-width: 600px; margin: 0 auto; }
    </style>
";
