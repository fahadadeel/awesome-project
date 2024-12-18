using System.Linq;
/// <summary>
/// Creates a new Word Document with structured content using 
/// <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a>.
/// Generates paragraphs with heading styles defined by the Word document template.
/// Adds normal paragraphs under each heading paragraph, including text runs with various fonts as per the template.
/// Saves the newly created Word Document.
/// </summary>
/// <param name="documentDirectory">
/// The directory where the Word Document will be saved (default is the root of your project).
/// </param>
/// <param name="filename">
/// The name of the Word Document file (default is "WordParagraphs.docx") to be saved to the directlry 
/// defined by documentDirectory..
/// </param>
static void CreateWordParagraphs(string documentDirectory = "../../../", string filename = "WordParagraphs.docx")
{
    try
    {
        // Initialize a new word document with the default template         
        var doc = new FileFormat.Words.Document();
        System.Console.WriteLine("Word Document with default template initialized");

        // Initialize the body with the new document
        var body = new FileFormat.Words.Body(doc);
        System.Console.WriteLine("Body of the Word Document initialized");

        // Get all paragraph styles
        var paragraphStyles = doc.GetElementStyles().ParagraphStyles;
        System.Console.WriteLine("Paragraph styles loaded");

        // Get all fonts defined by FontTable and Theme
        var fonts = doc.GetElementStyles().TableFonts;
        var fontsTheme = doc.GetElementStyles().ThemeFonts;
        System.Console.WriteLine("Fonts defined by FontsTable and Theme loaded");

        // Merge all fonts
        fonts.AddRange(fontsTheme);
        System.Console.WriteLine("All Fonts merged");

        // Create Headings Paragraph and append to the body.
        foreach (var paragraphStyle in paragraphStyles.Where(style => !style.Contains("Normal")))
        {
            var paragraphWithStyle = new FileFormat.Words.IElements.Paragraph { Style = paragraphStyle };
            paragraphWithStyle.AddRun(new FileFormat.Words.IElements.Run
            {
                Text = $"Paragraph with {paragraphStyle} Style"
            });
            System.Console.WriteLine($"Styled Paragraph with {paragraphStyle} Created");
            body.AppendChild(paragraphWithStyle);
            System.Console.WriteLine($"Styled Paragraph with {paragraphStyle} Appended to Word Document Body");

            // Create Normal Paragraph and include text runs with various fonts as per the template.
            var paragraphNormal = new FileFormat.Words.IElements.Paragraph();
            System.Console.WriteLine("Normal Paragraph Created");
            paragraphNormal.AddRun(new FileFormat.Words.IElements.Run
            {
                Text = $"Text in normal paragraph with default font and size but with bold " +
                       $"and underlined Gray Color ",
                Color = FileFormat.Words.IElements.Colors.Gray,
                Bold = true,
                Underline = true
            });
            foreach (var font in fonts)
            {
                paragraphNormal.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = $"Text in normal paragraph with font {font} and size 10 but with default " +
                           $"color, bold, and underlines. ",
                    FontFamily = font,
                    FontSize = 10
                });
            }
            System.Console.WriteLine("All Runs with all fonts Created for Normal Paragraph");
            body.AppendChild(paragraphNormal);
            System.Console.WriteLine($"Normal Paragraph Appended to Word Document Body");
        }

        // Save the newly created Word Document.
        doc.Save($"{documentDirectory}/{filename}");
        System.Console.WriteLine($"Word Document {filename} Created. Please check directory: " +
                                 $"{System.IO.Path.GetFullPath(documentDirectory)}");
    }
    catch (System.Exception ex)
    {
        throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
    }
}
