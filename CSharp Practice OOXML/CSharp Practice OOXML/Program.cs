using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace CSharp_Practice_OOXML
{
    class Program
    {
        static void Main(string[] args)
        {
            /* *** Define names and complete path of the file *** */
            string filename = "\\New Word Document With Styles.docx";
            string path = "C:\\Test OpenXML";
            string completePathOfNewFile = path + filename;

            /* *** Create document *** */
            using (WordprocessingDocument file = WordprocessingDocument.Create(completePathOfNewFile, WordprocessingDocumentType.Document)) 
            {
                file.AddMainDocumentPart();

                //Create styles definitions part
                StyleDefinitionsPart styleDefinitionsPart = file.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

                //Create styles
                Styles styles = new Styles();
                styles.Save(styleDefinitionsPart);
                styles = styleDefinitionsPart.Styles;

                //Create style
                Style style = new Style() 
                { 
                Type=StyleValues.Paragraph,
                StyleId="header1",
                CustomStyle=true,
                Default=false
                };

                style.Append(new StyleName() { Val = "Header 1" });

                StyleRunProperties styleRunProperties = new StyleRunProperties();
                styleRunProperties.Append(new Bold());
                styleRunProperties.Append(new RunFonts() { Ascii = "Arial" });
                styleRunProperties.Append(new FontSize() { Val = "24" });

                style.Append(styleRunProperties);
                styles.Append(style);

                /*** Use Open Office XML to contain content in Word File *** */
                Text text = new Text("This is text created with the purpose of setting a new style");
                Run run = new Run(text);
                Paragraph paragraph = new Paragraph(run);
                Body body = new Body(paragraph);
                Document document = new Document(body);

                paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = "header1" });

                file.MainDocumentPart.Document = document;
                file.MainDocumentPart.Document.Save();



            }

        }
    }
}
