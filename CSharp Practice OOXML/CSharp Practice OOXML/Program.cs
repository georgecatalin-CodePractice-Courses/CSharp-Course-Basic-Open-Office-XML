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
            string completePathToNewFile = "C:\\Test OpenXML\\myNewFile.docx";
            string completePathToExistingFile = "C:\\Test OpenXML\\myText.docx";

            /* *** Use Open Office XML to define content *** */


            /* *** Use System.IO to copy the existing file *** */
            if (File.Exists(completePathToNewFile))
            {
                File.Delete(completePathToNewFile);
            }

            File.Copy(completePathToExistingFile, completePathToNewFile);

            /* *** Use Open Office XML to open file and add its new content *** */
            Text text = new Text("This is the text to be added for testing the feature of setting the heading to a paragraph.");
            Run run = new Run(text);
            Paragraph paragraph = new Paragraph(run);
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = "Header1" });

            /* *** Use Open Office XML to construct the file *** */

            WordprocessingDocument file = WordprocessingDocument.Open(completePathToNewFile, true);

            file.MainDocumentPart.Document.AppendChild(paragraph);
            file.MainDocumentPart.Document.Save();

        }
    }
}
