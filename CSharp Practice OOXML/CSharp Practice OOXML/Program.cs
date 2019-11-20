using System;
using System.Collections.Generic;
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
            /* *** Define the complete path to the file which is about to be constructed *** */
            string pathToFolder = "C:\\Test OpenXML";
            string completePathToFile =pathToFolder+ "\\myText.docx";

            /* *** Use Open Office XML to define content *** */
            Text text = new Text("This is my first Open Office Application.");
            Run run = new Run(text);
            Paragraph paragraph = new Paragraph(run);
            Body body = new Body(paragraph);
            Document document = new Document(body);

            /* *** Use Open Office XML to build the file *** */
            using (var file=WordprocessingDocument.Create(completePathToFile,WordprocessingDocumentType.Document))
            {
                file.AddMainDocumentPart();
                file.MainDocumentPart.Document = document;
                file.MainDocumentPart.Document.Save();
            }
        }
    }
}
