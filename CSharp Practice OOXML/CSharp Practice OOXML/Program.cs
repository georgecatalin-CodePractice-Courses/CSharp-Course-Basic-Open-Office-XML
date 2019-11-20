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
            string pathToFile = @"C:\Test OpenXML";
            string fileName = @"\Set the text font.docx";
            string completePathFile = pathToFile + fileName;

            /* *** Use Open Office XML to define the styles *** */
            RunFonts runFonts = new RunFonts();
            runFonts.Ascii = "Tahoma";

            Run run = new Run();
            run.AppendChild(runFonts);
            run.AppendChild(new Text("This is new text created with the purpose to set the Font in OOXML."));

            Paragraph paragraph = new Paragraph();
            paragraph.AppendChild(run);

            Body body = new Body();
            body.AppendChild(paragraph);

            Document document = new Document();
            document.Append(body);

            /* *** Use Open Office XML to construct the file *** */
            using (WordprocessingDocument file = WordprocessingDocument.Create(completePathFile, WordprocessingDocumentType.Document))
            {
                file.AddMainDocumentPart();
                file.MainDocumentPart.Document = document;
                file.MainDocumentPart.Document.Save();
            }
        }
    }
}

