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
            string fileName = @"\Set the text color.docx";
            string completePathFile = pathToFile + fileName;

            /* *** Use Open Office XML to define the styles *** */
            RunProperties runProperties = new RunProperties();

            Color color = new Color();
            color.Val = "00FF00";

            runProperties.Append(color);

            /* *** Use Open Office XML to create content *** */
            Run run = new Run();
            run.Append(runProperties);
            run.Append(new Text("This is text to be added with a new color."));

            Paragraph paragraph = new Paragraph();
            paragraph.Append(run);

            Body body = new Body();
            body.Append(paragraph);

            Document document = new Document(body);

            /* *** Use Open Office XML to construct the file *** */
            using (WordprocessingDocument file=WordprocessingDocument.Create(completePathFile,WordprocessingDocumentType.Document))
            {
                file.AddMainDocumentPart();
                file.MainDocumentPart.Document = document;
                file.MainDocumentPart.Document.Save();
            }

            }

        }
    }

