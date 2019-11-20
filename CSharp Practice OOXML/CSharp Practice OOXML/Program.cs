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
            string pathToFolder = "C:\\Test OpenXML";
            string filename = "\\Set Background of an Element.docx";
            string completePathToFile = pathToFolder + filename;

            /* *** Use Open Office XML to set RunProperties *** */
            RunProperties runProperties = new RunProperties();

            Shading shading = new Shading();
            shading.Color = "auto";
            shading.Fill = "00FF00";
            shading.Val = ShadingPatternValues.Clear;

            runProperties.Append(shading);

            /* *** Use Open Office XML to add content *** */
            Run run = new Run();
            run.Append(runProperties);
            run.Append(new Text("This is new text with the purpose of changing background color."));

            Paragraph paragraph = new Paragraph(run);
            Body body = new Body(paragraph);
            Document document = new Document(body);

            /* *** Use Open Office XML to construct the document *** */
            using (WordprocessingDocument file = WordprocessingDocument.Create(completePathToFile, WordprocessingDocumentType.Document))
            {
                file.AddMainDocumentPart();
                file.MainDocumentPart.Document = document;
                file.MainDocumentPart.Document.Save();
            }

            }

        }
    }

