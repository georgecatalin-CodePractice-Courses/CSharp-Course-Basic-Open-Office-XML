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
            string fileName = @"\Create a table.docx";
            string completePathFile = pathToFile + fileName;

            /* *** Use Open Office XML to define the table *** */
            Table table = new Table();

            /* *** Use Open Office XML to create the first row of the table *** */
            TableRow tableRow = new TableRow();

            /* *** Use Open Office XML to create each cell(column) in the table heading *** */
            tableRow.Append(CreateTableCell("Column A"));
            tableRow.Append(CreateTableCell("Column B"));
            tableRow.Append(CreateTableCell("Column C"));
            tableRow.Append(CreateTableCell("Column D"));

            table.Append(tableRow);

            /* *** Use Open Office XML to create the other rows of the table *** */
            for (int i = 1; i <= 10; i++)
            {
               tableRow = new TableRow();

                tableRow.AppendChild(CreateTableCell("A" + i.ToString()));
                tableRow.AppendChild(CreateTableCell("B" + i.ToString()));
                tableRow.AppendChild(CreateTableCell("C" + i.ToString()));
                tableRow.AppendChild(CreateTableCell("D" + i.ToString()));

                table.AppendChild(tableRow);
            }

            Body body = new Body(table);
            Document document = new Document(body);

            /* *** Use Open Office XML to construct the file *** */
            using (WordprocessingDocument file=WordprocessingDocument.Create(completePathFile,WordprocessingDocumentType.Document))
            {
                file.AddMainDocumentPart();
                file.MainDocumentPart.Document = document;
                file.MainDocumentPart.Document.Save();
            }
        }


        /* *** This is a helper method that adds a text element in a run element in a paragraph element  *** */
        private static TableCell CreateTableCell(string text)
        {
            return new TableCell(new Paragraph(new Run(new Text(text))));
        }
    }
}

