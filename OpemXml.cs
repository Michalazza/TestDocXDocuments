using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace TestDocXDocuments
{
    public class OpemXml
    {        
        public void TestOpenXml(string originalFilePath, string modifiedFilePath)
        {
            File.Copy(originalFilePath, modifiedFilePath);

            using (var wordprocessingDocument = WordprocessingDocument.Open(modifiedFilePath, isEditable: true))
            {
                // Do changes here...
                Table table = wordprocessingDocument.MainDocumentPart.Document.FirstChild.Elements<Table>().Where(p => p.InnerXml.Contains("«Criterion»")).FirstOrDefault();
                TableRow row = table.Elements<TableRow>().ElementAt(1);
                GridColumn col = new GridColumn();
                table.ChildElements.Where(p => p.LocalName.Equals("tblGrid")).FirstOrDefault().AppendChild<GridColumn>(col);
                Document doc = wordprocessingDocument.MainDocumentPart.Document;
                row.AppendChild(new TableCell(new Paragraph(new Run(new Text(" 3")))));
                //row.InsertAfter<TableCell>(new TableCell(), row.ElementAt(3));
                wordprocessingDocument.Close();
            }
        }

        public void OpenXmlTest1()
        {
            string txt = "Append text in body - OpenAndAddTextToWordDocument";
            var filepath = @"C:\\Meh\\Assesment_mapping_stremlined_template_v1.14.dotx";
            WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filepath, true);

            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));
            Table table = wordprocessingDocument.MainDocumentPart.Document.Elements<Table>().Where(p => p.InnerText.Contains("«Criterion»")).FirstOrDefault();

            //wordprocessingDocument.MainDocumentPart.Document.Save(fileStream);
            // Close the handle explicitly.
            wordprocessingDocument.Close();
        }
    }
}
