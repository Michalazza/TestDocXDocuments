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
    class Program
    {
        static void Main(string[] args)
        {
            string originalFilePath = @"C:\Meh\Assesment_mapping_stremlined_template_v1.14.dotx";
            string modifiedFilePath = @"C:\Meh\Meh.dotx";


            
            

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

        public void Docx()
        {
            ////Need to load a document
            //DocX document = DocX.Load(@"C:\\Meh\\Assesment_mapping_stremlined_template_v1.14.dotx");
            ////Need to create a document
            //DocX documentsomething = DocX.Create(@"C:\\Meh\\Meh.docx");
            ////Need to insert this document into new one
            //documentsomething.InsertDocument(document, true);
            ////Replace Text
            //documentsomething.ReplaceText("«TeachingSection»", "Hellooo", false);

            ////Get table where it contains Criterion Element
            //Table tableCriteria = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Criterion»")).FirstOrDefault();
            //Table tableEvidence = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Performance»")).FirstOrDefault();
            //Table tableKnowledge = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Knowledge»")).FirstOrDefault();
            //Table tableAssessment = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Assessment»")).FirstOrDefault();

            //var row1 = tableCriteria.Rows[1].Xml;
            //Row row2 = tableCriteria.Rows[2];
            //for (var count = 0; count < 2; count++)
            //{
            //    tableCriteria.InsertColumn(count + 2);
            //}
            //for (var count = 0; count < 2; count++)
            //{
            //    tableCriteria.InsertColumn(count + 4);
            //}
            //tableCriteria.Rows[0].MergeCells(1, 3);
            //tableCriteria.Rows[1].Xml = row1;

            //tableAssessment.InsertColumn(2);
            //tableAssessment.InsertColumn(2);
            //Row row = tableAssessment.InsertRow();
            //row.MergeCells(1, 3);
            //tableAssessment.Rows.Add(row);
            //documentsomething.Save();
        }
    }
}
