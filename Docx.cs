using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestDocXDocuments
{
    public class Docx
    {
        public void TestDocx(string originalFilePath, string modifiedFilePath)
        {
            //Need to load a document
            DocX document = DocX.Load(originalFilePath);
            //Need to create a document
            DocX documentsomething = DocX.Create(modifiedFilePath);
            //Need to insert this document into new one
            documentsomething.InsertDocument(document, true);
            //Replace Text
            documentsomething.ReplaceText("«TeachingSection»", "Hellooo", false);

            //Get table where it contains Criterion Element
            Table tableCriteria = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Criterion»")).FirstOrDefault();
            Table tableEvidence = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Performance»")).FirstOrDefault();
            Table tableKnowledge = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Knowledge»")).FirstOrDefault();
            Table tableAssessment = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Assessment»")).FirstOrDefault();

            var row3 = tableCriteria.Rows[0].Xml;
            var row1 = tableCriteria.Rows[1].Xml;
            Row row2 = tableCriteria.Rows[2];
            for (var count = 0; count < 2; count++)
            {
                tableCriteria.InsertColumn(count + 2);
            }
            for (var count = 0; count < 2; count++)
            {
                tableCriteria.InsertColumn(count + 4);
            }
            tableCriteria.Rows[0].MergeCells(1, 3);
            tableCriteria.Rows[1].Xml = row1;

            tableAssessment.InsertColumn(2);
            tableAssessment.InsertColumn(2);
            Row row = tableAssessment.InsertRow();
            row.MergeCells(1, 3);
            tableAssessment.Rows.Add(row);
            documentsomething.Save();
        }
    }
}
