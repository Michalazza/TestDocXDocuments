using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestDocXDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            //Need to load a document
            DocX document = DocX.Load(@"C:\\Meh\\Assesment_mapping_stremlined_template_v1.14.dotx");
            //Need to create a document
            DocX documentsomething = DocX.Create(@"C:\\Meh\\Meh.docx");
            //Need to insert this document into new one
            documentsomething.InsertDocument(document, true);
            //Replace Text
            documentsomething.ReplaceText("«TeachingSection»", "Hellooo", false);
            
            //Get table where it contains Criterion Element
            Table tableCriteria = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Criterion»")).FirstOrDefault();
            Table tableEvidence = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Performance»")).FirstOrDefault();
            Table tableKnowledge = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Knowledge»")).FirstOrDefault();
            Table tableAssessment = documentsomething.Tables.Where(p => p.Xml.Value.Contains("«Assessment»")).FirstOrDefault();

            tableAssessment.InsertColumn(2);
            tableAssessment.InsertColumn(2);
            Row row = tableAssessment.InsertRow();
            row.MergeCells(1, 3);
            tableAssessment.Rows.Add(row);            
            documentsomething.Save();
        }
    }
}
