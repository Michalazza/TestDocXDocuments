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
            //var originalFilePath = @"C:\Meh\Assesment_mapping_stremlined_template_v1.14.dotx";
            //var modifiedFilePath = @"C:\Meh\Meh.dotx";

            var originalFilePath = @"Z:\Meh\Assesment_mapping_stremlined_template_v1.14.dotx";
            var modifiedFilePath = @"Z:\Meh\Meh.docx";

            Console.WriteLine("1. Use Docx to make Document");
            Console.WriteLine("2. Use Openxml to make Document");
            var docx = new Docx();
            var openXml = new OpemXml();
            var x = Console.ReadLine();
            if(x.Equals("1"))
            {
                docx.TestDocx(originalFilePath, modifiedFilePath);
            }
            else if(x.Equals("2"))
            {
                openXml.TestOpenXml(originalFilePath, modifiedFilePath);
            }
            Console.ReadKey();
        }        
    }
}
