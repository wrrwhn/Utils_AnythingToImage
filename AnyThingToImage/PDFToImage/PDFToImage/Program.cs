using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFToImage
{
    class Program
    {
        static void Main(string[] args)
        {
            testStringFormat();

            Console.ReadKey();
        }

        private static void testSplit()
        {
            String[] sperators = { "," };
            String[] suppportExtensions = "doc,docx,ppt,pptx,xls,xlsx".Split(sperators, StringSplitOptions.RemoveEmptyEntries);
            Console.WriteLine(suppportExtensions.Contains("pp"));
            Console.WriteLine(suppportExtensions.Contains("ppt"));
            Console.WriteLine(suppportExtensions.Contains("pptx"));
        }

        private static void testStringFormat()
        {
            Console.WriteLine(String.Format("http://localhost:8080/resource/writeback?topicId={0}", 1, 2));
            Console.WriteLine(String.Format("http://localhost:8080/resource/writeback?topicId={0}&amp;exerciseId={1}", 1, 2).Replace("&amp;", "&"));
        }
    }
}
