using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordConverter.Utils;

namespace WordConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            String docFile = @"D:\资料备份\资料\工作\测试\anythingToImage\from\test.doc";
            String docxFile = @"D:\资料备份\资料\工作\测试\anythingToImage\from\test.docx";
            String outputFile = @"D:\资料备份\资料\工作\测试\anythingToImage\to\";

            WordUtils.ConvertToPDF(docFile, outputFile + docFile.Substring(docFile.LastIndexOf("\\")) + ".pdf");
            WordUtils.ConvertToPDF(docxFile, outputFile + docxFile.Substring(docxFile.LastIndexOf("\\")) + ".pdf");
        }
    }
}
