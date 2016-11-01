using PPTConverter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTConverter
{
    public class Program
    {
        static void Main(string[] args)
        {
            String pptFile = @"D:\资料备份\资料\工作\测试\anythingToImage\from\test.ppt";
            String pptxFile = @"D:\资料备份\资料\工作\测试\anythingToImage\from\test.pptx";
            String outputFile = @"D:\资料备份\资料\工作\测试\anythingToImage\to\";

            //PPTUtils.ConvertToPDF(xlsFile, outputFile + xlsFile.Substring(xlsFile.LastIndexOf("\\")) + ".pdf");
            //PPTUtils.ConvertToPDF(xlsxFile, outputFile + xlsxFile.Substring(xlsxFile.LastIndexOf("\\")) + ".pdf");            
            PPTUtils.ConvertToIMAGE(pptFile, outputFile + "ppt");
            PPTUtils.ConvertToIMAGE(pptFile, outputFile+ "pptx");

            Console.ReadKey();
        }
    }
}
