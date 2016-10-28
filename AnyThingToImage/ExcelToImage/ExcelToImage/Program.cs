using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelConverter.Utils;

namespace ExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            String xlsFile = @"D:\资料备份\资料\工作\测试\anythingToImage\from\test.xls";
            String xlsxFile = @"D:\资料备份\资料\工作\测试\anythingToImage\from\test.xlsx";
            String outputFile = @"D:\资料备份\资料\工作\测试\anythingToImage\to\";

            ExcelUtils.ConvertToPDF(xlsFile, outputFile + xlsFile.Substring(xlsFile.LastIndexOf("\\")) + ".pdf");
            ExcelUtils.ConvertToPDF(xlsxFile, outputFile + xlsxFile.Substring(xlsxFile.LastIndexOf("\\")) + ".pdf");
        }
    }
}
