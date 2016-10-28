using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTConverter.Utils
{
    public class PPTUtils
    {
        public static String ConvertToPDF(String filePath, String destPath)
        {
            // file exist & check type
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.EndsWith("ppt") && !filePath.EndsWith("pptx"))
                throw new Exception("文件非 Powerpoint 类型");
            
            // TODO 特殊格式支持

            Application app = new Application();
            Presentation presentation = app.Presentations.Open2007(filePath, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoFalse
                , Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            presentation.SaveAs(destPath, PpSaveAsFileType.ppSaveAsPDF);
            presentation.Close();
            app.Quit();

            return destPath;
        }

        public static String ConvertToIMAGE(String filePath, String destPath)
        {
            Application app = new Application();
            Presentation presentation = app.Presentations.Open2007(filePath, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoFalse
                , Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            presentation.SaveAs(destPath, PpSaveAsFileType.ppSaveAsJPG);
            presentation.Close();
            app.Quit();

            return destPath;
        }
    }
}
