using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordConverter.Utils
{
    class WordUtils
    {
        public static void ConvertToImage(String filePath, String destpath, System.Drawing.Imaging.ImageFormat format)
        {
            // file exist & check type
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.EndsWith("doc") && !filePath.EndsWith("docx"))
                throw new Exception("文件非 word 类型");

            // open file
            object FileName = (object)filePath;
            object ReadOnly = (object)true;  
            object PrintToFile = (object)true;
            object OutPutFiletemp = (object)destpath + DateTime.Now.Ticks.ToString() + ".tiff";
            ApplicationClass app = new ApplicationClass();
            Document doc = app.Documents.Open(ref FileName, ReadOnly: ref ReadOnly);
            string defaultPrinter = app.ActivePrinter;
            app.ActivePrinter = "Microsoft Print to PDF";
            app.PrintOut(FileName: ref FileName,
                              PrintToFile: ref PrintToFile,
                              OutputFileName: ref OutPutFiletemp);
            //doc.Close();
            //app.Documents.Close();
            app.Documents[1].Close();
            app.ActivePrinter = defaultPrinter;
            app.Quit();
            System.Drawing.Image img = System.Drawing.Image.FromFile(OutPutFiletemp.ToString());
            img.Save(destpath, format);
            img.Dispose();
            File.Delete(OutPutFiletemp.ToString());
        }

        public static void ConvertToPDFWithPrinter(String filePath, String destpath)
        {
            // file exist & check type
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.EndsWith("doc") && !filePath.EndsWith("docx"))
                throw new Exception("文件非 word 类型");

            Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
            Type wordType = word.GetType();
            Microsoft.Office.Interop.Word.Documents docs = word.Documents;
            Type docsType = docs.GetType();
            Microsoft.Office.Interop.Word.Document doc = (Microsoft.Office.Interop.Word.Document)docsType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { filePath, true, true });
            doc.Application.ActivePrinter = "Microsoft Print to PDF";
            Type docType = doc.GetType();
            docType.InvokeMember("PrintOut", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { false, false, Microsoft.Office.Interop.Word.WdPrintOutRange.wdPrintAllDocument, destpath });
            wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, word, null);
        }

        public static void ConvertToPDF(String filePath, String destpath)
        {
            // file exist & check type
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.EndsWith("doc") && !filePath.EndsWith("docx"))
                throw new Exception("文件非 word 类型");

            Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(filePath);
            doc.ExportAsFixedFormat(destpath, WdExportFormat.wdExportFormatPDF);

            // rename
            List<String> list = new List<string>();
            String[] fileNames = Directory.GetFiles(destpath);
            foreach (var fileName in fileNames)
            {
                String newFileName = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + Path.GetExtension(fileName));
                File.Move(fileName, newFileName);
                list.Add(newFileName);
            }

        }
    }
}
