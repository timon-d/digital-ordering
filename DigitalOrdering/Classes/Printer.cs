using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    class Printer
    {
        private static String printServer = "ipm-int";
        private static String folderPath = @"C:\Bestellung\PrintTemp\PDF\";
        private String printer = "Bestellungen";
        private String printSharePath;

        public Printer() 
        {
            SetPrintSharePath();
        }

        private void SetPrintSharePath()
        {
            this.printSharePath = @"\\" + printServer + @"\" + printer + @"\";
        }


        public void PrintFile(String fileName, Boolean convertToPostScriptBeforePrint)
        {
            string toPrint = folderPath + fileName;
            if (convertToPostScriptBeforePrint)
            {
                PdfToPostScript converter = new PdfToPostScript(folderPath, fileName);
                toPrint = converter.ConvertPdfToPostScript();
                fileName = Path.GetFileName(toPrint);
            }
            File.Copy(toPrint, printSharePath + fileName);
        }
    }
}
