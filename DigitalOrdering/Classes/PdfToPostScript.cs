using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    class PdfToPostScript
    {
        
        private static String gsProgramPath = @"C:\Program Files\gs\gs9.20\bin\gswin64c.exe";
        private static String outputFolder = @"C:\Bestellung\PrintTemp\PS\";
        private static String logFolder = @"C:\Bestellung\Logs\PSPrints\";
        private static String logFileNameTail = "_GS_PSPrint.log";
        private static String conversionMode = "ps2write";
        private static String[] administrators = { "timon.dages@ipm.fraunhofer.de" };
        private String inputPath;
        private String outputPath;
        private String logFileName;
        private String logPath;
        private String conversionArgs;
        private Boolean successfulConversion;

        public PdfToPostScript(String folder, String fileName)
        {
            this.inputPath = folder + fileName;
            this.outputPath = outputFolder + fileName + ".ps";
            this.logFileName = fileName + logFileNameTail;
            this.logPath =  logFolder + this.logFileName;
            this.conversionArgs = " -sPAPERSIZE=a4 -dFIXEDMEDIA -dPDFFitPage -dBATCH -dNOPAUSE -sDEVICE=" + conversionMode + " -sCIDFMAP=lib/cidfmap -sOutputFile=" + this.outputPath + " -c \"<</BeginPage{0.9 0.9 scale 29.75 42.1 translate}>> setpagedevice\" -f " + inputPath + " ";
            //-sFONTPATH=C:/Windows/fonts -dEmbedAllFonts=true
        }
        public String ConvertPdfToPostScript()
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.Arguments = this.conversionArgs;
            startInfo.FileName = gsProgramPath;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardOutput = true;
            process = Process.Start(startInfo);
            
            //diese While-Schleife schreibt jede von GhostScript ausgegebene Zeile in die Stringliste "logs"
            List<string> logs = new List<string>();
            while (!process.StandardOutput.EndOfStream)
            {
                string line = process.StandardOutput.ReadLine();
                logs.Add(line);
            }
            string[] logText = logs.ToArray();
            File.WriteAllLines(this.logPath, logText);

            //Wenn der Prozess länger als drei Minuten läuft, wird er "gekillt"
            process.WaitForExit(30000);
            if (process.HasExited == false) process.Kill();
            this.successfulConversion = process.ExitCode == 0;
            if(!successfulConversion)
            {
                MailNotification notification = new MailNotification(administrators, "Fehler beim Konvertieren in PostScript (" + this.outputPath + ")", "");
                notification.InsertAttachment(File.ReadAllBytes(this.logPath), this.logFileName);
                notification.SendMailNotification();
            }
            return this.outputPath;
        }
    }
}
