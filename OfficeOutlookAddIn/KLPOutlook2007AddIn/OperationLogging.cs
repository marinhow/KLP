using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace KLPOutlookAddIn
{
    class OperationLogging
    {
        public string LogFilePath { get; set; }
        public StreamWriter writer;

        public OperationLogging()
        {
            LogFilePath = GetRootFolder();
        }

        public void WriteLog(string filename, MetadataInfo info)
        {
            try
            {
                writer = new StreamWriter(Path.Combine(LogFilePath, "OperationLog.txt"), true);
                writer.WriteLine(DateTime.Now.ToString());
                writer.WriteLine("FileName: "+filename);
                writer.WriteLine("Index Key: "+info.Indekseringsnokkel);
                writer.WriteLine("Fødselsnummer: " + info.Fodselsnr);
                writer.WriteLine("Document Code: "+info.Dokumentkode);
                writer.WriteLine("Document Code Descrition: "+ info.DokumentkodeBeskrivelse);
                writer.WriteLine(" ");
                writer.Close();
                
            }catch(Exception exp)
            {
                
            }
        }

        public void WriteErrorLog(string errorMessage)
        {
            try
            {
                writer = new StreamWriter(Path.Combine(LogFilePath, "OperationLog.txt"), true);
                writer.WriteLine(DateTime.Now.ToString());

                writer.WriteLine(errorMessage);

                writer.WriteLine(" ");
                writer.Close();

            }
            catch (Exception exp)
            {

            }
        }

        public void endLogEntry()
        {
            try
            {
                writer = new StreamWriter(Path.Combine(LogFilePath, "OperationLog.txt"), true);
                writer.WriteLine(" -------------------------------------- ");
                writer.Close();

            }
            catch (Exception exp)
            {

            }
        }

        private string GetRootFolder()
        {
            string rootFolder = @"C:\";
            rootFolder = Path.Combine(rootFolder, "DocTimTempFiles");

            try
            {
                if (!Directory.Exists(rootFolder))
                    Directory.CreateDirectory(rootFolder);
            }
            catch (System.UnauthorizedAccessException) { }

            return rootFolder;
        }
    }
}
