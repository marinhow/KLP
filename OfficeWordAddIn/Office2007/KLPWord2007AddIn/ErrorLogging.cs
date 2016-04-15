using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace KLPWordAddIn
{
    class ErrorLogging
    {
        public string LogFilePath { get; set; }
        public StreamWriter writer;

        public ErrorLogging()
        {
            LogFilePath = GetRootFolder();
        }

        public void WriteLog(System.Exception exception)
        {
            try
            {
                writer = new StreamWriter(Path.Combine(LogFilePath, "ErrorLog.txt"), true);
                writer.WriteLine(DateTime.Now.ToString("dddd, MMMM dd, yyyy h:mm tt") + " -> " + exception.Message);
                writer.WriteLine("Source -> " + exception.Source);
                writer.WriteLine("StackTrace -> " + exception.StackTrace);
                writer.Close();

            }
            catch (Exception exp)
            {

            }
        }

        private string GetRootFolder()
        {
            string rootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            rootFolder = Path.Combine(rootFolder, "KLPOfficeAddInsErrorLogs");

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
