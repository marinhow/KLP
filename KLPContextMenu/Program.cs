using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using KLP.Utils;

namespace KLPContextMenu
{
    static class Program
    {
        private static KLP.Utils.ErrorLogging _errorLogging;
        private static KLP.Utils.OperationLogging _operationsLogging;


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string [] args)
        {
            _errorLogging = new ErrorLogging();
            _operationsLogging = new OperationLogging();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1(""));
            
            HandleFile(args[0]);
           
          

        }

        static void HandleFile(string filePath)
        {
            try
            {

                var form = new Form1("");
                form.ShowDialog();

                var metadata = new MetadataInfo(form.Ankomstdato, form.Skadenummer, form.Dokumentkode,
                                                form.DokumentkodeBeskrivelse,
                                                form.Dokumentbeskrivelse, form.Fodselsnr, form.DokAnkomstStatus, form.ExternalLink,
                                                form.Folder,
                                                form.validated);
                if (metadata.Validated)
                {

                    var fileName = System.IO.Path.GetFileName(filePath);
                    var timeStamp = DateTime.Now.Ticks.ToString();

                    if (fileName.ToLower().EndsWith(".pdf"))
                    {
                        var path = Path.Combine(KLP.Utils.Utils.GetRootFolder(), timeStamp + "_" + fileName);
                        File.Copy(filePath, path);
                        KLP.Utils.Utils.SaveMetadataToXml(path, metadata, filePath);

                        MessageBox.Show("Overføring OK!", "KLP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        _operationsLogging.WriteLog(filePath, metadata);

                    }


                    else
                    {
                        MessageBox.Show(
                            "The attached file: " + filePath +
                            " cannot be saved to pdf! File extension not supported!", "KLP AddIn");
                    }


                }
            }catch(Exception exception)
            {
                MessageBox.Show("FEIL - se logg!");
                _errorLogging.WriteLog(exception);
                _operationsLogging.WriteErrorLog(exception.Message);
            }



        }

        

    }
}
