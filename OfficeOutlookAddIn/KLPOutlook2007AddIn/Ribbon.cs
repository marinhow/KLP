using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Threading;


namespace KLPOutlookAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {

        private Office.IRibbonUI ribbon;
        private ErrorLogging _errorLogging;
        private OperationLogging _operationsLogging;

        public Ribbon()
        {
            _errorLogging = new ErrorLogging();
            _operationsLogging = new OperationLogging();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            Inspector insp = Globals.ThisAddIn.Application.ActiveInspector();
            
            return GetResourceText("KLPOutlookAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            CreateCustomFolder();
        }

        public void exportToPDFButton_Click(IRibbonControl control)
        {
            try
            {
                Inspector insp = Globals.ThisAddIn.Application.ActiveInspector();

                MailItem mailItem = insp.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;

                if (mailItem.Recipients.Count > 0)
                {
                    Thread t = new Thread(WorkingThread);
                    t.SetApartmentState(ApartmentState.STA);
                    t.Start(insp);
                }
                else
                {
                    MessageBox.Show("Mottakeren mangler!", "KLP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

        }

        #endregion


        #region IO Operations

        private void WorkingThread(object insp)
        {

            try
            {
                Inspector inspector = (Inspector)insp;
                MetadataInfo metadata = GetMetadataInfoFromForm();

                if (metadata.Validated)
                {


                    if (inspector != null)
                    {
                        try
                        {
                            MailItem mailItem = inspector.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;
                            
                            FileHandler(mailItem, metadata);

                        }
                        catch (System.Exception exception)
                        {
                            _errorLogging.WriteLog(exception);
                        }
                    }
                }

            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private void FileHandler(MailItem mailItem, MetadataInfo metadata)
        {
            var operationCheckpoint = string.Empty;

            var indexingKeys = metadata.Indekseringsnokkel.Split(new string[] { "," },StringSplitOptions.RemoveEmptyEntries);

            if (indexingKeys.Length > 10)
            {
                MessageBox.Show("FEIL - Max skadenumre er 10");
                return;
            }


            try
            {
                foreach (var indexKey in indexingKeys)
                {
                    metadata.Indekseringsnokkel = indexKey;
                    
                    operationCheckpoint = "Started operation";
                    string timeStamp = DateTime.Now.Ticks.ToString();

                    string docFileName = string.Empty;

                    if (mailItem.Subject == null)
                        docFileName = timeStamp + "_epostutenemne" + ".doc".Trim();

                    else if (mailItem.Subject.Length > 20)
                        docFileName = timeStamp + "_" + KLP.Utils.Utils.HandleSpecialCharacters(mailItem.Subject) + ".doc".Trim();

                    else
                        docFileName = timeStamp + "_" + KLP.Utils.Utils.HandleSpecialCharacters(mailItem.Subject) + ".doc".Trim();

                    string filePath = Path.Combine(GetRootFolder(), docFileName);
                    string tempFilePath = Path.Combine(GetTempFolder(), docFileName);

                    operationCheckpoint = string.Format("File name and path completed: {0} {1}", filePath, tempFilePath);

                    //Handle email storage in PDF format and XML metadata file
                    mailItem.SaveAs(tempFilePath, OlSaveAsType.olDoc);
                    var emailPDFFile = ConvertToPdf(docFileName);

                    operationCheckpoint = string.Format("Email converted to PDF file: {0}", emailPDFFile);

                    //Handle attachments storage in PDF format and XML metadata file
                    Attachments attachments = mailItem.Attachments;
                    foreach (Attachment i in attachments)
                    {
                        timeStamp = DateTime.Now.Ticks.ToString();
                        if (i.FileName.ToLower().EndsWith(".pdf"))
                        {
                            var path = Path.Combine(GetTempFolder(), timeStamp + "_" + i.FileName);
                            i.SaveAsFile(path);

                            KLP.Utils.Utils.CombineMultiplePDFFiles(emailPDFFile, path);
                            File.Delete(path);
                        }

                        else if (i.FileName.ToLower().EndsWith(".doc") || i.FileName.ToLower().EndsWith(".docx") || i.FileName.ToLower().EndsWith(".xls")
                           || i.FileName.ToLower().EndsWith(".xlsx") || i.FileName.ToLower().EndsWith(".ppt") || i.FileName.ToLower().EndsWith(".pptx") ||
                           i.FileName.ToLower().EndsWith(".txt") || i.FileName.ToLower().EndsWith(".rtf"))
                        {
                            var path = Path.Combine(GetTempFolder(), timeStamp + "_" + i.FileName);
                            i.SaveAsFile(path);
                            var pdfFilePath = ConvertToPdf(timeStamp + "_" + i.FileName);

                            KLP.Utils.Utils.CombineMultiplePDFFiles(emailPDFFile, pdfFilePath);
                            File.Delete(pdfFilePath);
                        }
                        else if (i.FileName.ToLower().EndsWith(".jpeg") || i.FileName.ToLower().EndsWith(".jpg") || i.FileName.ToLower().EndsWith(".gif") ||
                            i.FileName.ToLower().EndsWith(".png") || i.FileName.ToLower().EndsWith(".bmp"))
                        {
                            if (!i.FileName.ToLower().Contains("image00"))
                            {
                                var path = Path.Combine(GetTempFolder(), timeStamp + "_" + i.FileName);
                                i.SaveAsFile(path);
                                var pdfFilePath = ImageToWord(timeStamp + "_" + i.FileName);

                                KLP.Utils.Utils.CombineMultiplePDFFiles(emailPDFFile, pdfFilePath);
                                File.Delete(pdfFilePath);

                            }
                        }

                        else if (i.FileName.ToLower().EndsWith(".tif"))
                        {
                            string inputFile = Path.Combine(GetTempFolder(), timeStamp + "_" + i.FileName);
                            i.SaveAsFile(inputFile);
                            var fileNameNoExt = Path.GetFileNameWithoutExtension(inputFile);
                            string outputFile = Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf");

                            string arguments = "-o " + "\"" + outputFile + "\"" + " " + "\"" + inputFile + "\"";

                            Process.Start(Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Tiff2Pdf.exe"), arguments);

                            while (!File.Exists(outputFile))
                            {
                                Thread.Sleep(500);
                            }

                            KLP.Utils.Utils.CombineMultiplePDFFiles(emailPDFFile, outputFile);
                            File.Delete(inputFile);
                            File.Delete(outputFile);
                        }



                        else
                        {
                            MessageBox.Show(
                                "The attached file: " + i.FileName +
                                " cannot be saved to pdf! File extension not supported!", "KLP AddIn");
                        }
                    }

                    operationCheckpoint = string.Format("Attachments saved");

                    File.Move(emailPDFFile, Path.Combine(GetRootFolder(), Path.GetFileNameWithoutExtension(emailPDFFile) + ".pdf"));
                    operationCheckpoint = string.Format("Files saved");
                    SaveMetadataToXml(docFileName, metadata, filePath);
                    operationCheckpoint = string.Format("XML saved");
                    _operationsLogging.WriteLog(emailPDFFile, metadata);
                }

                if (mailItem.EntryID == null)
                {
                    mailItem.Send();
                    MessageBox.Show("Konvertering til PDF OK! Epost sendt.", "KLP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                    MessageBox.Show("Konvertering til PDF OK!", "KLP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                

            }
            catch (System.Exception exception)
            {
                _operationsLogging.WriteErrorLog(operationCheckpoint);
                _operationsLogging.WriteErrorLog(exception.Message);
                _errorLogging.WriteLog(exception);
                MessageBox.Show("FEIL - se logg!");
            }
            finally
            {
                _operationsLogging.endLogEntry();

                //CleanRootFolder();
            }
        }

        private string ConvertToPdf(string file)
        {
            if (file.ToLower().EndsWith(".xls") || file.ToLower().EndsWith(".xlsx"))
                return ExcelToPdf(file);

            if (file.ToLower().EndsWith(".doc") || file.ToLower().EndsWith(".docx") || file.ToLower().EndsWith(".txt") ||
                    file.ToLower().EndsWith(".jpeg") || file.ToLower().EndsWith(".jpg") || file.ToLower().EndsWith(".gif") ||
                    file.ToLower().EndsWith(".png") || file.ToLower().EndsWith(".bmp") || file.ToLower().EndsWith(".rtf"))
                return WordToPdf(file);


            return "";
        }

        private string WordToPdf(string file)
        {
            try
            {

                // Create a new Microsoft Word application object
                var word = new Microsoft.Office.Interop.Word.Application();

                // C# doesn't have optional arguments so we'll need a dummy value
                object oMissing = System.Reflection.Missing.Value;

                // Cast as Object for word Open method
                Object filename = (Object)Path.Combine(GetTempFolder(), file);

                // Use the dummy value as a placeholder for optional arguments
                Document doc = word.Documents.Open(ref filename, ref oMissing,
                                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                var fileNameNoExt = Path.GetFileNameWithoutExtension(Path.Combine(GetTempFolder(), file));

                object outputFileName = (Object)Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                           ref fileFormat, ref oMissing, ref oMissing,
                           ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                           ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                           ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.                
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;

                word.Quit(ref oMissing, ref oMissing, ref oMissing);

                //File.Move(Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf"), Path.Combine(GetRootFolder(), fileNameNoExt + ".pdf"));

                //Delete the doc file after converted to PDF
                File.Delete(Path.Combine(GetTempFolder(), file));
                //File.Delete(Path.Combine(GetTempFolder(), file + ".pdf"));

                return outputFileName.ToString();

            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
                return "";
            }
        }

        private string ExcelToPdf(string file)
        {
            // Create a new Microsoft Word application object
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook doc = null;

            try
            {

                // C# doesn't have optional arguments so we'll need a dummy value
                object oMissing = System.Reflection.Missing.Value;

                // Cast as Object for word Open method
                string filename = Path.Combine(GetTempFolder(), file);

                // Use the dummy value as a placeholder for optional arguments
                doc = excel.Workbooks.Open(filename, oMissing,
                                                    oMissing, oMissing, oMissing, oMissing, oMissing,
                                                    oMissing, oMissing, oMissing);

                var fileNameNoExt = Path.GetFileNameWithoutExtension(Path.Combine(GetTempFolder(), file));
                string outputFileName = Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf");

                // Save document into PDF Format
                if (doc != null)
                    doc.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                        outputFileName, XlFixedFormatQuality.xlQualityStandard,
                        true, true, Type.Missing,
                        Type.Missing, false,
                        Type.Missing);

                doc.Close();
                excel.Quit();

                //File.Move(Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf"),
                //          Path.Combine(GetRootFolder(), fileNameNoExt + ".pdf"));

                //Delete the doc file after converted to PDF
                File.Delete(Path.Combine(GetTempFolder(), file));
                //File.Delete(Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf"));

                return outputFileName;

            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
                return "";
            }


        }

        private string ImageToWord(string ImageFile)
        {



            object missing = Type.Missing;
            object start = 0;
            object end = 0;

            // Create a new Microsoft Word application object
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = WordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Microsoft.Office.Interop.Word.Range rng = doc.Range(ref start, ref end);


            rng.InlineShapes.AddPicture(Path.Combine(GetTempFolder(), ImageFile), ref missing, ref missing,
                                        ref missing);

            var fileNameNoExt = Path.GetFileNameWithoutExtension(Path.Combine(GetTempFolder(), ImageFile));

            object outputFileName = (Object)Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf");

            object fileFormat = WdSaveFormat.wdFormatPDF;

            // Save document into PDF Format
            doc.SaveAs(ref outputFileName,
                       ref fileFormat, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing);

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)doc).Close(ref saveChanges, ref missing, ref missing);

            doc = null;
            WordApp.Quit(ref missing, ref missing, ref missing);

            //File.Move(Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf"),
            //          Path.Combine(GetRootFolder(), fileNameNoExt + ".pdf"));

            //Delete the doc file after converted to PDF
            //File.Delete(Path.Combine(GetTempFolder(), ImageFile));
            //File.Delete(Path.Combine(GetTempFolder(), fileNameNoExt + ".pdf"));

            File.Delete(Path.Combine(GetTempFolder(), ImageFile));

            return outputFileName.ToString();
        }

        private string GetRootFolder()
        {
            //string rootFolder = KLP.Utils.Utils.GetActiveEnvironment(Properties.Settings.Default.EnvironmentConfigurationFile);
            string rootFolder = KLP.Utils.Utils.GetActiveEnvironment(KLP.Utils.Utils.CONFIGURATIONXML);

            try
            {
                if (!Directory.Exists(rootFolder))
                    Directory.CreateDirectory(rootFolder);
            }
            catch (System.UnauthorizedAccessException exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return rootFolder;
        }

        private string GetTempFolder()
        {
            //string tempFolder = @"C:\";
            //tempFolder = Path.Combine(tempFolder, "DocTimTempFiles");
            string tempFolder = KLP.Utils.Utils.TEMPFOLDER;

            try
            {
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);
            }
            catch (System.UnauthorizedAccessException exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return tempFolder;
        }

        //Clean root folder from the .tmp Files
        //private void CleanRootFolder()
        //{
        //    var fileNetFolder = new DirectoryInfo(GetRootFolder());

        //    foreach (FileInfo file in fileNetFolder.GetFiles())
        //    {
        //        if (file.FullName.ToLower().EndsWith(".tmp"))
        //            file.Delete();
        //    }

        //}

        private void SaveMetadataToXml(string uniqueFileName, MetadataInfo metadataInfo, string fullPath)
        {

            try
            {

                var xml = new XmlDocument();
                XmlDeclaration xmldecl = xml.CreateXmlDeclaration("1.0", "ISO-8859-1", null);

                //Add the new node to the document.
                XmlElement root = xml.DocumentElement;
                xml.InsertBefore(xmldecl, root);

                XmlElement element;

                element = xml.CreateElement("document");
                xml.AppendChild(element);

                XmlElement childElement;
                XmlAttribute childAttr;

                #region Arkiv
                //Arkiv
                childElement = xml.CreateElement("arkiv");

                XmlElement grandChildElement;
                XmlAttribute grandChildAttr;

                grandChildElement = xml.CreateElement("objectstore");
                grandChildElement.InnerText = "Skade";
                childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("documentclass");
                grandChildElement.InnerText = "SkadeOppgjor";
                childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("documentfolder");
                grandChildElement.InnerText = @"\SkadeOppgjor";
                childElement.AppendChild(grandChildElement);

                element.AppendChild(childElement);

                #endregion

                #region Indexes
                //Indexes
                childElement = xml.CreateElement("indexes");

                //grandChildElement = xml.CreateElement("Ankomstdato");
                //grandChildAttr = xml.CreateAttribute("indexClass");
                //grandChildAttr.Value = "Date";
                //grandChildElement.Attributes.Append(grandChildAttr);
                //grandChildElement.InnerText = metadataInfo.Ankomstdato;
                //childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("Indekseringsnokkel");
                grandChildElement.InnerText = metadataInfo.Indekseringsnokkel;
                childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("Dokumentkode");
                grandChildElement.InnerText = metadataInfo.Dokumentkode.Trim();
                childElement.AppendChild(grandChildElement);

                //grandChildElement = xml.CreateElement("DokumentkodeBeskrivelse");
                //grandChildElement.InnerText = metadataInfo.DokumentkodeBeskrivelse;
                //childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("Dokumentbeskrivelse");
                grandChildElement.InnerText = metadataInfo.Dokumentbeskrivelse.Length < 100 ?
                                                metadataInfo.Dokumentbeskrivelse : metadataInfo.Dokumentbeskrivelse.Substring(0, 99);
                childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("Fodselsnr");
                grandChildElement.InnerText = metadataInfo.Fodselsnr;
                childElement.AppendChild(grandChildElement);

                grandChildElement = xml.CreateElement("DokAnkomstStatus");
                grandChildElement.InnerText = metadataInfo.DokAnkomstStatus;
                childElement.AppendChild(grandChildElement);


                grandChildElement = xml.CreateElement("ExternalLink");
                try
                {
                    grandChildElement.InnerText = "outlook:" + GetOutlookItemEntryID();//metadataInfo.ExternalLink;
                }
                catch (System.Exception exp)
                {
                    grandChildElement.InnerText = "outlook:";
                }

                childElement.AppendChild(grandChildElement);

                element.AppendChild(childElement);

                #region Log
                //Arkiv
                childElement = xml.CreateElement("logg");


                grandChildElement = xml.CreateElement("user");
                grandChildElement.InnerText = WindowsIdentity.GetCurrent().Name;
                childElement.AppendChild(grandChildElement);

                element.AppendChild(childElement);

                #endregion

                #endregion
                var fileNameNoExt = Path.GetFileNameWithoutExtension(Path.Combine(GetTempFolder(), uniqueFileName));

                if (!fileNameNoExt.ToLower().Contains("image001"))
                {
                    xml.Save(Path.Combine(GetRootFolder(), fileNameNoExt + ".pdf.xml"));
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private MetadataInfo GetMetadataInfoFromForm()
        {
            MetadataInfo metadata = null;
            try
            {
                Form1 form = new Form1(GetSkadenummerFromDocument());
                form.ShowDialog();

                metadata = new MetadataInfo(form.Ankomstdato, form.Skadenummer, form.Dokumentkode, form.DokumentkodeBeskrivelse,
                    form.Dokumentbeskrivelse, form.Fodselsnr, form.DokAnkomstStatus, form.ExternalLink, form.Folder, form.validated);
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return metadata;
        }


        private string GetSkadenummerFromDocument()
        {
            string skadenummer = string.Empty;
            string skadenummerFinal = string.Empty;
            string searchQuery = "VÅR REF";

            try
            {

            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return skadenummerFinal;
        }


        private void CreateCustomFolder()
        {
            var todayDate = DateTime.Now;

            var month = todayDate.ToString("MMM");
            var year = todayDate.Year.ToString();

            var app = Globals.ThisAddIn.Application;

            MAPIFolder inBox = (MAPIFolder)app.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            //Verify if Post Bygning or Post Motor
            MAPIFolder activeFolder = inBox;
            foreach (MAPIFolder subFolder in inBox.Folders)
            {
                if (subFolder.Name.Equals("Arkiv Bygning"))
                {
                    activeFolder = subFolder;
                    foreach (MAPIFolder yearFolder in subFolder.Folders)
                    {
                        if (yearFolder.Name.Equals(year))
                        {
                            activeFolder = yearFolder;
                            foreach (MAPIFolder monthFolder in yearFolder.Folders)
                            {
                                if (monthFolder.Name.Equals(month))
                                {
                                    activeFolder = monthFolder;
                                    break;
                                }
                            }
                        }
                    }  
                }
                else if (subFolder.Name.Equals("Arkiv Motor"))
                {
                    activeFolder = subFolder;
                    foreach (MAPIFolder yearFolder in subFolder.Folders)
                    {
                        if (yearFolder.Name.Equals(year))
                        {
                            activeFolder = yearFolder;
                            foreach (MAPIFolder monthFolder in yearFolder.Folders)
                            {
                                if (monthFolder.Name.Equals(month))
                                {
                                    activeFolder = monthFolder;
                                    break;
                                }
                            }
                        }
                    }

                }
            }

            MessageBox.Show("Current folder " + activeFolder.Name + ".");

            //string userName = (string)app.ActiveExplorer().Session.CurrentUser.Name;
            //MAPIFolder customFolder = null;
            //try
            //{
            //    customFolder = (MAPIFolder)inBox.Folders.Add(userName,OlDefaultFolders.olFolderInbox);
            //    MessageBox.Show("You have created a new folder named " + userName + ".");
            //    inBox.Folders[userName].Display();
            //}
            //catch (System.Exception ex)
            //{
            //    MessageBox.Show("The following error occurred: " + ex.Message);
            //}
        }

        private void MoveMailToFolder()
        {

        }

        #endregion

        #region Helpers

        private string GetOutlookItemEntryID()
        {
            Inspector insp = Globals.ThisAddIn.Application.ActiveInspector();
            MailItem mailItem = insp.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;

            return mailItem.EntryID;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion



    }
}
