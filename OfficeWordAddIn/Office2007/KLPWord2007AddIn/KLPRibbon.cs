using System;
using System.Collections;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Security.Principal;
using System.Xml;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Windows.Forms;

namespace KLPWordAddIn
{
    public partial class KLPRibbon
    {
        private ErrorLogging _errorLogging;
        private OperationLogging _operationLogging;

        private void KLPRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //PerformBackupFolderMaintainance();
            _errorLogging = new ErrorLogging();
            _operationLogging = new OperationLogging();
        }

        #region Help Methods

        //Get File Unique Name
        private string getUniqueFileName()
        {
            
            string filename = DateTime.Now.Ticks.ToString(CultureInfo.InvariantCulture) + "_" +
                Globals.ThisAddIn.Application.ActiveDocument.Name.Split(new char[] { '.' })[0];

            return filename;
        }

        //Get Children Form data to insert into the Metadata file
        private MetadataInfo getMetadataInfoFromForm()
        {

            MetadataInfo metadata = null;
            try
            {
                StringDictionary tiaValue = GetTiaValues();
                MetadataForm form = new MetadataForm(tiaValue);
                form.ShowDialog();

                metadata = new MetadataInfo(form.Ankomstdato, form.Skadenummer, form.Dokumentkode, form.DokumentkodeBeskrivelse,
                    form.Dokumentbeskrivelse, form.Fodselsnr, form.DokAnkomstStatus, form.ExternalLink, tiaValue, form.Folder,form.validated);
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return metadata;
        }

        private StringDictionary GetTiaValues()
        {

            StringDictionary tiaValues = new StringDictionary();

            try
            {

                Document document = Globals.ThisAddIn.Application.ActiveDocument;
                string text = string.Empty;


                if (document.Sections.Count > 0)
                {
                    Range range = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    range.TextRetrievalMode.IncludeHiddenText = true;

                    text = range.Text;

                    string[] tags = text.Replace("<", "").Replace(">", "").Split(new string[] {","},
                                                                                 StringSplitOptions.RemoveEmptyEntries);
                    if (tags.Count() > 1)
                    {

                        foreach (var tag in tags)
                        {
                            string[] vals = tag.Split(new char[] {':'});
                            tiaValues.Add(vals[0].Trim(), vals[1].TrimEnd().TrimStart());
                        }
                    }
                    else
                    {
                        tiaValues.Add("TIACLACNO", GetSkadenummerFromDocument());
                    }
                }
                else
                {
                    tiaValues.Add("TIACLACNO", GetSkadenummerFromDocument());
                }

            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return tiaValues;
        }


        private string GetSkadenummerFromDocument()
        {
            string skadenummer = string.Empty;
            string skadenummerFinal = string.Empty;
            string searchQuery = "VÅR REF";

            try
            {

                Microsoft.Office.Interop.Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

                if (document.Content.Text.Contains(searchQuery))
                {
                    Microsoft.Office.Interop.Word.Range range = document.Content;
                    range.WholeStory();
                    range.Find.ClearFormatting();
                    range.Find.Text = searchQuery;

                    int FirstChr = range.Text.IndexOf(searchQuery);

                    skadenummer = document.Content.Text.Substring(FirstChr + 8, 10).TrimStart().TrimEnd();

                    

                    skadenummerFinal =
                        skadenummer.Where(char.IsDigit).Aggregate(skadenummerFinal,
                        (current, item) => current + item);

                    if (skadenummerFinal.Length > 6)
                        skadenummerFinal = "-1";
                }

            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return skadenummerFinal;
        }

        #endregion

        #region IO Operations

        //Perform backup folder maintainance (clean files older than 5 days)
        private void PerformBackupFolderMaintainance()
        {
            try
            {

                if (Directory.Exists(GetBackupFolder()))
                {
                    string[] filePaths = Directory.GetFiles(GetBackupFolder());

                    if (filePaths.Length > 0)
                        foreach (string file in filePaths)
                            if (File.GetCreationTime(file).CompareTo((DateTime.Now).AddDays(-5)) < 0)
                                File.Delete(file);
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

        }

        //Clean root folder and move older files to backup folder
        private void CleanRootFolder()
        {
            try
            {
                if (Directory.Exists(GetRootFolder()))
                {
                    string[] filePaths = Directory.GetFiles(GetRootFolder());

                    if (filePaths.Length > 0)
                        foreach (string file in filePaths)
                            if (!File.Exists(Path.Combine(GetBackupFolder(), "Backup_" + System.IO.Path.GetFileName(file))))
                                File.Move(file, Path.Combine(GetBackupFolder(), "Backup_" + System.IO.Path.GetFileName(file)));
                            else
                            {
                                File.Delete(Path.Combine(GetBackupFolder(), "Backup_" + System.IO.Path.GetFileName(file)));
                                File.Move(file, Path.Combine(GetBackupFolder(), "Backup_" + System.IO.Path.GetFileName(file)));
                            }
                }
            }
            catch (System.Exception exception) {
                _errorLogging.WriteLog(exception);
            }
        }

        //Save the new PDF file generated to the root folder
        private void SaveNewFile()
        {
            //1 - Get Metadata Info so store in the XML file
            MetadataInfo metadataInfo = getMetadataInfoFromForm();

            if (metadataInfo.Validated)
            {
                //2 - Get Unique File Name
                string uniqueFileName = getUniqueFileName();

                //3 - Save the Metadata XML file
                SaveToXML(uniqueFileName, metadataInfo);

                //4 - Save the PDF file
                SaveToPDF(uniqueFileName);

                _operationLogging.WriteLog(Path.Combine(GetRootFolder(), uniqueFileName + ".pdf"),metadataInfo);
                MessageBox.Show("Konvertering til PDF OK!", "KLP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SaveToPDF(string uniqueFileName)
        {
            try
            {

                Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                    Path.Combine(GetTempFolder(), uniqueFileName),
                    Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false);

                File.Move(Path.Combine(GetTempFolder(), uniqueFileName+".pdf"), Path.Combine(GetRootFolder(), uniqueFileName+".pdf"));
                

            }
            catch (Exception exception)
            {
                _errorLogging.WriteLog(exception);
                _operationLogging.WriteErrorLog(exception.Message);
            }
        }

        private void SaveToXML(string uniqueFileName, MetadataInfo metadataInfo)
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

                //grandChildElement = xml.CreateElement("Mottaker");
                //grandChildElement.InnerText = metadataInfo.tiaValues["Mottaker"];
                //childElement.AppendChild(grandChildElement);


                //metadataInfo.tiaValues.Remove("tiaclalno");
                //metadataInfo.tiaValues.Remove("tiareqpgm");


                foreach (DictionaryEntry tiaValue in metadataInfo.tiaValues)
                {
                    if (tiaValue.Key.ToString().ToLower().Equals("tiarecino"))
                    {
                        grandChildElement = xml.CreateElement("Mottaker");
                        grandChildElement.InnerText = tiaValue.Value.ToString();
                        childElement.AppendChild(grandChildElement);
                    }
                }

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

                //save to file
                xml.Save(Path.Combine(GetRootFolder(), uniqueFileName + ".pdf.xml"));
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        //Get the root folder
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

        //Get the backup folder
        private string GetBackupFolder()
        {
            string backupFolder = Path.Combine(GetRootFolder(), "Backups");

            try
            {
                if (!Directory.Exists(backupFolder))
                    Directory.CreateDirectory(backupFolder);
            }
            catch (System.UnauthorizedAccessException) { MessageBox.Show(Properties.Resources.GetFolderError); }

            return backupFolder;
        }

        #endregion

        #region ButtonEventHandlers

        private void exportToPDFButton_Click(object sender, RibbonControlEventArgs e)
        {
            export();
        }

        private void export()
        {
            //Clean Root Folder or move existing files to the Backup Folder
            //CleanRootFolder();

            //Save the New File
            SaveNewFile();

            //Perform backup folder maintainance
            //PerformBackupFolderMaintainance();

        }

        #endregion

    }
}
