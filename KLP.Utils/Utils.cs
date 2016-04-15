using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace KLP.Utils
{
    public static class Utils
    {
        //Combine 2 PDF into 1
        public static void CombineMultiplePDFFiles(string mainFileName, string secondaryFile)
        {
            var filename1 = mainFileName;
            var filename2 = secondaryFile;


            //// Open the input files
            //PdfDocument inputDocument1 = PdfReader.Open(filename1, PdfDocumentOpenMode.Import);
            //PdfDocument inputDocument2 = PdfReader.Open(filename2, PdfDocumentOpenMode.Import);

            // Open the input files FIX
            PdfDocument inputDocument1 = CompatiblePdfReader.Open(filename1, PdfDocumentOpenMode.Import);
            PdfDocument inputDocument2 = CompatiblePdfReader.Open(filename2, PdfDocumentOpenMode.Import);

            // Create the output document
            PdfDocument outputDocument = new PdfDocument();


            // Show consecutive pages facing. Requires Acrobat 5 or higher.
            outputDocument.PageLayout = PdfPageLayout.OneColumn;

            for (int i = 0; i < inputDocument1.Pages.Count; i++)
            {
                outputDocument.AddPage(inputDocument1.Pages[i]);
            }

            for (int i = 0; i < inputDocument2.Pages.Count; i++)
            {
                outputDocument.AddPage(inputDocument2.Pages[i]);
            }


            // Save the document...
            var finalFilename = filename1;
            outputDocument.Save(finalFilename);
            
        }


        //Handle Special Characters in string
        public static string HandleSpecialCharacters(string original)
        {
            var final = original.Replace("<", "").Replace(">", "").Replace(":", "").Replace("\"", "").Replace(@"\", "").
                Replace("/", "").Replace("|", "").Replace("?", "").Replace("*", "").Replace("\"", "").Replace("@", "");
            
            return final;

        }


        #region Configuration

        public static string GetActiveEnvironment(string configurationFile)
        {
            string c = "";

            try
            {
                XDocument xDoc = XDocument.Load(configurationFile);
                //XDocument xDoc = XDocument.Parse(resourceName);

                var result = from b in xDoc.Descendants("Environment")
                             select new
                                        {
                                            environmentName = b.Element("EnvironmentName").Value,
                                            active = b.Element("Active").Value,
                                            server = b.Element("Server").Value
                                        };

                var activeServer = (from servers in result
                                    where servers.active.ToLower() == "true"
                                    select servers).First();


                c = activeServer.server;

            }catch(Exception exp)
            {
                throw exp;
            }

            return c;
        }

        //LOCAL-DEV ENVIRONMENT
        public static string CONFIGURATIONXML = @"C:\DocTimTempFiles\EnvironmentConfiguration.xml";
        public static string DOCUMENTCODESXML = @"C:\DocTimTempFiles\DocumentCodes.xml";
        public static string TEMPFOLDER = @"C:\DocTimTempFiles";

        ////KLP ENVIRONMENT
        //public static string CONFIGURATIONXML = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory,
        //                                                     "EnvironmentConfiguration.xml");
        //public static string DOCUMENTCODESXML = @"\\FIL1BP\filenet-skade-doclist\DocumentCodes.xml";
        //public static string TEMPFOLDER = @"C:\DocTimTempFiles";


        #endregion

        #region DocumentCodes
        //Fødselsnummer Validator Method
        public static Tuple<bool, string> CheckFødselsnummer(string number)
        {
            string result = string.Empty;

            //sjekker lengden 
            if (number.Length != 11)
            {
                result = "Fødselsnummer er ikke korrekt!";
                return new Tuple<bool, string>(false, result);
            }

            //D-nummer transformation
            if (Convert.ToInt32(number.Substring(0, 1)) > 3)
            {
                number = (Convert.ToInt32(number.Substring(0, 1)) - 4).ToString() + number.Substring(1, number.Length - 1);
            }

            long foo;
            //Sjekker om det faktisk er et tall det som er skrevet inn
            //Checks if the input is just number
            if (!Int64.TryParse(number, out foo))
            {
                result = "Fødselsnummer er ikke korrekt!";
                return new Tuple<bool, string>(false, result);
            }

            //bare fordi jeg ikke gidder å skrive inputssn.Text hele tiden.
            string num = number;

            //Deler opp i litt mer håndterlige deler
            string day = num.Substring(0, 2);
            string month = num.Substring(2, 2);
            string year = num.Substring(4, 2);
            string individual = num.Substring(6, 3);
            string k1 = num.Substring(9, 1);
            string k2 = num.Substring(10, 1);



            //Her kan du validere litt dato. Jeg gidder ikke skrive det nå, men du kan ta med sjekk på antall dager i den 
            // aktuelle mnd, osv... Husk at i et D-nummer legges det til 40 på dagen.

            //dersom individnummeret er mellom 500 og 750 er vedkommende født mellom 1855 og 1899
            if (Convert.ToInt32(individual) > 500 && Convert.ToInt32(individual) < 750)
                result = "Er du sikker på at denne personen er født FØR 1900? - ";

            //individnummerets tredje siffer bestemmer kjønn. partall: kvinne, oddetall: mann
            if (Convert.ToInt32(individual.Substring(2, 1)) % 2 == 0)
                result = "Dette er en kvinne - ";
            else
                result = "Dette er en mann - ";

            //Deler opp alle sifferne i hver sin int (bare for å gjøre utregningen lettere)
            int d1 = Convert.ToInt32(day.Substring(0, 1));
            int d2 = Convert.ToInt32(day.Substring(1, 1));
            int m1 = Convert.ToInt32(month.Substring(0, 1));
            int m2 = Convert.ToInt32(month.Substring(1, 1));
            int y1 = Convert.ToInt32(year.Substring(0, 1));
            int y2 = Convert.ToInt32(year.Substring(1, 1));
            int i1 = Convert.ToInt32(individual.Substring(0, 1));
            int i2 = Convert.ToInt32(individual.Substring(1, 1));
            int i3 = Convert.ToInt32(individual.Substring(2, 1));

            //Regner ut k1 (første kontrollsiffer)
            int k1Calculated = 11 -
                               (((3 * d1) + (7 * d2) + (6 * m1) + (1 * m2) + (8 * y1) + (9 * y2) + (4 * i1) + (5 * i2) + (2 * i3)) % 11);
            k1Calculated = (k1Calculated == 11 ? 0 : k1Calculated);

            //fødselsnummer som ville gitt k1 = 10 tildeles ikke
            if (k1Calculated == 10)
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k1 kan aldri bli 10";
                return new Tuple<bool, string>(false, result);
            }

            //Sjekker om den utregnede k1 er den samme som den som er tastet inn
            if (k1Calculated != Convert.ToInt32(k1))
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k1 feil!";
                return new Tuple<bool, string>(false, result);
            }

            //regner ut k2 (andre kontrolliffer)
            int k2Calculated = 11 -
                               (((5 * d1) + (4 * d2) + (3 * m1) + (2 * m2) + (7 * y1) + (6 * y2) + (5 * i1) + (4 * i2) + (3 * i3) +
                                 (2 * k1Calculated)) % 11);
            k2Calculated = (k2Calculated == 11 ? 0 : k2Calculated);

            //fødselsnummer som ville gitt k2 = 10 tildeles ikke
            if (k2Calculated == 10)
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k2 kan aldri bli 10";
                return new Tuple<bool, string>(false, result);
            }

            //sjekker om den utregnede k2 er den samme som den som er tatet inn
            if (k2Calculated != Convert.ToInt32(k2))
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k2 feil";
                return new Tuple<bool, string>(false, result);
            }

            //siden alle feil returnerer test-funksjonen, så har den aldrå nå passert :)
            result += "Passerte alle tester";
            return new Tuple<bool, string>(true, result);
        }

        //Read Product Codes File
        public static List<Tuple<string, string>> Read(string resourceName)
        {
            XDocument xDoc = XDocument.Load(resourceName);
            //XDocument xDoc = XDocument.Parse(resourceName);

            var result = from b in xDoc.Descendants("DocumentCodes")
                         select new
                         {
                             code = b.Element("BREVKODE").Value,
                             description = b.Element("Beskrivelse").Value.Replace("/", "|").Replace(@"\", "|")
                         };

            List<Tuple<string, string>> productCodes = new List<Tuple<string, string>>();

            foreach (var row in result)
                productCodes.Add(new Tuple<string, string>(row.code, row.description));


            return productCodes;
        }
        #endregion


        #region Special Folders
        public static string GetRootFolder()
        {
            //string rootFolder = @"C:\TESTESPARAAPAGAR\";

            //string rootFolder = KLP.Utils.Utils.GetActiveEnvironment(Properties.Settings.Default.EnvironmentConfigurationFile);
            string rootFolder = KLP.Utils.Utils.GetActiveEnvironment(KLP.Utils.Utils.CONFIGURATIONXML);

            try
            {
                if (!Directory.Exists(rootFolder))
                    Directory.CreateDirectory(rootFolder);
            }
            catch (System.UnauthorizedAccessException exception)
            {
                throw exception;
            }

            return rootFolder;
        }

        public static string GetTempFolder()
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
                throw exception;
            }

            return tempFolder;
        }
        #endregion

        public static void SaveMetadataToXml(string uniqueFileName, MetadataInfo metadataInfo, string fullPath)
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
                grandChildElement.InnerText = metadataInfo.ExternalLink;
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
                throw exception;
            }
        }

        public static string ConvertToPdf(string file)
        {
            if (file.ToLower().EndsWith(".xls") || file.ToLower().EndsWith(".xlsx"))
                return ExcelToPdf(file);

            if (file.ToLower().EndsWith(".doc") || file.ToLower().EndsWith(".docx") || file.ToLower().EndsWith(".txt") ||
                    file.ToLower().EndsWith(".jpeg") || file.ToLower().EndsWith(".jpg") || file.ToLower().EndsWith(".gif") ||
                    file.ToLower().EndsWith(".png") || file.ToLower().EndsWith(".bmp") || file.ToLower().EndsWith(".rtf"))
                return WordToPdf(file);


            return "";
        }

        private static string WordToPdf(string file)
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
                throw exception;
                return "";
            }
        }

        private static string ExcelToPdf(string file)
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
                throw exception;
                return "";
            }


        }

        public static string ImageToWord(string ImageFile)
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
    }
}
    