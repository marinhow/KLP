using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace KLPWordAddIn
{
    public partial class MetadataForm : Form
    {

        public string Ankomstdato { get; set; }
        public string Skadenummer { get; set; }
        public string Dokumentkode { get; set; }
        public string DokumentkodeBeskrivelse { get; set; }
        public string Dokumentbeskrivelse { get; set; }
        public string Fodselsnr { get; set; }
        public string DokAnkomstStatus { get; set; }
        public string ExternalLink { get; set; }
        public string Folder { get; set; }
        public bool validated;

        private ErrorLogging _errorLogging;

        private List<Tuple<string, string>> productCodes;

        private int comboSize;

        public MetadataForm(StringDictionary tiaValues)
        {
            InitializeComponent();
            _errorLogging = new ErrorLogging();
            //textBox5.Text = Properties.Settings.Default.FileServerPath;
            textBox5.Text = KLP.Utils.Utils.GetActiveEnvironment(KLP.Utils.Utils.CONFIGURATIONXML);
            textBox4.Text = "WF";
            comboSize = comboBox1.Width;

            AddProductCodes();
            try
            {

                if (tiaValues["TIACLACNO"].Equals("-1"))
                {
                    MessageBox.Show("Skadenummeret er ugyldig!");
                }
                else
                {
                    textBox1.Text = tiaValues["TIACLACNO"];
                }

                var docID = tiaValues["TIAREQPGM"].ToString();
                string docDesc = productCodes.Find(id => id.Item1.Equals(tiaValues["TIAREQPGM"].ToString())).Item2;

                comboBox1.Text = docID+"   ";
                DocumentCodeDescription.Text = docDesc;

               
                tiaValues.Remove("tiadocankst");
                tiaValues.Remove("TIADOCANKST");
                tiaValues.Remove("TIACLACNO");
                tiaValues.Remove("tiaclacno");
                //tiaValues.Remove("TIARECINO");
                tiaValues.Remove("TIAPGMDESC");
                tiaValues.Remove("TIAREQPGM");
                


            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
                tiaValues.Remove("TIACLACNO");
                tiaValues.Remove("tiaclacno");
            }

            validated = false;

        }

        private void AddProductCodes()
        {
            try
            {
                //List<Tuple<string, string>> productCodes = new Helper().Read(@"C:\DocumentCodes.xml");
                productCodes = KLP.Utils.Utils.Read(KLP.Utils.Utils.DOCUMENTCODESXML);


                foreach (var productCode in productCodes)
                    comboBox1.Items.Add(productCode.Item1 + " - " + productCode.Item2);

                comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormValidation())
                {
                    Ankomstdato = dateTimePicker1.Value.ToString("yyyy-MM-dd hh:mm:ss");
                    Skadenummer = textBox1.Text;
                    Dokumentkode = comboBox1.Text.TrimEnd();
                    DokumentkodeBeskrivelse = DocumentCodeDescription.Text;
                    //Dokumentkode = textBox2.Text;
                    //Dokumentbeskrivelse = "Invoice";
                    Dokumentbeskrivelse = DocumentCodeDescription.Text; 
                    Fodselsnr = textBox3.Text;
                    DokAnkomstStatus = textBox4.Text;
                    Folder = textBox5.Text;

                    if (Skadenummer.Equals("") || Skadenummer.Equals(" "))
                        Skadenummer = Fodselsnr;

                    //Properties.Settings.Default["FileServerPath"] = textBox5.Text;
                    //Properties.Settings.Default.Save();
                    

                    validated = true;
                    this.Close();
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private bool FormValidation()
        {
            bool valid = true;

            try
            {
                var validation = KLP.Utils.Utils.CheckFødselsnummer(textBox3.Text);

                if (textBox1.Text.Equals(""))
                {
                    if (!validation.Item1)
                    {
                        MessageBox.Show("Skadenummer må fylles ut!");
                        valid = false;
                    }
                }

                if ((textBox1.Text.Length > 6 || !Regex.IsMatch(textBox1.Text, @"^(\+|-)?\d+(\.\d+)?$")) && textBox3.Text.Equals(""))
                {
                    MessageBox.Show("Skadenummer feil!");
                    valid = false;
                }

                if (!validation.Item1)
                {
                    if (textBox1.Text.Equals(""))
                    {
                        MessageBox.Show(validation.Item2);
                        valid = false;
                    }
                }

                if (comboBox1.SelectedIndex == -1 && comboBox1.Text == string.Empty)
                {
                    MessageBox.Show("Dokumentkode må fylles ut!");
                    valid = false;
                }

                if (DocumentCodeDescription.Text.Length > 100)
                {
                    MessageBox.Show("Dokumentbeskrivelsen er for lang");
                    valid = false;
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }

            return valid;
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    textBox5.ReadOnly = false;
        //}

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (!char.IsControl(e.KeyChar)
                    && !char.IsDigit(e.KeyChar)
                    && e.KeyChar != '.')
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if (e.KeyChar == '.'
                    && (sender as TextBox).Text.IndexOf('.') > -1)
                {
                    e.Handled = true;
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private const Keys CopyKeys = Keys.Control | Keys.C;
        private const Keys PasteKeys = Keys.Control | Keys.V;

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if ((keyData == CopyKeys) || (keyData == PasteKeys))
            {
                return true;
            }
            else
            {
                return base.ProcessCmdKey(ref msg, keyData);
            }
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            try
            {
                comboBox1.Width = comboBox1.Width + DocumentCodeDescription.Width + 10;
                DocumentCodeDescription.Visible = false;
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                comboBox1.Width = comboSize;
                DocumentCodeDescription.Visible = true;
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                comboBox1.Width = comboSize;
                DocumentCodeDescription.Visible = true;

                if (comboBox1.SelectedIndex > -1)
                {
                    var documentCodeSelected = comboBox1.GetItemText(comboBox1.SelectedItem).Split(
                        new string[] {" - "}, StringSplitOptions.None);

                    BeginInvoke(new Action(() => comboBox1.Text = documentCodeSelected[0]));
                    
                    DocumentCodeDescription.Text = documentCodeSelected[1];
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }



            
    }
}
