using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace KLPOutlookAddIn
{
    public partial class Form1 : Form
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

        private int comboSize;

        private ErrorLogging _errorLogging;

        private List<string> validDocumentCodes;

        public Form1(string skadenummer)
        {
            try
            {
                InitializeComponent();
                _errorLogging = new ErrorLogging();
                textBox4.Text = "WF";
                comboSize = comboBox1.Width;
                textBox1.Text = skadenummer;
                textBox5.Text = KLP.Utils.Utils.GetActiveEnvironment(KLP.Utils.Utils.CONFIGURATIONXML);
                AddProductCodes(KLP.Utils.Utils.DOCUMENTCODESXML);

                validated = false;
            }
            catch (Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }


        private void AddProductCodes(string documentCodesFileLocation)
        {

            validDocumentCodes = new List<string>();

            //List<Tuple<string, string>> productCodes =
            //    new Helper().Read(Properties.Settings.Default.DocumentCodesFileLocation);

            List<Tuple<string, string>> productCodes = KLP.Utils.Utils.Read(documentCodesFileLocation);            


            foreach (var productCode in productCodes)
            {
                comboBox1.Items.Add(productCode.Item1 + " - " + productCode.Item2);
                validDocumentCodes.Add(productCode.Item1);
            }

            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
            
        }

        private void button1_Click(object sender, EventArgs e)
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

                //Properties.Settings.Default["FileServerPath"] = textBox5.Text;
                //Properties.Settings.Default.Save();
                Folder = textBox5.Text;

                if (Skadenummer.Equals("") || Skadenummer.Equals(" "))
                    Skadenummer = Fodselsnr;

                validated = true;
                this.Close();

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

                //if ((textBox1.Text.Length > 6 || !Regex.IsMatch(textBox1.Text, @"^(\+|-)?\d+(\.\d+)?$")) && textBox3.Text.Equals(""))
                //{
                //    MessageBox.Show("Skadenummer feil!");
                //    valid = false;
                //}

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

                //Document Codes Validation
                if (!validDocumentCodes.Contains(comboBox1.Text))
                {
                    MessageBox.Show("Dokumentkode feil!");
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


        //private void button2_Click(object sender, EventArgs e)
        //{
        //    textBox5.ReadOnly = false;
        //}

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
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

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            changeDocumentValues();
        }

        private void changeDocumentValues()
        {
            try
            {
                comboBox1.Width = comboSize;
                DocumentCodeDescription.Visible = true;

                if (comboBox1.SelectedIndex > -1)
                {
                    var documentCodeSelected = comboBox1.GetItemText(comboBox1.SelectedItem).Split(
                        new string[] { " - " }, StringSplitOptions.None);

                    BeginInvoke(new Action(() => comboBox1.Text = documentCodeSelected[0]));

                    DocumentCodeDescription.Text = documentCodeSelected[1];
                }
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            } 
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            try
            {
                //changeDocumentValues();

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
                changeDocumentValues();

                comboBox1.Width = comboSize;
                DocumentCodeDescription.Visible = true;
            }
            catch (System.Exception exception)
            {
                _errorLogging.WriteLog(exception);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                textBox3.Text = string.Empty;
                textBox3.Enabled = false;
            }
            else
            {
                textBox3.Enabled = true;
            }

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text.Length > 0)
            {
                textBox1.Text = string.Empty;
                textBox1.Enabled = false;
            }
            else
            {
                textBox1.Enabled = true;
            }
        }

    }
}
