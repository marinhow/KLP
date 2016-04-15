using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace KLPOutlookAddIn
{
    public partial class KLPRibbon
    {
        private ErrorLogging _errorLogging;

        private void KLPRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _errorLogging = new ErrorLogging();
        }

        private void exportToPDFButton_Click(object sender, RibbonControlEventArgs e)
        {

        }

    }
}
