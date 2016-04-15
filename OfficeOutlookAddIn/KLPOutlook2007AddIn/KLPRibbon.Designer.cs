namespace KLPOutlookAddIn
{
    partial class KLPRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public KLPRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(KLPRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.exportToPDFButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.KeyTip = "K";
            this.tab1.Label = "KLP Verktøy";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.exportToPDFButton);
            this.group1.Label = "Export";
            this.group1.Name = "group1";
            // 
            // exportToPDFButton
            // 
            this.exportToPDFButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.exportToPDFButton.Image = ((System.Drawing.Image)(resources.GetObject("exportToPDFButton.Image")));
            this.exportToPDFButton.KeyTip = "I";
            this.exportToPDFButton.Label = "PDF";
            this.exportToPDFButton.Name = "exportToPDFButton";
            this.exportToPDFButton.ShowImage = true;
            this.exportToPDFButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportToPDFButton_Click);
            // 
            // KLPRibbon
            // 
            this.Name = "KLPRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.KLPRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportToPDFButton;
    }

    partial class ThisRibbonCollection
    {
        internal KLPRibbon KLPRibbon
        {
            get { return this.GetRibbon<KLPRibbon>(); }
        }
    }
}
