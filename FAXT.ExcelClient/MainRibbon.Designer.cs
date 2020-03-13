namespace FAXT.ExcelClient
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.tabMain = this.Factory.CreateRibbonTab();
            this.groupInput = this.Factory.CreateRibbonGroup();
            this.btnDropdownHelper = this.Factory.CreateRibbonButton();
            this.groupImport = this.Factory.CreateRibbonGroup();
            this.btnXML = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tabMain.SuspendLayout();
            this.groupInput.SuspendLayout();
            this.groupImport.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Groups.Add(this.groupInput);
            this.tabMain.Groups.Add(this.groupImport);
            this.tabMain.Groups.Add(this.groupHelp);
            this.tabMain.Label = "Excel Helper";
            this.tabMain.Name = "tabMain";
            // 
            // groupInput
            // 
            this.groupInput.Items.Add(this.btnDropdownHelper);
            this.groupInput.Label = "Data Input";
            this.groupInput.Name = "groupInput";
            // 
            // btnDropdownHelper
            // 
            this.btnDropdownHelper.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDropdownHelper.Label = "Dropdown Helper";
            this.btnDropdownHelper.Name = "btnDropdownHelper";
            this.btnDropdownHelper.OfficeImageId = "ChartQuickExplore";
            this.btnDropdownHelper.ShowImage = true;
            this.btnDropdownHelper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropdownHelper_Click);
            // 
            // groupImport
            // 
            this.groupImport.Items.Add(this.btnXML);
            this.groupImport.Items.Add(this.button1);
            this.groupImport.Label = "Data Import";
            this.groupImport.Name = "groupImport";
            // 
            // btnXML
            // 
            this.btnXML.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnXML.Label = "XML Importer";
            this.btnXML.Name = "btnXML";
            this.btnXML.OfficeImageId = "ImportXmlFile";
            this.btnXML.ShowImage = true;
            this.btnXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.XMLImporter_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "PDF Table Grabber";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "ContTypeApplyToList";
            this.button1.ShowImage = true;
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btnHelp);
            this.groupHelp.Label = "About";
            this.groupHelp.Name = "groupHelp";
            // 
            // btnHelp
            // 
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Label = "Help";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.OfficeImageId = "Help";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Help_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tabMain.ResumeLayout(false);
            this.tabMain.PerformLayout();
            this.groupInput.ResumeLayout(false);
            this.groupInput.PerformLayout();
            this.groupImport.ResumeLayout(false);
            this.groupImport.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupInput;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDropdownHelper;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
