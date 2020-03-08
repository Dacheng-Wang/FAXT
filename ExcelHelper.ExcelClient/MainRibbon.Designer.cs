namespace ExcelHelper.ExcelClient
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
            this.groupDropdown = this.Factory.CreateRibbonGroup();
            this.btnDropdownHelper = this.Factory.CreateRibbonButton();
            this.tabMain.SuspendLayout();
            this.groupDropdown.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Groups.Add(this.groupDropdown);
            this.tabMain.Label = "Excel Helper";
            this.tabMain.Name = "tabMain";
            // 
            // groupDropdown
            // 
            this.groupDropdown.Items.Add(this.btnDropdownHelper);
            this.groupDropdown.Label = "Dropdown Helper";
            this.groupDropdown.Name = "groupDropdown";
            // 
            // btnDropdownHelper
            // 
            this.btnDropdownHelper.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDropdownHelper.Label = "Dropdown Helper";
            this.btnDropdownHelper.Name = "btnDropdownHelper";
            this.btnDropdownHelper.OfficeImageId = "ChartQuickExplore";
            this.btnDropdownHelper.ShowImage = true;
            this.btnDropdownHelper.Click += btnDropdownHelper_Click;
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tabMain.ResumeLayout(false);
            this.tabMain.PerformLayout();
            this.groupDropdown.ResumeLayout(false);
            this.groupDropdown.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDropdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDropdownHelper;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
