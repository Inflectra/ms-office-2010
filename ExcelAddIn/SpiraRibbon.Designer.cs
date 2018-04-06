namespace SpiraExcelAddIn
{
    partial class SpiraRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SpiraRibbon()
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
            this.tabInflectra = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnConnect = this.Factory.CreateRibbonButton();
            this.btnDisconnect = this.Factory.CreateRibbonButton();
            this.btnOptions = this.Factory.CreateRibbonButton();
            this.ddlProject = this.Factory.CreateRibbonDropDown();
            this.ddlArtifactType = this.Factory.CreateRibbonDropDown();
            this.box2 = this.Factory.CreateRibbonBox();
            this.btnImport = this.Factory.CreateRibbonButton();
            this.btnExport = this.Factory.CreateRibbonButton();
            this.btnClear = this.Factory.CreateRibbonButton();
            this.tabInflectra.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // tabInflectra
            // 
            this.tabInflectra.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabInflectra.Groups.Add(this.group1);
            this.tabInflectra.Label = "TabAddIns";
            this.tabInflectra.Name = "tabInflectra";
            // 
            // group1
            // 
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.ddlProject);
            this.group1.Items.Add(this.ddlArtifactType);
            this.group1.Items.Add(this.box2);
            this.group1.Label = "SpiraTeam";
            this.group1.Name = "group1";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.btnConnect);
            this.box1.Items.Add(this.btnDisconnect);
            this.box1.Items.Add(this.btnOptions);
            this.box1.Name = "box1";
            // 
            // btnConnect
            // 
            this.btnConnect.Image = global::SpiraExcelAddIn.Properties.Resources.SpiraIcon1;
            this.btnConnect.Label = "Connect";
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.ShowImage = true;
            this.btnConnect.SuperTip = "Connect to Spira";
            this.btnConnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConnect_Click);
            // 
            // btnDisconnect
            // 
            this.btnDisconnect.Image = global::SpiraExcelAddIn.Properties.Resources.DisconnectIcon;
            this.btnDisconnect.Label = "Disconnect";
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.ShowImage = true;
            this.btnDisconnect.SuperTip = "Disconnect from Spira";
            this.btnDisconnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisconnect_Click);
            // 
            // btnOptions
            // 
            this.btnOptions.Image = global::SpiraExcelAddIn.Properties.Resources.OptionsIcon;
            this.btnOptions.Label = "Options";
            this.btnOptions.Name = "btnOptions";
            this.btnOptions.ShowImage = true;
            this.btnOptions.SuperTip = "Change Import/Export Options";
            this.btnOptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOptions_Click);
            // 
            // ddlProject
            // 
            this.ddlProject.Label = "Project:";
            this.ddlProject.Name = "ddlProject";
            this.ddlProject.SizeString = "xxxxxxxxxxxxxxxxxxxx";
            this.ddlProject.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddlProject_SelectionChanged);
            // 
            // ddlArtifactType
            // 
            this.ddlArtifactType.Label = "Artifact:";
            this.ddlArtifactType.Name = "ddlArtifactType";
            this.ddlArtifactType.SizeString = "xxxxxxxxxxxxxxxxxxxx";
            // 
            // box2
            // 
            this.box2.Items.Add(this.btnImport);
            this.box2.Items.Add(this.btnExport);
            this.box2.Items.Add(this.btnClear);
            this.box2.Name = "box2";
            // 
            // btnImport
            // 
            this.btnImport.Image = global::SpiraExcelAddIn.Properties.Resources.ImportIcon1;
            this.btnImport.Label = "Import";
            this.btnImport.Name = "btnImport";
            this.btnImport.ShowImage = true;
            this.btnImport.SuperTip = "Import from Spira to Excel";
            this.btnImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImport_Click);
            // 
            // btnExport
            // 
            this.btnExport.Image = global::SpiraExcelAddIn.Properties.Resources.ExportIcon1;
            this.btnExport.Label = "Export";
            this.btnExport.Name = "btnExport";
            this.btnExport.ShowImage = true;
            this.btnExport.SuperTip = "Export to Spira from Excel";
            this.btnExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExport_Click);
            // 
            // btnClear
            // 
            this.btnClear.Image = global::SpiraExcelAddIn.Properties.Resources.ClearIcon1;
            this.btnClear.Label = "Clear";
            this.btnClear.Name = "btnClear";
            this.btnClear.ShowImage = true;
            this.btnClear.SuperTip = "Clear the current Excel worksheet";
            this.btnClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClear_Click);
            // 
            // SpiraRibbon
            // 
            this.Name = "SpiraRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabInflectra);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SpiraRibbon_Load);
            this.tabInflectra.ResumeLayout(false);
            this.tabInflectra.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabInflectra;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisconnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlArtifactType;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
    }

    partial class ThisRibbonCollection
    {
        internal SpiraRibbon SpiraRibbon
        {
            get { return this.GetRibbon<SpiraRibbon>(); }
        }
    }
}
