namespace SpiraWordAddIn
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnConnect = this.Factory.CreateRibbonButton();
            this.btnDisconnect = this.Factory.CreateRibbonButton();
            this.ddlProject = this.Factory.CreateRibbonDropDown();
            this.ddlArtifactType = this.Factory.CreateRibbonDropDown();
            this.box2 = this.Factory.CreateRibbonBox();
            this.btnParameters = this.Factory.CreateRibbonButton();
            this.btnExport = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
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
            this.box1.Name = "box1";
            // 
            // btnConnect
            // 
            this.btnConnect.Image = global::SpiraWordAddIn.Properties.Resources.SpiraIcon1;
            this.btnConnect.Label = "Connect";
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.ShowImage = true;
            this.btnConnect.SuperTip = "Connect to Spira";
            this.btnConnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConnect_Click);
            // 
            // btnDisconnect
            // 
            this.btnDisconnect.Image = global::SpiraWordAddIn.Properties.Resources.DisconnectIcon;
            this.btnDisconnect.Label = "Disconnect";
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.ShowImage = true;
            this.btnDisconnect.SuperTip = "Disconnect from Spira";
            this.btnDisconnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisconnect_Click);
            // 
            // ddlProject
            // 
            this.ddlProject.Label = "Project:";
            this.ddlProject.Name = "ddlProject";
            this.ddlProject.SizeString = "xxxxxxxxxxxxxxxxxxxx";
            // 
            // ddlArtifactType
            // 
            this.ddlArtifactType.Label = "Artifact:";
            this.ddlArtifactType.Name = "ddlArtifactType";
            this.ddlArtifactType.SizeString = "xxxxxxxxxxxxxxxxxxxx";
            // 
            // box2
            // 
            this.box2.Items.Add(this.btnParameters);
            this.box2.Items.Add(this.btnExport);
            this.box2.Name = "box2";
            // 
            // btnParameters
            // 
            this.btnParameters.Image = global::SpiraWordAddIn.Properties.Resources.ParametersIcon1;
            this.btnParameters.Label = "Style Mappings";
            this.btnParameters.Name = "btnParameters";
            this.btnParameters.ShowImage = true;
            this.btnParameters.SuperTip = "Edit the style mappings used in the current document";
            this.btnParameters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnParameters_Click);
            // 
            // btnExport
            // 
            this.btnExport.Image = global::SpiraWordAddIn.Properties.Resources.ExportIcon1;
            this.btnExport.Label = "Export";
            this.btnExport.Name = "btnExport";
            this.btnExport.ShowImage = true;
            this.btnExport.SuperTip = "Export to Spira from Excel";
            this.btnExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExport_Click);
            // 
            // SpiraRibbon
            // 
            this.Name = "SpiraRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SpiraRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisconnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlArtifactType;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnParameters;
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
