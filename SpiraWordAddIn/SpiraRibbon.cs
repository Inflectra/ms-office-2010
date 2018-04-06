using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;

using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Net;

namespace SpiraWordAddIn
{
    /// <summary>
    /// Contains the SpiraTeam Importer ribbon tab
    /// </summary>
    public partial class SpiraRibbon
    {
        private ConnectForm connectForm;
        private ProgressForm progressForm;
        private ParametersForm parametersForm;
        private Importer importer = null;
        private SpiraImportExport.RemoteProject[] remoteProjects = null;

        private Dictionary<MappedStyleKeys, string> mappedStyles = new Dictionary<MappedStyleKeys, string>();

        #region Enumerations

        public enum MappedStyleKeys
        {
            Requirement_Indent1,
            Requirement_Indent2,
            Requirement_Indent3,
            Requirement_Indent4,
            Requirement_Indent5,
            TestCase_Folder,
            TestCase_TestCase,
            TestStep_Description,
            TestStep_ExpectedResult,
            TestStep_SampleData
        }

        #endregion

        #region Properties

        /// <summary>
        /// Indicates if we have projects populated or not
        /// </summary>
        public bool ProjectsPopulated
        {
            get
            {
                if (this.remoteProjects == null || this.remoteProjects.Length == 0)
                {
                    return false;
                }
                return true;
            }
        }

        /// <summary>
        /// Used to store the password (since only stored in settings if RememberMe checked)
        /// </summary>
        public static string SpiraPassword
        {
            get;
            set;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Creates the WCF endpoints
        /// </summary>
        /// <param name="fullUri">The URI</param>
        /// <returns>The client class</returns>
        /// <remarks>We need to do this in code because the app.config file is not available in VSTO</remarks>
        public static SpiraImportExport.SoapServiceClient CreateClient(Uri fullUri)
        {
            //Configure the binding
            BasicHttpBinding httpBinding = new BasicHttpBinding();

            //Allow cookies and large messages
            httpBinding.AllowCookies = true;
            httpBinding.MaxBufferSize = 100000000; //100MB
            httpBinding.MaxReceivedMessageSize = 100000000; //100MB
            httpBinding.ReaderQuotas.MaxStringContentLength = 2147483647;
            httpBinding.ReaderQuotas.MaxDepth = 2147483647;
            httpBinding.ReaderQuotas.MaxBytesPerRead = 2147483647;
            httpBinding.ReaderQuotas.MaxNameTableCharCount = 2147483647;
            httpBinding.ReaderQuotas.MaxArrayLength = 2147483647;

            //Handle SSL if necessary
            if (fullUri.Scheme == "https")
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.Transport;
                httpBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;

                //Allow self-signed certificates
                PermissiveCertificatePolicy.Enact("");

                //Force TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            }
            else
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.None;
            }

            //Create the new client with endpoint and HTTP Binding
            EndpointAddress endpointAddress = new EndpointAddress(fullUri.AbsoluteUri);
            SpiraImportExport.SoapServiceClient spiraImportExport = new SpiraImportExport.SoapServiceClient(httpBinding, endpointAddress);

            //Modify the operation behaviors to allow unlimited objects in the graph
            foreach (var operation in spiraImportExport.Endpoint.Contract.Operations)
            {
                var behavior = operation.Behaviors.Find<DataContractSerializerOperationBehavior>() as DataContractSerializerOperationBehavior;
                if (behavior != null)
                {
                    behavior.MaxItemsInObjectGraph = 2147483647;
                }
            }

            return spiraImportExport;
        }

        /// <summary>
        /// Loads the list of projects and enables the rest of the toolbar
        /// </summary>
        private void LoadProjects()
        {
            //Create the full URL from the provided base URL
            Uri fullUri;
            if (!Importer.TryCreateFullUrl(Configuration.Default.SpiraUrl, out fullUri))
            {
                MessageBox.Show("The Server URL entered is not a valid URL", "Connect to Server", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            SpiraImportExport.SoapServiceClient spiraImportExport = CreateClient(fullUri);
            try
            {
                //Authenticate 
                bool success = spiraImportExport.Connection_Authenticate(Configuration.Default.SpiraUserName, SpiraPassword);
                if (!success)
                {
                    //Authentication failed
                    MessageBox.Show("Unable to authenticate with Server. Please check the username and password and try again.", "Load Projects", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //Now get the list of projects
                this.remoteProjects = spiraImportExport.Project_Retrieve();
                if (!this.ProjectsPopulated)
                {
                    MessageBox.Show("No projects were returned. Please make sure that you are the member of at least one project and try again.", "Load Projects", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //Make the drop-downs enabled
                this.ddlProject.Enabled = true;
                this.ddlArtifactType.Enabled = true;

                //Now bind the list of projects to the dropdown
                this.ddlProject.Items.Clear();
                ddlProject.ScreenTip = "Please Select Project";
                foreach (SpiraImportExport.RemoteProject remoteProject in this.remoteProjects)
                {
                    RibbonDropDownItem ddi = Factory.CreateRibbonDropDownItem();
                    ddi.Label = remoteProject.Name;
                    this.ddlProject.Items.Add(ddi);
                }

                //Now make the other buttons enabled
                this.btnConnect.Enabled = false;
                this.btnDisconnect.Enabled = true;
                this.btnExport.Enabled = true;
            }
            catch (TimeoutException exception)
            {
                // Handle the timeout exception.
                spiraImportExport.Abort();
                MessageBox.Show("A timeout error occurred! (" + exception.Message + ")", "Timeout Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (CommunicationException exception)
            {
                // Handle the communication exception.
                spiraImportExport.Abort();
                MessageBox.Show("A communication error occurred! (" + exception.Message + ")", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// Cancels the current operation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void progressForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (this.importer != null)
            {
                this.importer.AbortOperation();
            }
        }

        /// <summary>
        /// Called when the import has completed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void importer_OperationCompleted(object sender, ImportExportCompletedArgs e)
        {
            this.btnExport.Enabled = true;
            //Show success message
            MessageBox.Show("Export of " + e.RowCount.ToString() + " items successfully completed", "Export Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// This handler is called when the connection succeeds
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void connectForm_ConnectSucceeded(object sender, EventArgs e)
        {
            //Load the project list
            LoadProjects();
        }

        /// <summary>
        /// Called when an error occurs
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void importer_ErrorOccurred(object sender, ImportExportErrorArgs e)
        {
            this.btnExport.Enabled = true;
            MessageBox.Show(e.Message, "Error During Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        /// <summary>
        /// Called when the ribbon is first loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SpiraRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Populate the initial set of style mappings
            this.mappedStyles.Add(MappedStyleKeys.Requirement_Indent1, "Heading 1");
            this.mappedStyles.Add(MappedStyleKeys.Requirement_Indent2, "Heading 2");
            this.mappedStyles.Add(MappedStyleKeys.Requirement_Indent3, "Heading 3");
            this.mappedStyles.Add(MappedStyleKeys.Requirement_Indent4, "Heading 4");
            this.mappedStyles.Add(MappedStyleKeys.Requirement_Indent5, "Heading 4");
            this.mappedStyles.Add(MappedStyleKeys.TestCase_Folder, "Heading 1");
            this.mappedStyles.Add(MappedStyleKeys.TestCase_TestCase, "Heading 2");
            this.mappedStyles.Add(MappedStyleKeys.TestStep_Description, "Column 1");
            this.mappedStyles.Add(MappedStyleKeys.TestStep_ExpectedResult, "Column 2");
            this.mappedStyles.Add(MappedStyleKeys.TestStep_SampleData, "Column 3");

            //Populate the Artifact Type entries
            ddlArtifactType.ScreenTip = "Please Select Type";
            RibbonDropDownItem ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Requirements";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Test Cases";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();

            //By default the drop-down lists and import/export buttons are disabled
            ddlProject.ScreenTip = "Please Select Project";
            this.ddlProject.Enabled = false;
            this.ddlArtifactType.Enabled = false;
            this.btnExport.Enabled = false;
            this.btnDisconnect.Enabled = false;
            this.btnParameters.Enabled = true;
            this.btnParameters.Enabled = (Globals.ThisAddIn.Application.Documents.Count > 0);

            //Need to also create the popup connection dialog box
            //and attach the connect succeeded event
            this.connectForm = new ConnectForm();
            this.connectForm.ConnectSucceeded += new EventHandler(connectForm_ConnectSucceeded);

            //Need to also create the parameters dialog box
            this.parametersForm = new ParametersForm();

            //Trap the window active event so that we can update combo-box references
            //As they change whenever we open a new document
            Globals.ThisAddIn.Application.WindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate);
            Globals.ThisAddIn.Application.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            Globals.ThisAddIn.Application.DocumentBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
        }

        /// <summary>
        /// Display the connection dialog box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConnect_Click(object sender, RibbonControlEventArgs e)
        {
            //Display the connection dialog form
            this.connectForm.ShowDialog();
        }


        /// <summary>
        /// Allow the user to export the data into SpiraTeam
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, RibbonControlEventArgs e)
        {
            //Make sure the user really wanted to do an export
            DialogResult result = MessageBox.Show("Are you sure you want to Export data to Spira. This will insert new rows in the live system?", "Confirm Export", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
            {
                return;
            }

            //Make sure that a project and artifact type have been selected
            if (this.ddlProject.SelectedItem == null)
            {
                MessageBox.Show("You need to select a project before starting the Export.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (this.ddlArtifactType.SelectedItem == null)
            {
                MessageBox.Show("You need to select an artifact before starting the Export.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            //Display the Export progress bar
            this.progressForm = new ProgressForm();
            this.progressForm.Caption = "Exporting data to Spira";
            this.progressForm.Show();
            this.progressForm.FormClosed += new FormClosedEventHandler(progressForm_FormClosed);

            try
            {
                //Get the artifact type and project id from the selection
                string artifactTypeName = this.ddlArtifactType.SelectedItem.Label;
                string projectName = this.ddlProject.SelectedItem.Label;
                if (this.remoteProjects == null)
                {
                    throw new ApplicationException("No projects have been loaded, unable to export");
                }
                int projectId = -1;
                foreach (SpiraImportExport.RemoteProject remoteProject in this.remoteProjects)
                {
                    if (remoteProject.Name == projectName)
                    {
                        projectId = remoteProject.ProjectId.Value;
                    }
                }
                if (projectId == -1)
                {
                    throw new ApplicationException("Unable to find matching project name, please try disconnecting and reconnecting.");
                }

                //Start the import process and attach the event handlers in case of error or completed
                this.btnExport.Enabled = false; //Prevent multiple export attempts
                importer = new Importer();
                importer.WordApplication = Globals.ThisAddIn.Application;
                importer.ProgressForm = this.progressForm;
                importer.MappedStyles = mappedStyles;
                importer.ErrorOccurred += new Importer.ImportExportErrorHandler(importer_ErrorOccurred);
                importer.OperationCompleted += new Importer.ImportExportCompletedHandler(importer_OperationCompleted);
                importer.Export(projectId, artifactTypeName);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Error During Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.progressForm.Close();
            }
        }

        /// <summary>
        /// Make the ribbon disabled until we reconnect
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDisconnect_Click(object sender, RibbonControlEventArgs e)
        {
            //Grey out the various buttons
            this.btnConnect.Enabled = true;
            this.btnDisconnect.Enabled = false;
            this.btnExport.Enabled = false;
        }

        /// <summary>
        /// Called BEFORE a document closes
        /// </summary>
        /// <param name="Doc"></param>
        void Application_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            //Hide the parameters button if we don't have an active document after this closes
            if (Globals.ThisAddIn.Application.Documents.Count <= 1)
            {
                this.btnParameters.Enabled = false;
            }
        }

        /// <summary>
        /// Called when a document is opened
        /// </summary>
        /// <param name="Doc"></param>
        void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            //Make the parameters button enabled
            this.btnParameters.Enabled = true;
        }

        /// <summary>
        /// Enables the parameters icon when a new document is loaded
        /// </summary>
        /// <param name="Doc"></param>
        /// <param name="Wn"></param>
        void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            //Make the parameters button enabled
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                this.btnParameters.Enabled = true;
            }
        }

        /// <summary>
        /// Called when the style mappings icon is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnParameters_Click(object sender, RibbonControlEventArgs e)
        {
            //Make sure we have a document loaded
            try
            {
                if (Globals.ThisAddIn.Application.ActiveDocument == null || Globals.ThisAddIn.Application.Documents == null || Globals.ThisAddIn.Application.Documents.Count == 0)
                {
                    MessageBox.Show("No Word document is currently loaded. Please open the Word document you wish to export from.", "Document Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No Word document is currently loaded. Please open the Word document you wish to export from.", "Document Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            //Display the "Update Style Mapping" dialog box
            this.parametersForm.WordDocument = Globals.ThisAddIn.Application.ActiveDocument;
            this.parametersForm.MappedStyles = this.mappedStyles;
            this.parametersForm.ShowDialog();
        }

        #endregion
    }
}
