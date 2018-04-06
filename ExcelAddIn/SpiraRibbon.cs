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
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Net;

namespace SpiraExcelAddIn
{
    /// <summary>
    /// Contains the SpiraTeam Importer ribbon tab
    /// </summary>
    public partial class SpiraRibbon
    {
        private ConnectForm connectForm;
        private OptionsForm optionsForm;
        private ProgressForm progressForm;
        private Importer importer = null;
        private SpiraImportExport.RemoteProject[] remoteProjects = null;

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
            httpBinding.ReceiveTimeout = new TimeSpan(0, 30, 0);    //30 minutes
            httpBinding.SendTimeout = new TimeSpan(0, 30, 0);    //30 minutes

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
                this.btnImport.Enabled = true;
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
            //Show success message
            MessageBox.Show("Import/Export of " + e.RowCount.ToString() + " rows successfully completed", "Import/Export Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            MessageBox.Show(e.Message, "Error During Import/Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        /// <summary>
        /// Called when the ribbon is first loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SpiraRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Populate the Artifact Type entries
            ddlArtifactType.ScreenTip = "Please Select Type";
            RibbonDropDownItem ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Requirements";
            ddi.Tag = "Requirements";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Releases";
            ddi.Tag = "Releases";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Test Cases";
            ddi.Tag = "Test Cases";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Test Sets";
            ddi.Tag = "Test Sets";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Test Runs";
            ddi.Tag = "Test Runs";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Incidents";
            ddi.Tag = "Incidents";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Tasks";
            ddi.Tag = "Tasks";
            ddlArtifactType.Items.Add(ddi);
            ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = "Custom Values";
            ddi.Tag = "Custom Values";
            ddlArtifactType.Items.Add(ddi);

            //By default the drop-down lists and import/export buttons are disabled
            ddlProject.ScreenTip = "Please Select Project";
            this.ddlProject.Enabled = false;
            this.ddlArtifactType.Enabled = false;
            this.btnImport.Enabled = false;
            this.btnExport.Enabled = false;
            this.btnDisconnect.Enabled = false;
            this.btnClear.Enabled = true;
            this.btnOptions.Enabled = true;

            //Need to also create the popup connection dialog box
            //and attach the connect succeeded event
            this.connectForm = new ConnectForm();
            this.connectForm.ConnectSucceeded += new EventHandler(connectForm_ConnectSucceeded);

            //Also create the options form
            this.optionsForm = new OptionsForm();
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
        /// Allow the user to import the data to SpiraTeam
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, RibbonControlEventArgs e)
        {
            //Make sure the user really wanted to do an import
            DialogResult result = MessageBox.Show("Are you sure you want to Import data from Spira. This will overwrite any information in the spreadsheet?", "Confirm Import", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
            {
                return;
            }

            //Make sure that a project and artifact type have been selected
            if (this.ddlProject.SelectedItem == null)
            {
                MessageBox.Show("You need to select a project before starting the Import.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (this.ddlArtifactType.SelectedItem == null)
            {
                MessageBox.Show("You need to select an artifact before starting the Import.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            //Display the Import progress bar
            this.progressForm = new ProgressForm();
            this.progressForm.Caption = "Importing data from Spira";
            this.progressForm.Show();
            this.progressForm.FormClosed += new FormClosedEventHandler(progressForm_FormClosed);

            try
            {
                //Get the artifact type and project id from the selection
                string artifactTypeName = this.ddlArtifactType.SelectedItem.Label;
                string projectName = this.ddlProject.SelectedItem.Label;
                if (this.remoteProjects == null)
                {
                    throw new ApplicationException("No projects have been loaded, unable to import");
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
                importer = new Importer();
                importer.ExcelApplication = Globals.ThisAddIn.Application;
                importer.ProgressForm = this.progressForm;
                importer.ErrorOccurred += new Importer.ImportExportErrorHandler(importer_ErrorOccurred);
                importer.OperationCompleted += new Importer.ImportExportCompletedHandler(importer_OperationCompleted);
                importer.Import(projectId, artifactTypeName);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Error During Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.progressForm.Close();
            }
        }

        /// <summary>
        /// Allow the user to export the data into SpiraTeam
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, RibbonControlEventArgs e)
        {
            //Make sure the user really wanted to do an export
            DialogResult result = MessageBox.Show("Are you sure you want to Export data to Spira. This will insert and update rows in the live system?", "Confirm Export", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                importer = new Importer();
                importer.ExcelApplication = Globals.ThisAddIn.Application;
                importer.ProgressForm = this.progressForm;
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
        /// Clear the active worksheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClear_Click(object sender, RibbonControlEventArgs e)
        {
            //Clear the active worksheet apart from the first two rows that contain header information
            DialogResult result = MessageBox.Show("Are you sure you want to Clear the active worksheet. This will remove any data in the current sheet?", "Confirm Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
            {
                return;
            }
            Microsoft.Office.Interop.Excel.Worksheet workSheet = Globals.ThisAddIn.Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            if (workSheet == null)
            {
                MessageBox.Show("There is no active worksheet. Please open up the Import template and try again.", "No Active Worksheet", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                Range usedRange = workSheet.UsedRange;
                if (usedRange != null && usedRange.Rows.Count > 2)
                {
                    Range rangeToClear = usedRange.Range[usedRange[3, 1], usedRange[usedRange.Rows.Count, usedRange.Columns.Count]];
                    rangeToClear.ClearContents();
                }
            }
        }

        /// <summary>
        /// Called when the options button is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOptions_Click(object sender, RibbonControlEventArgs e)
        {
            //Display the options dialog form
            this.optionsForm.ShowDialog();
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
            this.btnImport.Enabled = false;
        }

        /// <summary>
        /// Called when the project dropdown is changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ddlProject_SelectionChanged(object sender, RibbonControlEventArgs e)
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

                //Make sure that a project has been selected
                if (this.ddlProject.SelectedItem == null)
                {
                    return;
                }

                string projectName = this.ddlProject.SelectedItem.Label;
                if (this.remoteProjects == null)
                {
                    return;
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

                //Now connect to the project
                success = spiraImportExport.Connection_ConnectToProject(projectId);
                if (!success)
                {
                    //Authentication failed
                    throw new ApplicationException("Unable to connect to project PR" + projectId + ". Please check that the user is a member of the project.");
                }

                //Make sure we have a Lookups worksheet available
                Worksheet lookupWorksheet = null;
                foreach (Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    if (worksheet.Name.Trim().ToLowerInvariant() == "lookups")
                    {
                        lookupWorksheet = worksheet;
                        break;
                    }
                }
                if (lookupWorksheet == null)
                {
                    return;
                }

                //Get the Incident project-specific dynamic lookup ranges for types, statuses, priorities and severities
                Range lookupCell = lookupWorksheet.Range["Inc_Lookups"];
                int startColumn = lookupCell.Column;
                int startRow = lookupCell.Row + 2;  //Two rows for the header

                //Incident Types
                SpiraImportExport.RemoteIncidentType[] incidentTypes = spiraImportExport.Incident_RetrieveTypes();
                Range dataRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn], lookupWorksheet.Cells[startRow + incidentTypes.Length, startColumn + 1]];
                object[,] dataValues = (object[,])dataRange.Value2;
                int row = 1;
                foreach (SpiraImportExport.RemoteIncidentType remoteType in incidentTypes)
                {
                    dataValues[row, 1] = remoteType.Name;
                    dataValues[row, 2] = remoteType.IncidentTypeId;
                    row++;
                }
                dataRange.Value2 = dataValues;
                Range validationRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn], lookupWorksheet.Cells[startRow + incidentTypes.Length - 1, startColumn]];
                validationRange.Name = "Inc_Type";

                //Incident Statuses
                SpiraImportExport.RemoteIncidentStatus[] incidentStatuses = spiraImportExport.Incident_RetrieveStatuses();
                dataRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn + 2], lookupWorksheet.Cells[startRow + incidentStatuses.Length, startColumn + 3]];
                dataValues = (object[,])dataRange.Value2;
                row = 1;
                foreach (SpiraImportExport.RemoteIncidentStatus remoteStatus in incidentStatuses)
                {
                    dataValues[row, 1] = remoteStatus.Name;
                    dataValues[row, 2] = remoteStatus.IncidentStatusId;
                    row++;
                }
                dataRange.Value2 = dataValues;
                validationRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn + 2], lookupWorksheet.Cells[startRow + incidentStatuses.Length - 1, startColumn + 2]];
                validationRange.Name = "Inc_Status";

                //Incident Priorities
                SpiraImportExport.RemoteIncidentPriority[] incidentPriorities = spiraImportExport.Incident_RetrievePriorities();
                dataRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn + 4], lookupWorksheet.Cells[startRow + incidentPriorities.Length, startColumn + 5]];
                dataValues = (object[,])dataRange.Value2;
                row = 1;
                foreach (SpiraImportExport.RemoteIncidentPriority remotePriority in incidentPriorities)
                {
                    dataValues[row, 1] = remotePriority.Name;
                    dataValues[row, 2] = remotePriority.PriorityId;
                    row++;
                }
                dataRange.Value2 = dataValues;
                validationRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn + 4], lookupWorksheet.Cells[startRow + incidentPriorities.Length - 1, startColumn + 4]];
                validationRange.Name = "Inc_Priority";

                //Incident Severities
                SpiraImportExport.RemoteIncidentSeverity[] incidentSeverities = spiraImportExport.Incident_RetrieveSeverities();
                dataRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn + 6], lookupWorksheet.Cells[startRow + incidentSeverities.Length, startColumn + 7]];
                dataValues = (object[,])dataRange.Value2;
                row = 1;
                foreach (SpiraImportExport.RemoteIncidentSeverity remoteSeverity in incidentSeverities)
                {
                    dataValues[row, 1] = remoteSeverity.Name;
                    dataValues[row, 2] = remoteSeverity.SeverityId;
                    row++;
                }
                dataRange.Value2 = dataValues;
                validationRange = (Range)lookupWorksheet.Range[lookupWorksheet.Cells[startRow, startColumn + 6], lookupWorksheet.Cells[startRow + incidentSeverities.Length - 1, startColumn + 6]];
                validationRange.Name = "Inc_Severity";
            }
            catch (TimeoutException exception)
            {
                // Handle the timeout exception.
                MessageBox.Show("A timeout error occurred! (" + exception.Message + ")", "Timeout Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (CommunicationException exception)
            {
                // Handle the communication exception.
                MessageBox.Show("A communication error occurred! (" + exception.Message + ")", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
