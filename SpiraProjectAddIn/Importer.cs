using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Xml;
using System.IO;

using Microsoft.Office.Interop.MSProject;

namespace SpiraProjectAddIn
{
    /// <summary>
    /// Contains the logic to import/export data from SpiraTeam to/from the MS Project project
    /// </summary>
    public class Importer
    {
        public const string SOAP_RELATIVE_URL = "Services/v5_0/SoapService.svc";

        private const int COM_TRUE = -1;
        private const int COM_FALSE = 0;

        private const string TEXT1_PREFIX = "Spira-";
        private const string FORMAT_DATE_TIME_INVARIANT = "{0:yyyy-MM-ddTHH:mm:ss.fff}";
        private const string FORMAT_DATE_TIME_INVARIANT_PARSE = "yyyy-MM-ddTHH:mm:ss.fff";
        private const string TEXT1_IGNORE_TOKEN = "IGNORE";

        //This event is called when an error occurs
        public delegate void ImportExportErrorHandler(object sender, ImportExportErrorArgs e);
        public event ImportExportErrorHandler ErrorOccurred;

        //This event is called when the import/export is completed
        public delegate void ImportExportCompletedHandler(object sender, ImportExportCompletedArgs e);
        public event ImportExportCompletedHandler OperationCompleted;

        //This private delegate callback is used to update the progress form in a thread-safe manner
        private delegate void UpdateProgressCallback(Nullable<int> currentValue, Nullable<int> maximumValue);
        private delegate void CloseProgressCallback();

        //Used to handle missing parameters to Interop code
        object missing = System.Reflection.Missing.Value;

        #region Properties

        /// <summary>
        /// Has the operation been aborted
        /// </summary>
        public bool IsAborted
        {
            get
            {
                return this.isAborted;
            }
        }
        protected bool isAborted = false;

        /// <summary>
        /// The handle to the MS-Word application instance
        /// </summary>
        public Microsoft.Office.Interop.MSProject.Application ProjectApplication
        {
            get;
            set;
        }

        /// <summary>
        /// The handle to the progress form
        /// </summary>
        public ProgressForm ProgressForm
        {
            get;
            set;
        }

        #endregion

        /// <summary>
        /// Called to raise the ErrorOccurred event
        /// </summary>
        /// <param name="message">The error message to display</param>
        public void OnErrorOccurred(string message)
        {
            //Close Progress
            CloseProgress();

            if (ErrorOccurred != null)
            {
                ErrorOccurred(this, new ImportExportErrorArgs(message));
            }
        }

        /// <summary>
        /// Called to raise the OperationCompleted event
        /// </summary>
        /// <param name="numberRows">The number of rows imported/exported</param>
        public void OnOperationCompleted(int numberRows)
        {
            //Close Progress
            CloseProgress();

            if (OperationCompleted != null)
            {
                OperationCompleted(this, new ImportExportCompletedArgs(numberRows));
            }
        }

        /// <summary>
        /// Closes the progress window in a thread-safe manner
        /// </summary>
        protected void CloseProgress()
        {
            if (this.ProgressForm == null)
            {
                return;
            }
            if (this.ProgressForm.InvokeRequired)
            {
                CloseProgressCallback closeProgressCallback = new CloseProgressCallback(CloseProgress);
                this.ProgressForm.Invoke(closeProgressCallback);
            }
            else
            {
                this.ProgressForm.Close();
            }
        }

        /// <summary>
        /// Updates the progress display of the form
        /// </summary>
        /// <param name="currentValue">The current value</param>
        /// <param name="maximumValue">The maximum value</param>
        protected void UpdateProgress(Nullable<int> currentValue, Nullable<int> maximumValue)
        {
            if (this.ProgressForm == null)
            {
                return;
            }

            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.ProgressForm.InvokeRequired)
            {
                UpdateProgressCallback updateProgressDelegate = new UpdateProgressCallback(UpdateProgress);
                this.ProgressForm.Invoke(updateProgressDelegate, new object[] { currentValue, maximumValue });
            }
            else
            {
                if (currentValue.HasValue)
                {
                    this.ProgressForm.ProgressValue = currentValue.Value;
                }
                if (maximumValue.HasValue)
                {
                    this.ProgressForm.ProgressMaximumValue = maximumValue.Value;
                }
            }
        }

        /// <summary>
        /// Creates the full web service URL from the provided base URL
        /// </summary>
        /// <param name="baseUrl">The provided base url</param>
        /// 
        /// <returns>The full URL to connect to</returns>
        public static bool TryCreateFullUrl(string baseUrl, out Uri fullUri)
        {
            //Add the suffix onto the base URL (need to make sure that it has the trailing slash)
            if (baseUrl[baseUrl.Length - 1] != '/')
            {
                baseUrl += "/";
            }
            Uri uri = new Uri(baseUrl, UriKind.Absolute);
            if (!Uri.TryCreate(uri, new Uri(Importer.SOAP_RELATIVE_URL, UriKind.Relative), out fullUri))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Aborts the current operation
        /// </summary>
        public void AbortOperation()
        {
            //Set the abort flag, which the background thread will see
            this.isAborted = true;
        }

        /// <summary>
        /// Imports data from Spira to the current MS-Project project
        /// </summary>
        /// <param name="projectId">The id of the project to import from</param>
        public void Import(int projectId)
        {
            //By default, it's not aborted
            this.isAborted = false;

            //Make sure that the handles to the progress dialog and application are available
            if (this.ProgressForm == null)
            {
                throw new ApplicationException("Unable to get handle to progress form. Aborting Import");
            }
            if (this.ProjectApplication == null)
            {
                throw new ApplicationException("Unable to get handle to MS-Project application instance. Aborting Import");
            }

            //Make sure we have a project loaded
            if (this.ProjectApplication.ActiveProject == null || this.ProjectApplication.Projects == null || this.ProjectApplication.Projects.Count == 0)
            {
                throw new ApplicationException("No MS-Project project is currently loaded. Please open the MS-Project file you wish to import to.");
            }

             //Start the background thread that performs the import
            ImportState importState = new ImportState();
            importState.ProjectId = projectId;
            importState.MSProject = this.ProjectApplication.ActiveProject;
            ThreadPool.QueueUserWorkItem(new WaitCallback(this.Import_Process), importState);
        }

        /// <summary>
        /// Exports data from the current MS-Project project to Spira
        /// </summary>
        /// <param name="projectId">The id of the project to export to</param>
        public void Export(int projectId)
        {
            //By default, it's not aborted
            this.isAborted = false;

            //Make sure that the handles to the progress dialog and application are available
            if (this.ProgressForm == null)
            {
                throw new ApplicationException("Unable to get handle to progress form. Aborting Export");
            }
            if (this.ProjectApplication == null)
            {
                throw new ApplicationException("Unable to get handle to MS-Project application instance. Aborting Export");
            }

            //Make sure we have a project loaded
            if (this.ProjectApplication.ActiveProject == null || this.ProjectApplication.Projects == null || this.ProjectApplication.Projects.Count == 0)
            {
                throw new ApplicationException("No MS-Project project is currently loaded. Please open the MS-Project file you wish to export from.");
            }

            //Start the background thread that performs the export
            ImportState importState = new ImportState();
            importState.ProjectId = projectId;
            importState.MSProject = this.ProjectApplication.ActiveProject;
            ThreadPool.QueueUserWorkItem(new WaitCallback(this.Export_Process), importState);
        }

        /// <summary>
        /// This method is responsible for actually importing the data
        /// </summary>
        /// <param name="stateInfo">State information handle</param>
        /// <remarks>This runs in background thread to avoid freezing the progress form</remarks>
        protected void Import_Process(object stateInfo)
        {
            try
            {
                //Get the passed state info
                ImportState importState = (ImportState)stateInfo;

                //Set the progress indicator to 0
                UpdateProgress(0, null);

                //Create the full URL from the provided base URL
                Uri fullUri;
                if (!Importer.TryCreateFullUrl(Configuration.Default.SpiraUrl, out fullUri))
                {
                    throw new ApplicationException("The Server URL entered is not a valid URL");
                }

                SpiraImportExport.SoapServiceClient spiraImportExport = SpiraRibbon.CreateClient(fullUri);

                //Authenticate 
                bool success = spiraImportExport.Connection_Authenticate(Configuration.Default.SpiraUserName, SpiraRibbon.SpiraPassword);
                if (!success)
                {
                    //Authentication failed
                    throw new ApplicationException("Unable to authenticate with Server. Please check the username and password and try again.");
                }

                //Now connect to the project
                success = spiraImportExport.Connection_ConnectToProject(importState.ProjectId);
                if (!success)
                {
                    //Authentication failed
                    throw new ApplicationException("Unable to connect to project PR" + importState.ProjectId + ". Please check that the user is a member of the project.");
                }

                //Now we need to import the various artifacts in turn
                //First lets get all the data from Spira so that we know how many items need to get loaded
                SpiraImportExport.RemoteRelease[] remoteReleases = spiraImportExport.Release_Retrieve(true);
                SpiraImportExport.RemoteRequirement[] remoteRequirements = spiraImportExport.Requirement_Retrieve(null, 1, Int32.MaxValue);
                SpiraImportExport.RemoteSort remoteSort = new SpiraProjectAddIn.SpiraImportExport.RemoteSort();
                remoteSort.PropertyName = "StartDate";
                remoteSort.SortAscending = true;
                SpiraImportExport.RemoteTask[] remoteTasks = spiraImportExport.Task_Retrieve(null, remoteSort, 1, Int32.MaxValue);

                int importCount = remoteReleases.Length + remoteRequirements.Length + remoteTasks.Length;
                int importProgress = 0;
                int errorCount = 0;

                //Set the progress indicator to 0%
                UpdateProgress(importProgress, importCount);

                //Get a handle to the list of tasks in the MS-Project project
                Project project = importState.MSProject;
                Tasks tasks = project.Tasks;

                //First open up the textfile that we will log information to (used for debugging purposes)
                string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_ProjectImportExport.log";
                StreamWriter streamWriter = File.CreateText(debugFile);

                //First we need to iterate through the Releases/Iterations and add as milestones
                //Dictionary<int, int> releaseMapping = new Dictionary<int, int>();
                foreach (SpiraImportExport.RemoteRelease remoteRelease in remoteReleases)
                {
                    try
                    {
                        //See if we already have this task in the MS-Project project
                        bool newTask = false;
                        Task task = FindExistingTask("RL" + remoteRelease.ReleaseId.Value, project);
                        if (task == null)
                        {
                            //Add a task to the project for this requirement
                            task = tasks.Add(remoteRelease.Name, missing);
                            newTask = true;
                        }
                        else
                        {
                            task.Name = remoteRelease.Name;
                        }
                        if (!String.IsNullOrEmpty(remoteRelease.Description))
                        {
                            task.Notes = remoteRelease.Description;
                        }

                        //Get the outline level from the release indent level
                        task.OutlineLevel = (short)(remoteRelease.IndentLevel.Length / 3);

                        //Specify the schedule/effort information
                        //We can't update the date for a 2nd-level item
                        if (newTask || !(bool)task.Summary)
                        {
                            task.Start = remoteRelease.StartDate.ToShortDateString();
                            task.Finish = remoteRelease.EndDate.ToShortDateString();
                        }

                        //Store the SpiraID in the Notes field
                        task.Text1 = TEXT1_PREFIX + "RL" + remoteRelease.ReleaseId.Value;
                        task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteRelease.LastUpdateDate);

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Import aborted by user.");
                        }
                    }
                    catch (System.Exception exception)
                    {
                        //Record the error in the log and add to the error count, then continue
                        streamWriter.WriteLine("Error During Import (RL" + remoteRelease.ReleaseId.Value + ") from Spira > MS-Project: " + exception.Message + " (" + exception.StackTrace + ")");
                        streamWriter.Flush();
                        errorCount++;
                    }
                }

                //Now we need to iterate through the Requirements
                foreach (SpiraImportExport.RemoteRequirement remoteRequirement in remoteRequirements)
                {
                    try
                    {
                        //See if we already have this task in the MS-Project project
                        Task task = FindExistingTask("RQ" + remoteRequirement.RequirementId.Value, project);
                        if (task == null)
                        {
                            //Add a task to the project for this requirement
                            task = tasks.Add(remoteRequirement.Name, missing);
                        }
                        else
                        {
                            task.Name = remoteRequirement.Name;
                        }
                        //The requirements have different INDENT levels, so we arrange the MS-Project tasks according
                        //to the requirements hierarchy, with the Tasks nested underneath.
                        //For now we'll have 3-levels only:
                        short requirementOutlineLevel = (short)(remoteRequirement.IndentLevel.Length / 3);
                        task.OutlineLevel = requirementOutlineLevel;
                        if (!String.IsNullOrEmpty(remoteRequirement.Description))
                        {
                            task.Notes = remoteRequirement.Description;
                        }

                        //If we have a release, set a constraint on the requirement
                        if (remoteRequirement.ReleaseId.HasValue)
                        {
                            //Locate the release
                            int releaseId = remoteRequirement.ReleaseId.Value;
                            foreach (SpiraImportExport.RemoteRelease remoteRelease in remoteReleases)
                            {
                                if (remoteRelease.ReleaseId.Value == releaseId)
                                {
                                    task.ConstraintDate = remoteRelease.EndDate.ToShortDateString();
                                    task.ConstraintType = PjConstraint.pjFNLT;
                                    break;
                                }
                            }
                        }

                        //If this requirement has no tasks, set as a milestone so that it round-trips correctly
                        //(i.e. it doesn't get turned into a task itself on Export)
                        if (remoteRequirement.TaskCount == 0)
                        {
                            task.Milestone = true;
                        }

                        //Store the SpiraID in the Notes field
                        task.Text1 = TEXT1_PREFIX + "RQ" + remoteRequirement.RequirementId.Value;
                        task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteRequirement.LastUpdateDate);

                        //Next find all the tasks that are part of this requirement
                        foreach (SpiraImportExport.RemoteTask remoteTask in remoteTasks)
                        {
                            try
                            {
                                if (remoteTask.RequirementId == remoteRequirement.RequirementId)
                                {
                                    //See if we already have this task in the MS-Project project
                                    task = FindExistingTask("TK" + remoteTask.TaskId.Value, project);
                                    if (task == null)
                                    {
                                        //Add a task to the project for this task
                                        task = tasks.Add(remoteTask.Name, missing);
                                    }
                                    else
                                    {
                                        task.Name = remoteTask.Name;
                                    }
                                    //Nest the task under the requirement
                                    task.OutlineLevel = (short)(requirementOutlineLevel + 1);
                                    if (!String.IsNullOrEmpty(remoteTask.Description))
                                    {
                                        task.Notes = remoteTask.Description;
                                    }

                                    //Specify the schedule/effort information
                                    if (remoteTask.StartDate.HasValue)
                                    {
                                        task.Start = remoteTask.StartDate.Value.ToShortDateString();
                                    }
                                    if (remoteTask.EndDate.HasValue)
                                    {
                                        task.Finish = remoteTask.EndDate.Value.ToShortDateString();
                                    }
                                    if (remoteTask.EstimatedEffort.HasValue)
                                    {
                                        task.Work = remoteTask.EstimatedEffort;
                                        task.Baseline1Work = remoteTask.EstimatedEffort;
                                    }
                                    if (remoteTask.ActualEffort.HasValue)
                                    {
                                        task.ActualWork = remoteTask.ActualEffort;
                                    }
                                    if (remoteTask.RemainingEffort.HasValue)
                                    {
                                        task.RemainingWork = remoteTask.RemainingEffort;
                                    }

                                    //Store the SpiraID in the Notes field and the last updated date for concurrent management
                                    task.Text1 = TEXT1_PREFIX + "TK" + remoteTask.TaskId.Value;
                                    task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteTask.LastUpdateDate);

                                    //Update the progress
                                    importProgress++;
                                    UpdateProgress(importProgress, importCount);
                                }
                                //Check for abort condition
                                if (this.IsAborted)
                                {
                                    throw new ApplicationException("Import aborted by user.");
                                }
                            }
                            catch (System.Exception exception)
                            {
                                //Record the error in the log and add to the error count, then continue
                                streamWriter.WriteLine("Error During Import (TK" + remoteTask.TaskId.Value + ") from Spira > MS-Project: " + exception.Message + " (" + exception.StackTrace + ")");
                                streamWriter.Flush();
                                errorCount++;
                            }
                        }

                        //Update the progress
                        importProgress++;
                        UpdateProgress(importProgress, importCount);

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Import aborted by user.");
                        }
                    }
                    catch (System.Exception exception)
                    {
                        //Record the error in the log and add to the error count, then continue
                        streamWriter.WriteLine("Error During Import (RQ" + remoteRequirement.RequirementId.Value + ") from Spira > MS-Project: " + exception.Message + " (" + exception.StackTrace + ")");
                        streamWriter.Flush();
                        errorCount++;
                    }
                }

                //Only throw one message if an error occurred
                if (errorCount > 0)
                {
                    streamWriter.Close();
                    throw new ApplicationException("Import failed with " + errorCount + " errors. Please check the import error log to view the details.");
                }
                streamWriter.Close();


                //Set the progress indicator to 100%
                UpdateProgress(importCount, importCount);

                //Raise the success event
                OnOperationCompleted(importCount);
            }
            catch (FaultException exception)
            {
                //If we get an exception need to raise an error event that the form displays
                //Need to get the SoapException detail
                MessageFault messageFault = exception.CreateMessageFault();
                if (messageFault == null)
                {
                    OnErrorOccurred(exception.Message);
                }
                else
                {
                    if (messageFault.HasDetail)
                    {
                        XmlElement soapDetails = messageFault.GetDetail<XmlElement>();
                        if (soapDetails == null || soapDetails.ChildNodes.Count == 0)
                        {
                            OnErrorOccurred(messageFault.Reason.ToString());
                        }
                        else
                        {
                            OnErrorOccurred(soapDetails.ChildNodes[0].Value);
                        }
                    }
                    else
                    {
                        OnErrorOccurred(messageFault.Reason.ToString());
                    }
                }
            }
            catch (System.Exception exception)
            {
                //If we get an exception need to raise an error event that the form displays
                OnErrorOccurred(exception.Message);
            }
        }

        /// </summary>
        /// <param name="input">The input string to be truncated</param>
        /// <param name="maxLength">The max-length to truncate to</param>
        /// <returns>The truncated string</returns>
        public static string SafeSubstring(string input, int maxLength)
        {
            if (input.Length > maxLength)
            {
                string output = input.Substring(0, maxLength);
                return output;
            }
            else
            {
                return input;
            }
        }

        /// <summary>
        /// This method is responsible for actually exporting the data
        /// </summary>
        /// <param name="stateInfo">State information handle</param>
        /// <remarks>This runs in background thread to avoid freezing the progress form</remarks>
        protected void Export_Process(object stateInfo)
        {
            try
            {
                //Get the passed state info
                ImportState importState = (ImportState)stateInfo;

                //Set the progress indicator to 0
                UpdateProgress(0, null);

                //Create the full URL from the provided base URL
                Uri fullUri;
                if (!Importer.TryCreateFullUrl(Configuration.Default.SpiraUrl, out fullUri))
                {
                    throw new ApplicationException("The Server URL entered is not a valid URL");
                }

                SpiraImportExport.SoapServiceClient spiraImportExport = SpiraRibbon.CreateClient(fullUri);

                //Authenticate 
                bool success = spiraImportExport.Connection_Authenticate(Configuration.Default.SpiraUserName, SpiraRibbon.SpiraPassword);
                if (!success)
                {
                    //Authentication failed
                    throw new ApplicationException("Unable to authenticate with Server. Please check the username and password and try again.");
                }

                //Now connect to the project
                success = spiraImportExport.Connection_ConnectToProject(importState.ProjectId);
                if (!success)
                {
                    //Authentication failed
                    throw new ApplicationException("Unable to connect to project PR" + importState.ProjectId + ". Please check that the user is a member of the project.");
                }

                //Get a handle to the list of tasks in the MS-Project project
                Project project = importState.MSProject;
                Tasks tasks = project.Tasks;
                int importCount = tasks.Count;
                int importProgress = 0;
                int errorCount = 0;

                //Set the progress indicator to 0%
                UpdateProgress(importProgress, importCount);
                
                //First open up the textfile that we will log information to (used for debugging purposes)
                string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_ProjectImportExport.log";
                StreamWriter streamWriter = File.CreateText(debugFile);

                //Next we need to iterate through all the MS-Project tasks
                short lastOutlineLevel = 1;
                int lastRequirementId = -1;
                foreach (Task task in tasks)
                {
                    try
                    {
                        //If we have the exact text 'IGNORE' in the Text01 custom field we don't import into Spira
                        if (task.Text1 == TEXT1_IGNORE_TOKEN)
                        {
                            streamWriter.WriteLine(String.Format("Ignoring MS-Project Task {0} since it has the IGNORE value set in its Text1 field.", task.ID));
                        }
                        else
                        {
                            //See if this is a Requirement or Task
                            //1) All summary items are Requirements
                            //2) Non-summary items are Tasks unless they are (a) milestones (zero effort) or (b) top-level items
                            bool isRequirement = false;
                            if ((bool)task.Summary || (bool)task.Milestone || task.OutlineLevel == 1)
                            {
                                isRequirement = true;
                            }

                            //Create/Update either the requirement or task
                            if (isRequirement)
                            {
                                //See if this is the Insert or Update case based on whether we have a matching value in Text1
                                int existingRequirementId = -1;
                                int prefixLength = TEXT1_PREFIX.Length;
                                if (!String.IsNullOrEmpty(task.Text1) && SafeSubstring(task.Text1, prefixLength) == TEXT1_PREFIX)
                                {
                                    //Make sure it is already listed as a requirements and not a task
                                    //If it was a task and it's now a requirement (i.e. we have added child items), treat as the Insert case
                                    if (task.Text1.Length > prefixLength + 2 && task.Text1.Substring(prefixLength, 2) == "RQ")
                                    {
                                        string requirementIdString = task.Text1.Substring(prefixLength + 2);
                                        int parsedValue;
                                        if (Int32.TryParse(requirementIdString, out parsedValue))
                                        {
                                            existingRequirementId = parsedValue;
                                        }
                                    }
                                }

                                //Create the new data object or retrieve the existing one
                                SpiraImportExport.RemoteRequirement remoteRequirement;
                                if (existingRequirementId == -1)
                                {
                                    remoteRequirement = new SpiraProjectAddIn.SpiraImportExport.RemoteRequirement();
                                }
                                else
                                {
                                    remoteRequirement = spiraImportExport.Requirement_RetrieveById(existingRequirementId);

                                    //See if we have a date-time in Text2 for concurrency management
                                    if (String.IsNullOrEmpty(task.Text2))
                                    {
                                        task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteRequirement.LastUpdateDate);
                                    }
                                    else
                                    {
                                        DateTime concurrencyDateTime;
                                        if (DateTime.TryParseExact(task.Text2, FORMAT_DATE_TIME_INVARIANT_PARSE, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out concurrencyDateTime))
                                        {
                                            remoteRequirement.LastUpdateDate = concurrencyDateTime;
                                        }
                                    }
                                }
                                //Populate the data
                                remoteRequirement.Name = task.Name;
                                remoteRequirement.Description = task.Notes;
                                remoteRequirement.StatusId = 2; //Planned
                                remoteRequirement.RequirementTypeId = 2; //Feature
                                //streamWriter.WriteLine(task.PercentComplete.GetType().ToString());
                                //streamWriter.Flush();
                                if ((short)task.PercentComplete == 0)
                                {
                                    remoteRequirement.StatusId = 2; //Planned;
                                }
                                if ((short)task.PercentComplete > 0 && (short)task.PercentComplete < 100)
                                {
                                    remoteRequirement.StatusId = 3; //In Progress;
                                }
                                if ((short)task.PercentComplete == 100)
                                {
                                    remoteRequirement.StatusId = 4; //Completed;
                                }
                                if (task.Work != null)
                                {
                                    //streamWriter.WriteLine(task.Work.GetType().ToString());
                                    //streamWriter.Flush();
                                    double work = (double)task.Work;

                                    //We convert into function points assuming 8 hours = 1 point.
                                    remoteRequirement.EstimatePoints = (decimal)(work / 8.0);
                                }

                                //Save the data and add the ID to the task
                                int indentOffset = task.OutlineLevel - lastOutlineLevel;
                                if (existingRequirementId == -1)
                                {
                                    remoteRequirement = spiraImportExport.Requirement_Create1(remoteRequirement, indentOffset);
                                    task.Text1 = TEXT1_PREFIX + "RQ" + remoteRequirement.RequirementId.Value;
                                }
                                else
                                {
                                    spiraImportExport.Requirement_Update(remoteRequirement);
                                }
                                lastOutlineLevel = task.OutlineLevel;
                                lastRequirementId = remoteRequirement.RequirementId.Value;

                                //Now need to re-retrieve the requirement to update the date for concurrency management
                                remoteRequirement = spiraImportExport.Requirement_RetrieveById(remoteRequirement.RequirementId.Value);
                                task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteRequirement.LastUpdateDate);
                            }
                            else
                            {
                                //See if this is the Insert or Update case based on whether we have a matching value in Text1
                                int existingTaskId = -1;
                                int prefixLength = TEXT1_PREFIX.Length;
                                if (!String.IsNullOrEmpty(task.Text1) && SafeSubstring(task.Text1, prefixLength) == TEXT1_PREFIX)
                                {
                                    //Make sure it is already listed as a task and not a requirement
                                    //If it was a requirement and it's now a task (i.e. we have removed child items), treat as the Insert case
                                    if (task.Text1.Length > prefixLength + 2 && task.Text1.Substring(prefixLength, 2) == "TK")
                                    {
                                        string taskIdString = task.Text1.Substring(prefixLength + 2);
                                        int parsedValue;
                                        if (Int32.TryParse(taskIdString, out parsedValue))
                                        {
                                            existingTaskId = parsedValue;
                                        }
                                    }
                                }

                                //Create the new data object or retrieve the existing one
                                SpiraImportExport.RemoteTask remoteTask;
                                if (existingTaskId == -1)
                                {
                                    remoteTask = new SpiraProjectAddIn.SpiraImportExport.RemoteTask();
                                }
                                else
                                {
                                    remoteTask = spiraImportExport.Task_RetrieveById(existingTaskId);

                                    //See if we have a date-time in Text2 for concurrency management
                                    if (String.IsNullOrEmpty(task.Text2))
                                    {
                                        task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteTask.LastUpdateDate);
                                    }
                                    else
                                    {
                                        DateTime concurrencyDateTime;
                                        if (DateTime.TryParseExact(task.Text2, FORMAT_DATE_TIME_INVARIANT_PARSE, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out concurrencyDateTime))
                                        {
                                            remoteTask.LastUpdateDate = concurrencyDateTime;
                                        }
                                    }
                                }

                                //Populate the data
                                remoteTask.Name = task.Name;
                                remoteTask.Description = task.Notes;
                                remoteTask.RequirementId = lastRequirementId;
                                //streamWriter.WriteLine(task.Start.GetType().ToString());
                                //streamWriter.Flush();
                                remoteTask.StartDate = (DateTime)task.Start;
                                remoteTask.EndDate = (DateTime)task.Finish;

                                remoteTask.TaskStatusId = 1; //Not Started
                                remoteTask.TaskTypeId = 5; //Other
                                //streamWriter.WriteLine(task.PercentComplete.GetType().ToString());
                                //streamWriter.Flush();
                                if ((short)task.PercentComplete == 0)
                                {
                                    remoteTask.TaskStatusId = 1; //Not Started
                                }
                                if ((short)task.PercentComplete > 0 && (short)task.PercentComplete < 100)
                                {
                                    remoteTask.TaskStatusId = 2; //In Progress;
                                }
                                if ((short)task.PercentComplete == 100)
                                {
                                    remoteTask.TaskStatusId = 3; //Completed;
                                }
                                if (task.Work != null)
                                {
                                    //streamWriter.WriteLine(task.Work.GetType().ToString());
                                    //streamWriter.Flush();
                                    double work = (double)task.Work;
                                    remoteTask.EstimatedEffort = (int)work;
                                }
                                //Baseline is used to store the original work if populated
                                if (task.BaselineWork != null)
                                {
                                    double work = (double)task.BaselineWork;
                                    remoteTask.EstimatedEffort = (int)work;
                                }
                                if (task.ActualWork != null)
                                {
                                    //streamWriter.WriteLine(task.Work.GetType().ToString());
                                    //streamWriter.Flush();
                                    double actualWork = (double)task.ActualWork;
                                    remoteTask.ActualEffort = (int)actualWork;
                                }
                                if (task.RemainingWork != null)
                                {
                                    //streamWriter.WriteLine(task.Work.GetType().ToString());
                                    //streamWriter.Flush();
                                    double remainingWork = (double)task.RemainingWork;
                                    remoteTask.RemainingEffort = (int)remainingWork;
                                }

                                //Save the data and add the ID to the task
                                if (existingTaskId == -1)
                                {
                                    remoteTask = spiraImportExport.Task_Create(remoteTask);
                                    task.Text1 = TEXT1_PREFIX + "TK" + remoteTask.TaskId.Value;
                                }
                                else
                                {
                                    spiraImportExport.Task_Update(remoteTask);
                                }

                                //Now need to re-retrieve the task to update the date for concurrency management
                                remoteTask = spiraImportExport.Task_RetrieveById(remoteTask.TaskId.Value);
                                task.Text2 = String.Format(FORMAT_DATE_TIME_INVARIANT, remoteTask.LastUpdateDate);
                            }
                        }

                        //Update the progress
                        importProgress++;
                        UpdateProgress(importProgress, importCount);

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Export aborted by user.");
                        }
                    }
                    catch (System.Exception exception)
                    {
                        //Record the error in the log and add to the error count, then continue
                        streamWriter.WriteLine("Error During Export (Task " + task.ID + ") from MS-Project > Spira: " + exception.Message + " (" + exception.StackTrace + ")");
                        streamWriter.Flush();
                        errorCount++;
                    }
                }

                //Only throw one message if an error occurred
                if (errorCount > 0)
                {
                    streamWriter.Close();
                    throw new ApplicationException("Export failed with " + errorCount + " errors. Please check the export error log to view the details.");
                }
                streamWriter.Close();


                //Set the progress indicator to 100%
                UpdateProgress(importCount, importCount);

                //Raise the success event
                OnOperationCompleted(importCount);
            }
            catch (FaultException exception)
            {
                //If we get an exception need to raise an error event that the form displays
                //Need to get the SoapException detail
                MessageFault messageFault = exception.CreateMessageFault();
                if (messageFault == null)
                {
                    OnErrorOccurred(exception.Message);
                }
                else
                {
                    if (messageFault.HasDetail)
                    {
                        XmlElement soapDetails = messageFault.GetDetail<XmlElement>();
                        if (soapDetails == null || soapDetails.ChildNodes.Count == 0)
                        {
                            OnErrorOccurred(messageFault.Reason.ToString());
                        }
                        else
                        {
                            OnErrorOccurred(soapDetails.ChildNodes[0].Value);
                        }
                    }
                    else
                    {
                        OnErrorOccurred(messageFault.Reason.ToString());
                    }
                }
            }
            catch (System.Exception exception)
            {
                //If we get an exception need to raise an error event that the form displays
                OnErrorOccurred(exception.Message);
            }
        }

        /// <summary>
        /// Sees if we have the artifact already in the MS-Project file
        /// </summary>
        /// <param name="artifactPrefixAndId">The artifact prefix and ID (e.g. RQ1)</param>
        /// <returns>The task if we have a match or null if not</returns>\
        /// <param name="project">Handle to the MS-Project project</param>
        private Task FindExistingTask(string artifactPrefixAndId, Project project)
        {
            Task matchedTask = null;
            int prefixLength = TEXT1_PREFIX.Length;
            foreach (Task task in project.Tasks)
            {
                if (!String.IsNullOrEmpty(task.Text1) && SafeSubstring(task.Text1, prefixLength) == TEXT1_PREFIX)
                {
                    //See if there is a match on the field
                    if (task.Text1.Length > prefixLength && task.Text1.Substring(prefixLength) == artifactPrefixAndId)
                    {
                        matchedTask = task;
                    }
                }

            }

            return matchedTask;
        }
    }

    /// <summary>
    /// The arguments for the ImportExportError event
    /// </summary>
    public class ImportExportErrorArgs : System.EventArgs
    {
        public ImportExportErrorArgs(string message)
        {
            this.Message = message;
        }

        public string Message
        {
            get;
            set;
        }
    }

    /// <summary>
    /// The arguments for the ImportExportCompleted event
    /// </summary>
    public class ImportExportCompletedArgs : System.EventArgs
    {
        public ImportExportCompletedArgs(int rowCount)
        {
            this.RowCount = rowCount;
        }

        public int RowCount
        {
            get;
            set;
        }
    }
}
