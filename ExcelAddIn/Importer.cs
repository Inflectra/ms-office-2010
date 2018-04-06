using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Xml;

using Microsoft.Office.Interop.Excel;
using SpiraExcelAddIn.SpiraImportExport;

namespace SpiraExcelAddIn
{
    /// <summary>
    /// Contains the logic to import/export data from SpiraTeam to/from the Excel sheet
    /// </summary>
    public class Importer
    {
        public const string SOAP_RELATIVE_URL = "Services/v5_0/SoapService.svc";

        public const int EXCEL_MAX_CELL_LENGTH = 8000;  //Office 2007 has a greater length than previous versions

        private const int RETRIEVE_PAGE_SIZE = 50;  //We retrieve items in batches of 50

        /// <summary>Enumeration of the different Artifact Types.</summary>
        public enum ArtifactTypeEnum : int
        {
            //Should match entries in TST_ARTIFACT_TYPE, with exception of NONE.
            User = -3,
            Message = -2,
            Project = -1,
            None = 0,
            Requirement = 1,
            TestCase = 2,
            Incident = 3,
            Release = 4,
            TestRun = 5,
            Task = 6,
            TestStep = 7,
            TestSet = 8,
            AutomationHost = 9,
            AutomationEngine = 10,
            Placeholder = 11,
            RequirementStep = 12
        }

        /// <summary>
        /// The various custom property types
        /// </summary>
        public enum CustomPropertyTypeEnum
        {
            Text = 1,
            Integer = 2,
            Decimal = 3,
            Boolean = 4,
            Date = 5,
            List = 6,
            MultiList = 7,
            User = 8
        }

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
        /// The handle to the excel sheet
        /// </summary>
        public Microsoft.Office.Interop.Excel._Application ExcelApplication
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
        /// Sets the string value on an object, making sure that if the existing value is greater than 8000
        /// we leave it alone since Excel2007+ can't handle larger than 8000 character strings, and we don't want
        /// to damage the data already on the server.
        /// </summary>
        /// <param name="propertyInfo">The property descriptor</param>
        /// <param name="value">The value we're setting</param>
        /// <param name="dataObject">The data object</param>
        protected void SafeSetStringValue(PropertyInfo propertyInfo, object dataObject,string value)
        {
            //Get the existing value
            string existingValue = (string)propertyInfo.GetValue(dataObject, null);
            //Set the value if we don't have truncation
            if (String.IsNullOrEmpty(existingValue) || existingValue.Length < EXCEL_MAX_CELL_LENGTH)
            {
                propertyInfo.SetValue(dataObject, value, null);
            }
            //Otherwise do nothing
        }

        /// <summary>
        /// Handles the long description fields so that they don't throw an 'Excel string too long' exception
        /// for text that exceeds 8000 characters
        /// </summary>
        /// <param name="value">The long description value</param>
        /// <returns>The cleaned, truncated string</returns>
        /// <param name="truncated">lets the caller know if it was truncated or not</param>
        /// <remarks>We truncate the string down to 8000 characters after removing formatting (if specified)</remarks>
        protected string CleanTruncateLongText(object value, ref bool truncated)
        {
            if (value == null)
            {
                return "";
            }
            string input = value.ToString();

            //First strip off the formatting (depending on the setting)
            string output;
            if (Configuration.Default.StripRichText)
            {
                output = HtmlRenderAsPlainText(input);
            }
            else
            {
                output = input;
            }

            //Next truncate if necessary
            if (output.Length > EXCEL_MAX_CELL_LENGTH)
            {
                truncated = true;
                output = output.Substring(0, EXCEL_MAX_CELL_LENGTH);
            }
            return output;
        }

        /// <summary>
        /// Renders HTML content as plain text, used to display titles, etc.
        /// </summary>
        /// <param name="source">The HTML markup</param>
        /// <returns>Plain text representation</returns>
        /// <remarks>Handles line-breaks, etc.</remarks>
        public static string HtmlRenderAsPlainText(string source)
        {
            try
            {
                string result;

                //Remove any of the MS-Word special strings
                result = Regex.Replace(source, @"<([ovwxp]:\w+)>.*</([ovwxp]:\w+)>", "", RegexOptions.IgnoreCase);

                // Remove HTML Development formatting
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\r", " ");
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\n", " ");
                // Remove step-formatting
                result = result.Replace("\t", string.Empty);
                // Remove repeating speces becuase browsers ignore them
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"( )+", " ");

                // Remove the header (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*head([^>])*>", "<head>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<( )*(/)( )*head( )*>)", "</head>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(<head>).*(</head>)", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all scripts (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*script([^>])*>", "<script>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<( )*(/)( )*script( )*>)", "</script>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                //result = System.Text.RegularExpressions.Regex.Replace(result, 
                //         @"(<script>)([^(<script>\.</script>)])*(</script>)",
                //         string.Empty, 
                //         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<script>).*(</script>)", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all styles (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*style([^>])*>", "<style>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"(<( )*(/)( )*style( )*>)", "</style>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(<style>).*(</style>)", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert tabs in spaces of <td> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*td([^>])*>", "\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line breaks in places of <BR> and <LI> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*br( )*>", "\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*li( )*>", "\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line paragraphs (double line breaks) in place
                // if <P>, <DIV> and <TR> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*div([^>])*>", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*tr([^>])*>", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<( )*p([^>])*>", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Remove remaining tags like <a>, links, images,
                // comments etc - anything thats enclosed inside < >
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<[^>]*>", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // replace special characters:
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&nbsp;", " ",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&bull;", " * ",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&lsaquo;", "<",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&rsaquo;", ">",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&trade;", "(tm)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&frasl;", "/",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"<", "<",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @">", ">",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&copy;", "(c)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&reg;", "(r)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove all others. More can be added, see
                // http://hotwired.lycos.com/webmonkey/reference/special_characters/
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    @"&(.{2,6});", string.Empty,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // for testng
                //System.Text.RegularExpressions.Regex.Replace(result, 
                //       this.txtRegex.Text,string.Empty, 
                //       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // make line breaking consistent
                result = result.Replace("\n", "\r");

                // Remove extra line breaks and tabs:
                // replace over 2 breaks with 2 and over 4 tabs with 4. 
                // Prepare first to remove any whitespaces inbetween
                // the escaped characters and remove redundant tabs inbetween linebreaks
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)( )+(\r)", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\t)( )+(\t)", "\t\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\t)( )+(\r)", "\t\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)( )+(\t)", "\r\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove redundant tabs
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)(\t)+(\r)", "\r\r",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove multible tabs followind a linebreak with just one tab
                result = System.Text.RegularExpressions.Regex.Replace(result,
                    "(\r)(\t)+", "\r\t",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Initial replacement target string for linebreaks
                string breaks = "\r\r\r";
                // Initial replacement target string for tabs
                string tabs = "\t\t\t\t\t";
                for (int index = 0; index < result.Length; index++)
                {
                    result = result.Replace(breaks, "\r\r");
                    result = result.Replace(tabs, "\t\t\t\t");
                    breaks = breaks + "\r";
                    tabs = tabs + "\t";
                }

                // Thats it.
                return result;

            }
            catch
            {
                return source;
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
                if (currentValue.HasValue && currentValue.Value <= this.ProgressForm.ProgressMaximumValue)
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
        /// Makes a string safe for use in XML (e.g. web service)
        /// </summary>
        /// <param name="input">The input string (as object)</param>
        /// <returns>The output string</returns>
        protected static string MakeXmlSafe(object input)
        {
            //Handle null reference case
            if (input == null)
            {
                return "";
            }

            //Handle empty string case
            string inputString = (string)input;
            if (inputString == "")
            {
                return inputString;
            }

            string output = inputString.Replace("\x00", "");
            output = output.Replace("\n", "");
            output = output.Replace("\r", "");
            output = output.Replace("\x01", "");
            output = output.Replace("\x02", "");
            output = output.Replace("\x03", "");
            output = output.Replace("\x04", "");
            output = output.Replace("\x05", "");
            output = output.Replace("\x06", "");
            output = output.Replace("\x07", "");
            output = output.Replace("\x08", "");
            output = output.Replace("\x0B", "");
            output = output.Replace("\x0C", "");
            output = output.Replace("\x0E", "");
            output = output.Replace("\x0F", "");
            output = output.Replace("\x10", "");
            output = output.Replace("\x11", "");
            output = output.Replace("\x12", "");
            output = output.Replace("\x13", "");
            output = output.Replace("\x14", "");
            output = output.Replace("\x15", "");
            output = output.Replace("\x16", "");
            output = output.Replace("\x17", "");
            output = output.Replace("\x18", "");
            output = output.Replace("\x19", "");
            output = output.Replace("\x1A", "");
            output = output.Replace("\x1B", "");
            output = output.Replace("\x1C", "");
            output = output.Replace("\x1D", "");
            output = output.Replace("\x1E", "");
            output = output.Replace("\x1F", "");
            return output;
        }

        /// <summary>
        /// Returns the index of the error column
        /// </summary>
        /// <param name="fieldColumnMapping"></param>
        /// <param name="customPropertyMapping"></param>
        /// <returns></returns>
        private int GetErrorColumn(Dictionary<string, int> fieldColumnMapping, Dictionary<int, int> customPropertyMapping = null)
        {
            //The error column is the column after the last data column
            int errorColumn = 1;
            foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
            {
                int columnIndex = fieldColumnPair.Value;
                if (columnIndex > errorColumn)
                {
                    errorColumn = columnIndex;
                }
            }
            if (customPropertyMapping != null)
            {
                foreach (KeyValuePair<int, int> fieldColumnPair in customPropertyMapping)
                {
                    int columnIndex = fieldColumnPair.Value;
                    if (columnIndex > errorColumn)
                    {
                        errorColumn = columnIndex;
                    }
                }
            }
            errorColumn++;
            return errorColumn;
        }

        /// <summary>
        /// Imports data from Spira to the excel sheet
        /// </summary>
        /// <param name="artifactTypeName">The type of data we're importing</param>
        /// <param name="projectId">The id of the project to import from</param>
        public void Import(int projectId, string artifactTypeName)
        {
            //By default, it's not aborted
            this.isAborted = false;

            //Make sure that the handles to the progress dialog and application are available
            if (this.ProgressForm == null)
            {
                throw new ApplicationException("Unable to get handle to progress form. Aborting Import");
            }
            if (this.ExcelApplication == null)
            {
                throw new ApplicationException("Unable to get handle to Excel application instance. Aborting Import");
            }

            //Make sure we have a workbook loaded
            if (this.ExcelApplication.ActiveWorkbook == null || this.ExcelApplication.Worksheets == null || this.ExcelApplication.Worksheets.Count == 0)
            {
                throw new ApplicationException("No Excel worksheet is currently loaded. Please open the Excel import template");
            }

            //Make sure that the required worksheets exist
            Worksheet importWorksheet = null;
            foreach (Worksheet worksheet in this.ExcelApplication.Worksheets)
            {
                if (worksheet.Name.Trim().ToLowerInvariant() == artifactTypeName.Trim().ToLowerInvariant())
                {
                    importWorksheet = worksheet;
                    break;
                }
            }
            if (importWorksheet == null)
            {
                throw new ApplicationException("Unable to locate a worksheet with name '" + artifactTypeName + "'. Aborting Import");
            }

            //Worksheet containing lookups
            Worksheet lookupWorksheet = null;
            foreach (Worksheet worksheet in this.ExcelApplication.Worksheets)
            {
                if (worksheet.Name.Trim().ToLowerInvariant() == "lookups")
                {
                    lookupWorksheet = worksheet;
                    break;
                }
            }
            if (lookupWorksheet == null)
            {
                throw new ApplicationException("Unable to locate a worksheet with name 'Lookups'. Aborting Import");
            }

            //Start the background thread that performs the import
            ImportState importState = new ImportState();
            importState.ProjectId = projectId;
            importState.ArtifactTypeName = artifactTypeName;
            importState.ExcelWorksheet = importWorksheet;
            importState.LookupWorksheet = lookupWorksheet;
            ThreadPool.QueueUserWorkItem(new WaitCallback(this.Import_Process), importState);
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

                //Now see which data is being imported and handle accordingly
                int importCount = 0;
                switch (importState.ArtifactTypeName)
                {
                    case "Requirements":
                        importCount = ImportRequirements(spiraImportExport, importState);
                        break;
                    case "Releases":
                        importCount = ImportReleases(spiraImportExport, importState);
                        break;
                    case "Test Sets":
                        importCount = ImportTestSets(spiraImportExport, importState);
                        break;
                    case "Test Cases":
                        importCount = ImportTestCases(spiraImportExport, importState);
                        break;
                    case "Test Runs":
                        importCount = ImportTestRuns(spiraImportExport, importState);
                        break;
                    case "Incidents":
                        importCount = ImportIncidents(spiraImportExport, importState);
                        break;
                    case "Tasks":
                        importCount = ImportTasks(spiraImportExport, importState);
                        break;
                    case "Custom Values":
                        importCount = ImportCustomValues(spiraImportExport, importState);
                        break;
                }

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
            catch (Exception exception)
            {
                //If we get an exception need to raise an error event that the form displays
                OnErrorOccurred(exception.Message);
            }
        }

        /// <summary>
        /// Imports incidents
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportIncidents(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Inc_Type");
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Inc_Status");
            Dictionary<int, string> priorityMapping = LoadLookup(importState.LookupWorksheet, "Inc_Priority");
            Dictionary<int, string> severityMapping = LoadLookup(importState.LookupWorksheet, "Inc_Severity");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Incident, false);

            //Get the list of components currently in this project
            RemoteComponent[] components = spiraImportExport.Component_Retrieve(true, false);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Inc #", "IncidentId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Incident Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Incident Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "IncidentTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "IncidentStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Priority", "PriorityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Severity", "SeverityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Detected Release", "DetectedReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Resolved Release", "ResolvedReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Detector", "OpenerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Est. Effort", "EstimatedEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Act. Effort", "ActualEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Rem. Effort", "RemainingEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Detected Date", "CreationDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Closed Date", "ClosedDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Components", "ComponentIds", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("IncidentId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident Name'");
            }
            if (!fieldColumnMapping.ContainsKey("Description"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident Description'");
            }
            if (!fieldColumnMapping.ContainsKey("IncidentStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("IncidentTypeId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Type'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Retrieve all the incidents in the project
            SpiraImportExport.RemoteSort remoteSort = new SpiraExcelAddIn.SpiraImportExport.RemoteSort();
            remoteSort.PropertyName = "IncidentId";
            remoteSort.SortAscending = true;
            int startRow = 1;
            bool noMoreData = false;
            int artifactCount = 0;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            int rowIndex = 1;
            while (!noMoreData)
            {
                SpiraImportExport.RemoteIncident[] remoteIncidents = spiraImportExport.Incident_Retrieve(null, remoteSort, startRow, RETRIEVE_PAGE_SIZE);
                if (remoteIncidents == null || remoteIncidents.Length == 0)
                {
                    noMoreData = true;
                    break;
                }
                artifactCount += remoteIncidents.Length;

                //Set the progress bar accordingly
                this.UpdateProgress(0, RETRIEVE_PAGE_SIZE);

                //Now iterate through the incidents and populate the fields
                int currentItemInPage = 1;
                foreach (SpiraImportExport.RemoteIncident remoteIncident in remoteIncidents)
                {
                    try
                    {
                        //For performance using VSTO Interop we need to update all the fields in the row in one go
                        Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                        object[,] dataValues = (object[,])dataRange.Value2;

                        //Iterate through the various mapped fields
                        bool oldTruncated = truncated;
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //See if this field exists on the remote object
                            Type remoteObjectType = remoteIncident.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanRead)
                            {
                                //See if we have one of the special known lookups
                                //or if we have to convert data-types
                                object propertyValue = propertyInfo.GetValue(remoteIncident, null);
                                if (fieldName == "IncidentTypeId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (typeMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = typeMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "IncidentStatusId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (statusMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = statusMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "PriorityId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (priorityMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = priorityMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "SeverityId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (severityMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = severityMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "ComponentIds")
                                {
                                    if (propertyValue == null)
                                    {
                                        dataValues[1, columnIndex] = "";
                                    }
                                    else
                                    {
                                        int[] componentIds = (int[])propertyValue;
                                        if (componentIds.Length > 0)
                                        {
                                            string componentNames = "";
                                            foreach (int componentId in componentIds)
                                            {
                                                RemoteComponent component = components.FirstOrDefault(c => c.ComponentId == componentId);
                                                if (component != null)
                                                {
                                                    if (componentNames == "")
                                                    {
                                                        componentNames = component.Name;
                                                    }
                                                    else
                                                    {
                                                        componentNames += "," + component.Name;
                                                    }
                                                }
                                            }
                                            dataValues[1, columnIndex] = componentNames;
                                        }
                                        else
                                        {
                                            dataValues[1, columnIndex] = "";
                                        }
                                    }
                                }
                                else if (fieldName == "Description")
                                {
                                    //Need to strip off any formatting and make sure it's not too long
                                    dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                }
                                else
                                {
                                    if (propertyInfo.PropertyType == typeof(bool))
                                    {
                                        bool flagValue = (bool)propertyValue;
                                        dataValues[1, columnIndex] = (flagValue) ? "Y" : "N";
                                    }
                                    else if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        //For strings we need to verify length and truncate if necessary
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
                                    {
                                        //For dates we need to convert to local-time
                                        if (propertyValue is DateTime)
                                        {
                                            DateTime dateTimeValue = (DateTime)propertyValue;
                                            dataValues[1, columnIndex] = dateTimeValue.ToLocalTime();
                                        }
                                        else if (propertyValue is DateTime?)
                                        {
                                            DateTime? dateTimeValue = (DateTime?)propertyValue;
                                            if (dateTimeValue.HasValue)
                                            {
                                                dataValues[1, columnIndex] = dateTimeValue.Value.ToLocalTime();
                                            }
                                            else
                                            {
                                                dataValues[1, columnIndex] = null;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dataValues[1, columnIndex] = propertyValue;
                                    }
                                }
                            }
                        }

                        //Iterate through all the custom properties
                        ImportCustomProperties(remoteIncident, customProperties, dataValues, customPropertyMapping);

                        //Now commit the data
                        dataRange.Value2 = dataValues;

                        //If it was truncated on this row, display a message in the right-most column
                        if (truncated && !oldTruncated)
                        {
                            Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                            errorCell.Value2 = "This row had data truncated.";
                        }

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Import aborted by user.");
                        }

                        //Move to the next row and update progress bar
                        rowIndex++;
                        importCount++;
                        this.UpdateProgress(currentItemInPage, null);
                    }
                    catch (Exception exception)
                    {
                        //Record the error on the sheet and add to the error count, then continue
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = exception.Message;
                        errorCount++;
                    }
                    currentItemInPage++;
                }
                startRow += RETRIEVE_PAGE_SIZE;
            }
            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Loads the lookups used in the importer into the various mapping dictionaries
        /// </summary>
        /// <param name="lookupWorksheet">The worksheet containing the lookups</param>
        /// <returns>The mapping dictionary</returns>
        /// <param name="rangeName">The name of the range</param>
        private Dictionary<int, string> LoadLookup(Worksheet lookupWorksheet, string rangeName)
        {
            Dictionary<int, string> lookupMapping = new Dictionary<int, string>();
            Range namedRange = lookupWorksheet.Range[rangeName, missing];
            
            //We need to get the range that consists of this named range and the matching cells to the right
            Range cell1 = (Range)lookupWorksheet.Cells[namedRange.Row, namedRange.Column];
            Range cell2 = (Range)lookupWorksheet.Cells[namedRange.Row + namedRange.Rows.Count - 1, namedRange.Column+1];
            Range lookupRange = lookupWorksheet.Range[cell1, cell2];
            object[,] lookupData = (object[,])lookupRange.Value2;

            for (int i = 1; i <= namedRange.Rows.Count; i++)
            {
                //Get the data in this range (that contains the lookup name)
                string lookupName = (string)lookupData[i, 1];

                //Now by convention, the cell adjacent to the right contains the value
                //(it's not in the named range)
                object lookupValue = lookupData[i, 2];
                int lookupId = -1;
                if (lookupValue is string)
                {
                    if (!Int32.TryParse((string)lookupValue, out lookupId))
                    {
                        lookupId = -1;
                    }
                }
                if (lookupValue is double)
                {
                    lookupId = (int)((double)lookupValue);
                }
                if (lookupValue is int)
                {
                    lookupId = (int)lookupValue;
                }
                if (!String.IsNullOrEmpty(lookupName) && lookupId != -1)
                {
                    if (!lookupMapping.ContainsKey(lookupId))
                    {
                        lookupMapping.Add(lookupId, lookupName);
                    }
                }
            }

            return lookupMapping;
        }

        /// <summary>
        /// Imports requirements
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportRequirements(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Req_Status");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Req_Type");
            Dictionary<int, string> importanceMapping = LoadLookup(importState.LookupWorksheet, "Req_Importance");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Requirement, false);

            //Get the list of components currently in this project
            RemoteComponent[] components = spiraImportExport.Component_Retrieve(true, false);

            //Set the progress bar accordingly
            this.UpdateProgress(0, RETRIEVE_PAGE_SIZE);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Req #", "RequirementId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Version", "ReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Importance", "ImportanceId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "StatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "RequirementTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Author", "AuthorId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Component", "ComponentId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Estimate", "EstimatePoints", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("RequirementId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Req #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Requirement Name'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Now iterate through the requirements and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            int startRow = 1;
            int artifactCount = 0;
            bool noMoreData = false;

            //Retrieve all the requirements in the project in batches
            while (!noMoreData)
            {
                SpiraImportExport.RemoteRequirement[] remoteRequirements = spiraImportExport.Requirement_Retrieve(null, startRow, RETRIEVE_PAGE_SIZE);
                if (remoteRequirements == null || remoteRequirements.Length == 0)
                {
                    noMoreData = true;
                    break;
                }
                artifactCount += remoteRequirements.Length;

                int currentItemInPage = 0;
                foreach (SpiraImportExport.RemoteRequirement remoteRequirement in remoteRequirements)
                {
                    try
                    {
                        //For performance using VSTO Interop we need to update all the fields in the row in one go
                        Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                        object[,] dataValues = (object[,])dataRange.Value2;

                        //Iterate through the various mapped standard fields
                        bool oldTruncated = truncated;
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //See if this field exists on the remote object
                            Type remoteObjectType = remoteRequirement.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanRead)
                            {
                                //If we have the Name field, we need to appropriate update the indent level
                                if (fieldName == "Name")
                                {
                                    Range nameCell = (Range)worksheet.Cells[rowIndex + 2, columnIndex];
                                    nameCell.IndentLevel = (remoteRequirement.IndentLevel.Length / 3) - 1;
                                }

                                //See if we have one of the special known lookups
                                object propertyValue = propertyInfo.GetValue(remoteRequirement, null);
                                if (fieldName == "StatusId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (statusMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = statusMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "ImportanceId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (importanceMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = importanceMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "RequirementTypeId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (typeMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = typeMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "ComponentId")
                                {
                                    if (propertyValue != null && propertyValue is Int32)
                                    {
                                        int componentId = (int)propertyValue;
                                        RemoteComponent component = components.FirstOrDefault(c => c.ComponentId == componentId);
                                        if (component != null)
                                        {
                                            dataValues[1, columnIndex] = component.Name;
                                        }
                                    }
                                }
                                else if (fieldName == "Description")
                                {
                                    //Need to strip off any formatting and make sure it's not too long
                                    dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                }
                                else if (propertyInfo.PropertyType == typeof(string))
                                {
                                    //For strings we need to verify length and truncate if necessary
                                    dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                }
                                else
                                {
                                    dataValues[1, columnIndex] = propertyValue;
                                }
                            }
                        }

                        //Iterate through all the custom properties
                        ImportCustomProperties(remoteRequirement, customProperties, dataValues, customPropertyMapping);

                        //Now commit the data
                        dataRange.Value2 = dataValues;

                        //If the requirement is a summary one, mark field as Bold
                        dataRange.Font.Bold = (remoteRequirement.Summary);

                        //If it was truncated on this row, display a message in the right-most column
                        if (truncated && !oldTruncated)
                        {
                            Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                            errorCell.Value2 = "This row had data truncated.";
                        }

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Import aborted by user.");
                        }

                        //Move to the next row and update progress bar
                        rowIndex++;
                        importCount++;
                        this.UpdateProgress(currentItemInPage, null);
                    }
                    catch (Exception exception)
                    {
                        //Record the error on the sheet and add to the error count, then continue
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = exception.Message;
                        errorCount++;
                    }
                    currentItemInPage++;
                }
                startRow += RETRIEVE_PAGE_SIZE;
            }
            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Imports releases
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportReleases(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Release_Status");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Release_Type");

            //Retrieve all the releases in the project
            SpiraImportExport.RemoteRelease[] remoteReleases = spiraImportExport.Release_Retrieve(false);
            int artifactCount = remoteReleases.Length;

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Release, false);

            //Set the progress bar accordingly
            this.UpdateProgress(0, artifactCount);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Rel #", "ReleaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Version Number", "VersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "ReleaseStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "ReleaseTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Creator", "CreatorId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Start Date", "StartDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "End Date", "EndDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "# Resources", "ResourceCount", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Non-Wk Days", "DaysNonWorking", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("ReleaseId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Rel #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Release Name'");
            }
            if (!fieldColumnMapping.ContainsKey("VersionNumber"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Version Number'");
            }
            if (!fieldColumnMapping.ContainsKey("ReleaseStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("ReleaseTypeId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Type'");
            }
            if (!fieldColumnMapping.ContainsKey("StartDate"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Start Date'");
            }
            if (!fieldColumnMapping.ContainsKey("EndDate"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'End Date'");
            }
            if (!fieldColumnMapping.ContainsKey("ResourceCount"))
            {
                throw new ApplicationException("Unable to find a column heading with name '# Resources'");
            }
            if (!fieldColumnMapping.ContainsKey("DaysNonWorking"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Non-Wk Days'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);
            
            //Now iterate through the releases and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            foreach (SpiraImportExport.RemoteRelease remoteRelease in remoteReleases)
            {
                try
                {
                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated;
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //See if this field exists on the remote object
                        Type remoteObjectType = remoteRelease.GetType();
                        PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                        if (propertyInfo != null && propertyInfo.CanRead)
                        {
                            //If we have the name field, need to update the indent-level
                            if (fieldName == "Name")
                            {
                                Range nameCell = (Range)worksheet.Cells[rowIndex + 2, columnIndex];
                                nameCell.IndentLevel = (remoteRelease.IndentLevel.Length / 3) - 1;
                            }

                            //See if we have one of the special known lookups
                            //or if we have to convert data-types
                            object propertyValue = propertyInfo.GetValue(remoteRelease, null);
                            if (fieldName == "ReleaseStatusId")
                            {
                                if (propertyValue != null)
                                {
                                    int fieldValue = (int)propertyValue;
                                    if (statusMapping.ContainsKey(fieldValue))
                                    {
                                        string lookupValue = statusMapping[fieldValue];
                                        dataValues[1, columnIndex] = lookupValue;
                                    }
                                }
                            }
                            else if (fieldName == "ReleaseTypeId")
                            {
                                if (propertyValue != null)
                                {
                                    int fieldValue = (int)propertyValue;
                                    if (typeMapping.ContainsKey(fieldValue))
                                    {
                                        string lookupValue = typeMapping[fieldValue];
                                        dataValues[1, columnIndex] = lookupValue;
                                    }
                                }
                            }
                            else if (propertyInfo.PropertyType == typeof(bool))
                            {
                                bool flagValue = (bool)propertyValue;
                                dataValues[1, columnIndex] = (flagValue) ? "Y" : "N";
                            }
                            else if (fieldName == "Description")
                            {
                                //Need to strip off any formatting and make sure it's not too long
                                dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                            }
                            else if (propertyInfo.PropertyType == typeof(string))
                            {
                                //For strings we need to verify length and truncate if necessary
                                dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                            }
                            else if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
                            {
                                //For dates we need to convert to local-time
                                if (propertyValue is DateTime)
                                {
                                    DateTime dateTimeValue = (DateTime)propertyValue;
                                    dataValues[1, columnIndex] = dateTimeValue.ToLocalTime();
                                }
                                else if (propertyValue is DateTime?)
                                {
                                    DateTime? dateTimeValue = (DateTime?)propertyValue;
                                    if (dateTimeValue.HasValue)
                                    {
                                        dataValues[1, columnIndex] = dateTimeValue.Value.ToLocalTime();
                                    }
                                    else
                                    {
                                        dataValues[1, columnIndex] = null;
                                    }
                                }
                            }
                            else
                            {
                                dataValues[1, columnIndex] = propertyValue;
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ImportCustomProperties(remoteRelease, customProperties, dataValues, customPropertyMapping);

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //If the release is a summary one, mark field as Bold. If an iteration or phase, mark as italic
                    dataRange.Font.Bold = (remoteRelease.Summary);
                    dataRange.Font.Italic = (remoteRelease.ReleaseTypeId == 3 || remoteRelease.ReleaseTypeId == 4);

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                    //Move to the next row and update progress bar
                    rowIndex++;
                    importCount++;
                    this.UpdateProgress(importCount, null);
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Imports tasks
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportTasks(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Task_Status");
            Dictionary<int, string> priorityMapping = LoadLookup(importState.LookupWorksheet, "Task_Priority");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Task_Type");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Task, false);

            //Retrieve all the tasks in the project
            SpiraImportExport.RemoteSort remoteSort = new SpiraExcelAddIn.SpiraImportExport.RemoteSort();
            remoteSort.PropertyName = "TaskId";
            remoteSort.SortAscending = true;
            SpiraImportExport.RemoteTask[] remoteTasks = spiraImportExport.Task_Retrieve(null, remoteSort, 1, Int32.MaxValue);
            int artifactCount = remoteTasks.Length;

            //Set the progress bar accordingly
            this.UpdateProgress(0, artifactCount);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Task #", "TaskId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Task Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Task Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "TaskStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "TaskTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Priority", "TaskPriorityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement #", "RequirementId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Version", "ReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Start Date", "StartDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "End Date", "EndDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Est. Effort", "EstimatedEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Act. Effort", "ActualEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Rem. Effort", "RemainingEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TaskId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Task #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Task Name'");
            }
            if (!fieldColumnMapping.ContainsKey("TaskStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("TaskTypeId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Type'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Now iterate through the tasks and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            foreach (SpiraImportExport.RemoteTask remoteTask in remoteTasks)
            {
                try
                {
                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated;
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //See if this field exists on the remote object
                        Type remoteObjectType = remoteTask.GetType();
                        PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                        if (propertyInfo != null && propertyInfo.CanRead)
                        {
                            //See if we have one of the special known lookups
                            //or if we have to convert data-types
                            object propertyValue = propertyInfo.GetValue(remoteTask, null);
                            if (fieldName == "TaskStatusId")
                            {
                                if (propertyValue != null)
                                {
                                    int fieldValue = (int)propertyValue;
                                    if (statusMapping.ContainsKey(fieldValue))
                                    {
                                        string lookupValue = statusMapping[fieldValue];
                                        dataValues[1, columnIndex] = lookupValue;
                                    }
                                }
                            }
                            else if (fieldName == "TaskTypeId")
                            {
                                if (propertyValue != null)
                                {
                                    int fieldValue = (int)propertyValue;
                                    if (typeMapping.ContainsKey(fieldValue))
                                    {
                                        string lookupValue = typeMapping[fieldValue];
                                        dataValues[1, columnIndex] = lookupValue;
                                    }
                                }
                            }
                            else if (fieldName == "TaskPriorityId")
                            {
                                if (propertyValue != null)
                                {
                                    int fieldValue = (int)propertyValue;
                                    if (priorityMapping.ContainsKey(fieldValue))
                                    {
                                        string lookupValue = priorityMapping[fieldValue];
                                        dataValues[1, columnIndex] = lookupValue;
                                    }
                                }
                            }
                            else if (fieldName == "Description")
                            {
                                //Need to strip off any formatting and make sure it's not too long
                                dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                            }
                            else
                            {
                                if (propertyInfo.PropertyType == typeof(bool))
                                {
                                    bool flagValue = (bool)propertyValue;
                                    dataValues[1, columnIndex] = (flagValue) ? "Y" : "N";
                                }
                                else if (propertyInfo.PropertyType == typeof(string))
                                {
                                    //For strings we need to verify length and truncate if necessary
                                    dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                }
                                else if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
                                {
                                    //For dates we need to convert to local-time
                                    if (propertyValue is DateTime)
                                    {
                                        DateTime dateTimeValue = (DateTime)propertyValue;
                                        dataValues[1, columnIndex] = dateTimeValue.ToLocalTime();
                                    }
                                    else if (propertyValue is DateTime?)
                                    {
                                        DateTime? dateTimeValue = (DateTime?)propertyValue;
                                        if (dateTimeValue.HasValue)
                                        {
                                            dataValues[1, columnIndex] = dateTimeValue.Value.ToLocalTime();
                                        }
                                        else
                                        {
                                            dataValues[1, columnIndex] = null;
                                        }
                                    }
                                }
                                else
                                {
                                    dataValues[1, columnIndex] = propertyValue;
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ImportCustomProperties(remoteTask, customProperties, dataValues, customPropertyMapping);

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                    //Move to the next row and update progress bar
                    rowIndex++;
                    importCount++;
                    this.UpdateProgress(importCount, null);
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Exports releases
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportReleases(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Release_Status");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Release_Type");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Release, false);

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Rel #", "ReleaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Version Number", "VersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "ReleaseStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "ReleaseTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Creator", "CreatorId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Start Date", "StartDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "End Date", "EndDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "# Resources", "ResourceCount", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Non-Wk Days", "DaysNonWorking", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("ReleaseId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Rel #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Release Name'");
            }
            if (!fieldColumnMapping.ContainsKey("VersionNumber"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Version Number'");
            }
            if (!fieldColumnMapping.ContainsKey("ReleaseStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("ReleaseTypeId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Type'");
            }
            if (!fieldColumnMapping.ContainsKey("StartDate"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Start Date'");
            }
            if (!fieldColumnMapping.ContainsKey("EndDate"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'End Date'");
            }
            if (!fieldColumnMapping.ContainsKey("ResourceCount"))
            {
                throw new ApplicationException("Unable to find a column heading with name '# Resources'");
            }
            if (!fieldColumnMapping.ContainsKey("DaysNonWorking"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Non-Wk Days'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //Now iterate through the rows in the sheet that have data
            Dictionary<int, int> parentPrimaryKeys = new Dictionary<int, int>();
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //See if we are inserting a new release or updating an existing one
                    int primaryKeyColumn = fieldColumnMapping["ReleaseId"];
                    SpiraImportExport.RemoteRelease remoteRelease = null;
                    if (dataValues[1, primaryKeyColumn] == null)
                    {
                        //We have insert case
                        remoteRelease = new SpiraExcelAddIn.SpiraImportExport.RemoteRelease();
                    }
                    else
                    {
                        //We have update case
                        string releaseIdString = dataValues[1, primaryKeyColumn].ToString();
                        int releaseId;
                        if (!Int32.TryParse(releaseIdString, out releaseId))
                        {
                            throw new ApplicationException("Release ID '" + releaseIdString + "' is not valid. It needs to be a purely integer value.");
                        }
                        remoteRelease = spiraImportExport.Release_RetrieveById(releaseId);
                    }

                    //Iterate through the various mapped fields
                    int indentLevel = 0;
                    string newComment = "";
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //If we have the name field, need to use that to determine the indent
                        //and also to know when we've reached the end of the import
                        if (fieldName == "Name")
                        {
                            Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                            if (nameCell == null || nameCell.Value2 == null)
                            {
                                lastRecord = true;
                                break;
                            }
                            else
                            {
                                indentLevel = (int)nameCell.IndentLevel;
                            }
                        }

                        //See if this field exists on the remote object (except Comment which is handled separately)
                        if (fieldName == "Comment")
                        {
                            object dataValue = dataValues[1, columnIndex];
                            if (dataValue != null)
                            {
                                newComment = (MakeXmlSafe(dataValue)).Trim();
                            }
                        }
                        else
                        {
                            Type remoteObjectType = remoteRelease.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];

                                //See if we have one of the special known lookups
                                if (fieldName == "ReleaseStatusId")
                                {
                                    if (dataValue == null)
                                    {
                                        //This field is not nullable, so we need to pass 1 to default to 'Planned'
                                        propertyInfo.SetValue(remoteRelease, 1, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = 1; //Default to Planned
                                        foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteRelease, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "ReleaseTypeId")
                                {
                                    if (dataValue == null)
                                    {
                                        //This field is not nullable, so we need to pass 1 to default to 'Major Release'
                                        propertyInfo.SetValue(remoteRelease, 1, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = 1; //Default to Major Release
                                        foreach (KeyValuePair<int, string> mappingEntry in typeMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteRelease, fieldValue, null);
                                        }
                                    }
                                }
                                else
                                {
                                    //Make sure that we do any necessary type conversion
                                    //Make sure the field handles nullable types
                                    if (dataValue == null)
                                    {
                                        if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                        {
                                            propertyInfo.SetValue(remoteRelease, null, null);
                                        }
                                    }
                                    else
                                    {
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                //Need to handle large string issue
                                                SafeSetStringValue(propertyInfo, remoteRelease, MakeXmlSafe(dataValue));

                                            }
                                            else
                                            {
                                                propertyInfo.SetValue(remoteRelease, dataValue.ToString(), null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(bool))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                bool flagValue = (MakeXmlSafe(dataValue) == "Y");
                                                propertyInfo.SetValue(remoteRelease, flagValue, null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(Nullable<DateTime>))
                                        {
                                            if (dataValue.GetType() == typeof(DateTime))
                                            {
                                                DateTime dateTimeValue = (DateTime)dataValue;
                                                propertyInfo.SetValue(remoteRelease, dateTimeValue.ToUniversalTime(), null);
                                            }
                                            else if (dataValue.GetType() == typeof(double))
                                            {
                                                DateTime dateTimeValue = DateTime.FromOADate((double)dataValue).ToUniversalTime();
                                                propertyInfo.SetValue(remoteRelease, dateTimeValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                DateTime dateTimeValue;
                                                if (DateTime.TryParse(stringValue, out dateTimeValue))
                                                {
                                                    propertyInfo.SetValue(remoteRelease, dateTimeValue.ToUniversalTime(), null);
                                                }
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                        {
                                            if (dataValue.GetType() == typeof(int))
                                            {
                                                propertyInfo.SetValue(remoteRelease, dataValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                int intValue;
                                                if (Int32.TryParse(stringValue, out intValue))
                                                {
                                                    propertyInfo.SetValue(remoteRelease, intValue, null);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ExportCustomProperties(remoteRelease, customProperties, dataValues, customPropertyMapping);

                    if (lastRecord)
                    {
                        break;
                    }

                    //Now either insert or update the release
                    if (remoteRelease.ReleaseId.HasValue)
                    {
                        spiraImportExport.Release_Update(remoteRelease);
                    }
                    else
                    {
                        //Insert case

                        //If we have an item already loaded that is a parent of this, then use the insert child API method
                        if (parentPrimaryKeys.ContainsKey(indentLevel - 1))
                        {
                            //Specify defaults for any required fields that might not have been set
                            if (remoteRelease.ResourceCount == 0)
                            {
                                remoteRelease.ResourceCount = 1;
                            }
                            remoteRelease = spiraImportExport.Release_Create(remoteRelease, parentPrimaryKeys[indentLevel - 1]);
                        }
                        else
                        {
                            //Otherwise just insert at the end
                            remoteRelease = spiraImportExport.Release_Create(remoteRelease, null);
                        }
                        //Update the cell with the release ID to prevent multiple-appends
                        Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                        newKeyCell.Value2 = remoteRelease.ReleaseId;
                    }
                    //Add to the parent indent dictionary (not iterations)
                    if (remoteRelease.ReleaseId.HasValue && (remoteRelease.ReleaseTypeId == 1 || remoteRelease.ReleaseTypeId == 2))
                    {
                        if (!parentPrimaryKeys.ContainsKey(indentLevel))
                        {
                            parentPrimaryKeys.Add(indentLevel, remoteRelease.ReleaseId.Value);
                        }
                        else
                        {
                            parentPrimaryKeys[indentLevel] = remoteRelease.ReleaseId.Value;
                        }
                    }

                    //Add a comment if necessary
                    if (newComment != "")
                    {
                        SpiraImportExport.RemoteComment remoteComment = new SpiraImportExport.RemoteComment();
                        remoteComment.ArtifactId = remoteRelease.ReleaseId.Value;
                        remoteComment.Text = newComment;
                        remoteComment.CreationDate = DateTime.UtcNow;
                        spiraImportExport.Release_CreateComment(remoteComment);
                    }

                    //Move to the next row and update progress bar
                    exportCount++;
                    this.UpdateProgress(exportCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Exports tasks
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportTasks(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Task_Status");
            Dictionary<int, string> priorityMapping = LoadLookup(importState.LookupWorksheet, "Task_Priority");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Task_Type");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Task, false);

            //Get the list of releases currently in this project
            RemoteRelease[] releases = spiraImportExport.Release_Retrieve(true);

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Task #", "TaskId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Task Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Task Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "TaskStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "TaskTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Priority", "TaskPriorityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement #", "RequirementId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Version", "ReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Start Date", "StartDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "End Date", "EndDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Est. Effort", "EstimatedEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Act. Effort", "ActualEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Rem. Effort", "RemainingEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TaskId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Task #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Task Name'");
            }
            if (!fieldColumnMapping.ContainsKey("TaskStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("TaskTypeId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Type'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //Now iterate through the rows in the sheet that have data
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //See if we are inserting a new task or updating an existing one
                    int primaryKeyColumn = fieldColumnMapping["TaskId"];
                    SpiraImportExport.RemoteTask remoteTask = null;
                    if (dataValues[1, primaryKeyColumn] == null)
                    {
                        //We have insert case
                        remoteTask = new SpiraExcelAddIn.SpiraImportExport.RemoteTask();
                    }
                    else
                    {
                        //We have update case
                        string taskIdString = dataValues[1, primaryKeyColumn].ToString();
                        int taskId;
                        if (!Int32.TryParse(taskIdString, out taskId))
                        {
                            throw new ApplicationException("Task ID '" + taskIdString + "' is not valid. It needs to be a purely integer value.");
                        }
                        remoteTask = spiraImportExport.Task_RetrieveById(taskId);
                    }

                    //Iterate through the various mapped fields
                    string newComment = "";
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //If we have the name field, need to use that to know when we've reached the end of the import
                        if (fieldName == "Name")
                        {
                            Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                            if (nameCell == null || nameCell.Value2 == null)
                            {
                                lastRecord = true;
                                break;
                            }
                        }

                        //See if this field exists on the remote object (except Comment which is handled separately)
                        if (fieldName == "Comment")
                        {
                            object dataValue = dataValues[1, columnIndex];
                            if (dataValue != null)
                            {
                                newComment = (MakeXmlSafe(dataValue)).Trim();
                            }
                        }
                        else
                        {
                            Type remoteObjectType = remoteTask.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];

                                //See if we have one of the special known lookups
                                if (fieldName == "TaskStatusId")
                                {
                                    if (dataValue == null)
                                    {
                                        //This field is not nullable, so we need to pass 1 to default to 'Not Started'
                                        propertyInfo.SetValue(remoteTask, 1, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = 1; //Default to requested
                                        foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteTask, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "TaskTypeId")
                                {
                                    if (dataValue == null)
                                    {
                                        //This field is not nullable, so we need to pass 1 to default to 'Development'
                                        propertyInfo.SetValue(remoteTask, 1, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = 1; //Default to requested
                                        foreach (KeyValuePair<int, string> mappingEntry in typeMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteTask, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "TaskPriorityId")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteTask, null, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = -1;
                                        foreach (KeyValuePair<int, string> mappingEntry in priorityMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteTask, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "ReleaseVersionNumber")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteTask, null, null);
                                    }
                                    else if (dataValue is String)
                                    {
                                        //We need to get the version number and find the corresponding ReleaseId, if it exists
                                        string versionNumber = (string)dataValue;
                                        RemoteRelease release = releases.FirstOrDefault(r => r.VersionNumber.Trim() == versionNumber.Trim());
                                        if (release != null && release.ReleaseId.HasValue)
                                        {
                                            remoteTask.ReleaseId = release.ReleaseId.Value;
                                        }
                                    }
                                }
                                else
                                {
                                    //Make sure that we do any necessary type conversion
                                    //Make sure the field handles nullable types
                                    if (dataValue == null)
                                    {
                                        if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                        {
                                            propertyInfo.SetValue(remoteTask, null, null);
                                        }
                                    }
                                    else
                                    {
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                //Need to handle large string issue
                                                SafeSetStringValue(propertyInfo, remoteTask, MakeXmlSafe(dataValue));

                                            }
                                            else
                                            {
                                                propertyInfo.SetValue(remoteTask, dataValue.ToString(), null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(bool))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                bool flagValue = (MakeXmlSafe(dataValue) == "Y");
                                                propertyInfo.SetValue(remoteTask, flagValue, null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(Nullable<DateTime>))
                                        {
                                            if (dataValue.GetType() == typeof(DateTime))
                                            {
                                                DateTime dateTimeValue = (DateTime)dataValue;
                                                propertyInfo.SetValue(remoteTask, dateTimeValue.ToUniversalTime(), null);
                                            }
                                            else if (dataValue.GetType() == typeof(double))
                                            {
                                                DateTime dateTimeValue = DateTime.FromOADate((double)dataValue).ToUniversalTime();
                                                propertyInfo.SetValue(remoteTask, dateTimeValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                DateTime dateTimeValue;
                                                if (DateTime.TryParse(stringValue, out dateTimeValue))
                                                {
                                                    propertyInfo.SetValue(remoteTask, dateTimeValue.ToUniversalTime(), null);
                                                }
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                        {
                                            if (dataValue.GetType() == typeof(int))
                                            {
                                                propertyInfo.SetValue(remoteTask, dataValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                int intValue;
                                                if (Int32.TryParse(stringValue, out intValue))
                                                {
                                                    propertyInfo.SetValue(remoteTask, intValue, null);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ExportCustomProperties(remoteTask, customProperties, dataValues, customPropertyMapping);

                    if (lastRecord)
                    {
                        break;
                    }

                    //Now either insert or update the task
                    if (remoteTask.TaskId.HasValue)
                    {
                        spiraImportExport.Task_Update(remoteTask);
                    }
                    else
                    {
                        //Insert case
                        remoteTask = spiraImportExport.Task_Create(remoteTask);

                        //Update the cell with the task ID to prevent multiple-appends
                        Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                        newKeyCell.Value2 = remoteTask.TaskId;
                    }

                    //Add a comment if necessary
                    if (newComment != "")
                    {
                        SpiraImportExport.RemoteComment remoteComment = new SpiraImportExport.RemoteComment();
                        remoteComment.ArtifactId = remoteTask.TaskId.Value;
                        remoteComment.Text = newComment;
                        remoteComment.CreationDate = DateTime.UtcNow;
                        spiraImportExport.Task_CreateComment(remoteComment);
                    }

                    //Move to the next row and update progress bar
                    exportCount++;
                    this.UpdateProgress(exportCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Imports custom property list values
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportCustomValues(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Retrieve all the custom lists and their values in the project
            RemoteCustomList[] remoteCustomLists = spiraImportExport.CustomProperty_RetrieveCustomLists();
            List<RemoteCustomListValue> remoteCustomValues = new List<RemoteCustomListValue>();
            //Get the values for each list
            foreach (RemoteCustomList remoteCustomList in remoteCustomLists)
            {
                int remoteCustomListId = remoteCustomList.CustomPropertyListId.Value;
                RemoteCustomList remoteCustomListWithValues = spiraImportExport.CustomProperty_RetrieveCustomListById(remoteCustomListId);
                if (remoteCustomListWithValues.Values != null)
                {
                    remoteCustomValues.AddRange(remoteCustomListWithValues.Values);
                }
            }
            int artifactCount = remoteCustomValues.Count;

            //Set the progress bar accordingly
            this.UpdateProgress(0, artifactCount);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Value #", "CustomPropertyValueId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Custom List #", "CustomPropertyListId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Custom Value Name", "Name", columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("CustomPropertyValueId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Value #'");
            }
            if (!fieldColumnMapping.ContainsKey("CustomPropertyListId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Custom List #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Custom Value Name'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping);

            //Now iterate through the tasks and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            foreach (RemoteCustomListValue remoteCustomValue in remoteCustomValues)
            {
                try
                {
                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated;
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //See if this field exists on the remote object
                        Type remoteObjectType = remoteCustomValue.GetType();
                        PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                        if (propertyInfo != null && propertyInfo.CanRead)
                        {
                            object propertyValue = propertyInfo.GetValue(remoteCustomValue, null);
                            if (propertyInfo.PropertyType == typeof(string))
                            {
                                //For strings we need to verify length and truncate if necessary
                                dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                            }
                            else
                            {
                                dataValues[1, columnIndex] = propertyValue;
                            }
                        }
                    }

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                    //Move to the next row and update progress bar
                    rowIndex++;
                    importCount++;
                    this.UpdateProgress(importCount, null);
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Exports custom values into Spira
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportCustomValues(SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Value #", "CustomPropertyValueId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Custom List #", "CustomPropertyListId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Custom Value Name", "Name", columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("CustomPropertyValueId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Value #'");
            }
            if (!fieldColumnMapping.ContainsKey("CustomPropertyListId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Custom List #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Custom Value Name'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping);

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //Now get all the custom lists in the project
            RemoteCustomList[] remoteCustomLists = spiraImportExport.CustomProperty_RetrieveCustomLists();
            Dictionary<int, RemoteCustomList> remoteCustomListValues = new Dictionary<int,RemoteCustomList>();
            //Get the values for each list
            foreach (RemoteCustomList remoteCustomList in remoteCustomLists)
            {
                int remoteCustomListId = remoteCustomList.CustomPropertyListId.Value;
                RemoteCustomList remoteCustomListWithValues = spiraImportExport.CustomProperty_RetrieveCustomListById(remoteCustomListId);
                if (!remoteCustomListValues.ContainsKey(remoteCustomListId))
                {
                    remoteCustomListValues.Add(remoteCustomListId, remoteCustomListWithValues);
                }
            }

            //Track which lists have changed
            List<int> changedLists = new List<int>();
            Dictionary<int, int> listRowIndex = new Dictionary<int, int>();

            //Now iterate through the rows in the sheet that have data
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //See if we are inserting a new task or updating an existing one
                    int primaryKeyColumn = fieldColumnMapping["CustomPropertyValueId"];
                    RemoteCustomListValue remoteCustomListValue = null;
                    if (dataValues[1, primaryKeyColumn] == null)
                    {
                        //We have insert case
                        remoteCustomListValue = new RemoteCustomListValue();
                    }
                    else
                    {
                        //We have update case
                        string customValueIdString = dataValues[1, primaryKeyColumn].ToString();
                        int customValueId;
                        if (!Int32.TryParse(customValueIdString, out customValueId))
                        {
                            throw new ApplicationException("Custom Value # '" + customValueIdString + "' is not valid. It needs to be a purely integer value.");
                        }
                        foreach (KeyValuePair<int, RemoteCustomList> kvp in remoteCustomListValues)
                        {
                            int customListId = kvp.Key;
                            if (kvp.Value.Values != null)
                            {
                                RemoteCustomListValue matchingValue = kvp.Value.Values.FirstOrDefault(v => v.CustomPropertyValueId == customValueId);
                                if (matchingValue != null)
                                {
                                    remoteCustomListValue = matchingValue;
                                    //Add to the list of changed lists
                                    if (!changedLists.Contains(customListId))
                                    {
                                        changedLists.Add(customListId);
                                    }
                                    if (!listRowIndex.ContainsKey(customListId))
                                    {
                                        listRowIndex.Add(customListId, rowIndex);
                                    }
                                }
                            }
                        }
                    }

                    //Make sure we either are adding new or have a match
                    if (remoteCustomListValue == null)
                    {
                        lastRecord = true;
                        break;
                    }
                    else
                    {
                        //Iterate through the various mapped fields
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //If we have the name/list id field, need to use that to know when we've reached the end of the import
                            if (fieldName == "Name" || fieldName == "CustomPropertyListId")
                            {
                                Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                                if (nameCell == null || nameCell.Value2 == null)
                                {
                                    lastRecord = true;
                                    break;
                                }
                            }

                            //See if this field exists on the remote object 
                            Type remoteObjectType = remoteCustomListValue.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];
                                //Make sure that we do any necessary type conversion
                                //Make sure the field handles nullable types
                                if (dataValue == null)
                                {
                                    if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                    {
                                        propertyInfo.SetValue(remoteCustomListValue, null, null);
                                    }
                                }
                                else
                                {
                                    if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        if (dataValue.GetType() == typeof(string))
                                        {
                                            //Need to handle large string issue
                                            SafeSetStringValue(propertyInfo, remoteCustomListValue, MakeXmlSafe(dataValue));
                                        }
                                        else
                                        {
                                            propertyInfo.SetValue(remoteCustomListValue, dataValue.ToString(), null);
                                        }
                                    }
                                    if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                    {
                                        if (dataValue.GetType() == typeof(int))
                                        {
                                            propertyInfo.SetValue(remoteCustomListValue, dataValue, null);
                                        }
                                        else
                                        {
                                            string stringValue = dataValue.ToString();
                                            int intValue;
                                            if (Int32.TryParse(stringValue, out intValue))
                                            {
                                                propertyInfo.SetValue(remoteCustomListValue, intValue, null);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (lastRecord)
                        {
                            break;
                        }

                        //For the insert case, we perform it immediately, the updates are all done at the end
                        if (!remoteCustomListValue.CustomPropertyValueId.HasValue)
                        {
                            //Insert case
                            remoteCustomListValue = spiraImportExport.CustomProperty_AddCustomListValue(remoteCustomListValue);

                            //Update the cell with the custom value ID to prevent multiple-appends
                            Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                            newKeyCell.Value2 = remoteCustomListValue.CustomPropertyValueId;
                            exportCount++;
                        }
                    }

                    //Move to the next row and update progress bar
                    this.UpdateProgress(exportCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = String.Format ("Error in row #{0}: {1}", rowIndex, exception.Message);
                    errorCount++;
                }
            }

            //Now we do the updates at the end
            foreach (int customListId in changedLists)
            {
                try
                {
                    //Get the object from the dictionary
                    if (remoteCustomListValues.ContainsKey(customListId))
                    {
                        RemoteCustomList remoteCustomList = remoteCustomListValues[customListId];
                        spiraImportExport.CustomProperty_UpdateCustomList(remoteCustomList);
                        if (remoteCustomList.Values != null)
                        {
                            exportCount += remoteCustomList.Values.Length;
                        }
                    }

                    //Move to the next row and update progress bar
                    this.UpdateProgress(exportCount, null);
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    if (listRowIndex.ContainsKey(customListId))
                    {
                        int rowIndex = listRowIndex[customListId];
                        Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                        errorCell.Value2 = String.Format("Error updating custom list CL{0}: {1}.", customListId , exception.Message);
                    }
                    errorCount++;
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Exports incidents
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportIncidents(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Inc_Type");
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Inc_Status");
            Dictionary<int, string> priorityMapping = LoadLookup(importState.LookupWorksheet, "Inc_Priority");
            Dictionary<int, string> severityMapping = LoadLookup(importState.LookupWorksheet, "Inc_Severity");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Incident, false);

            //Get the list of components currently in this project
            RemoteComponent[] components = spiraImportExport.Component_Retrieve(true, false);

            //Get the list of releases currently in this project
            RemoteRelease[] releases = spiraImportExport.Release_Retrieve(true);

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Inc #", "IncidentId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Incident Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Incident Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "IncidentTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "IncidentStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Priority", "PriorityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Severity", "SeverityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Detected Release", "DetectedReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Resolved Release", "ResolvedReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Detector", "OpenerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Est. Effort", "EstimatedEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Act. Effort", "ActualEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Rem. Effort", "RemainingEffort", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Detected Date", "CreationDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Closed Date", "ClosedDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Components", "ComponentIds", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("IncidentId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident Name'");
            }
            if (!fieldColumnMapping.ContainsKey("Description"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident Description'");
            }
            if (!fieldColumnMapping.ContainsKey("IncidentStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("IncidentTypeId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Type'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //Now iterate through the rows in the sheet that have data
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //See if we are inserting a new incident or updating an existing one
                    int primaryKeyColumn = fieldColumnMapping["IncidentId"];
                    SpiraImportExport.RemoteIncident remoteIncident = null;
                    if (dataValues[1, primaryKeyColumn] == null)
                    {
                        //We have insert case
                        remoteIncident = new SpiraExcelAddIn.SpiraImportExport.RemoteIncident();
                    }
                    else
                    {
                        //We have update case
                        string incidentIdString = dataValues[1, primaryKeyColumn].ToString();
                        int incidentId;
                        if (!Int32.TryParse(incidentIdString, out incidentId))
                        {
                            throw new ApplicationException("Incident ID '" + incidentIdString + "' is not valid. It needs to be a purely integer value.");
                        }
                        remoteIncident = spiraImportExport.Incident_RetrieveById(incidentId);
                    }

                    //Iterate through the various mapped fields
                    string incidentResolution = "";
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //If we have the name field, need to use that to know when we've reached the end of the import
                        if (fieldName == "Name")
                        {
                            Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                            if (nameCell == null || nameCell.Value2 == null)
                            {
                                lastRecord = true;
                                break;
                            }
                        }

                        //See if this field exists on the remote object (except Comment which is handled separately)
                        if (fieldName == "Comment")
                        {
                            object dataValue = dataValues[1, columnIndex];
                            if (dataValue != null)
                            {
                                incidentResolution = (MakeXmlSafe(dataValue)).Trim();
                            }
                        }
                        else
                        {
                            Type remoteObjectType = remoteIncident.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];

                                //See if we have one of the special known lookups
                                if (fieldName == "IncidentStatusId")
                                {
                                    string lookupValue = MakeXmlSafe(dataValue);
                                    Nullable<int> fieldValue = null; //Default value for project
                                    foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                    {
                                        if (mappingEntry.Value == lookupValue)
                                        {
                                            fieldValue = mappingEntry.Key;
                                            break;
                                        }
                                    }
                                    if (fieldValue != -1)
                                    {
                                        propertyInfo.SetValue(remoteIncident, fieldValue, null);
                                    }
                                }
                                if (fieldName == "IncidentTypeId")
                                {
                                    string lookupValue = MakeXmlSafe(dataValue);
                                    Nullable<int> fieldValue = null; //Default value for project
                                    foreach (KeyValuePair<int, string> mappingEntry in typeMapping)
                                    {
                                        if (mappingEntry.Value == lookupValue)
                                        {
                                            fieldValue = mappingEntry.Key;
                                            break;
                                        }
                                    }
                                    if (fieldValue != -1)
                                    {
                                        propertyInfo.SetValue(remoteIncident, fieldValue, null);
                                    }
                                }
                                else if (fieldName == "PriorityId")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteIncident, null, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = -1;
                                        foreach (KeyValuePair<int, string> mappingEntry in priorityMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteIncident, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "SeverityId")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteIncident, null, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = -1;
                                        foreach (KeyValuePair<int, string> mappingEntry in severityMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteIncident, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "DetectedReleaseVersionNumber")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteIncident, null, null);
                                    }
                                    else if (dataValue is String)
                                    {
                                        //We need to get the version number and find the corresponding ReleaseId, if it exists
                                        string versionNumber = (string)dataValue;
                                        RemoteRelease release = releases.FirstOrDefault(r => r.VersionNumber.Trim() == versionNumber.Trim());
                                        if (release != null && release.ReleaseId.HasValue)
                                        {
                                            remoteIncident.DetectedReleaseId = release.ReleaseId.Value;
                                        }
                                    }
                                }
                                else if (fieldName == "ResolvedReleaseVersionNumber")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteIncident, null, null);
                                    }
                                    else if (dataValue is String)
                                    {
                                        //We need to get the version number and find the corresponding ReleaseId, if it exists
                                        string versionNumber = (string)dataValue;
                                        RemoteRelease release = releases.FirstOrDefault(r => r.VersionNumber.Trim() == versionNumber.Trim());
                                        if (release != null && release.ReleaseId.HasValue)
                                        {
                                            remoteIncident.ResolvedReleaseId = release.ReleaseId.Value;
                                        }
                                    }
                                }
                                else if (fieldName == "ComponentIds")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteIncident, new int[0], null);
                                    }
                                    else if (dataValue is String)
                                    {
                                        //We need to get the list of component names
                                        string[] componentNames = ((string)dataValue).Split(',');

                                        //Convert to component IDs
                                        List<int> componentIds = new List<int>();
                                        foreach (string componentName in componentNames)
                                        {
                                            RemoteComponent component = components.FirstOrDefault(c => c.Name.Trim() == componentName.Trim());
                                            if (component != null && component.ComponentId.HasValue)
                                            {
                                                componentIds.Add(component.ComponentId.Value);
                                            }
                                        }
                                        if (componentIds.Count > 0)
                                        {
                                            remoteIncident.ComponentIds = componentIds.ToArray();
                                        }
                                    }
                                }
                                else
                                {
                                    //Make sure that we do any necessary type conversion
                                    //Make sure the field handles nullable types
                                    if (dataValue == null)
                                    {
                                        if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                        {
                                            propertyInfo.SetValue(remoteIncident, null, null);
                                        }
                                    }
                                    else
                                    {
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                //Need to handle large string issue
                                                SafeSetStringValue(propertyInfo, remoteIncident, MakeXmlSafe(dataValue));

                                            }
                                            else
                                            {
                                                propertyInfo.SetValue(remoteIncident, dataValue.ToString(), null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(bool))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                bool flagValue = (MakeXmlSafe(dataValue) == "Y");
                                                propertyInfo.SetValue(remoteIncident, flagValue, null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(Nullable<DateTime>))
                                        {
                                            if (dataValue.GetType() == typeof(DateTime))
                                            {
                                                DateTime dateTimeValue = (DateTime)dataValue;
                                                propertyInfo.SetValue(remoteIncident, dateTimeValue.ToUniversalTime(), null);
                                            }
                                            else if (dataValue.GetType() == typeof(double))
                                            {
                                                DateTime dateTimeValue = DateTime.FromOADate((double)dataValue).ToUniversalTime();
                                                propertyInfo.SetValue(remoteIncident, dateTimeValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                DateTime dateTimeValue;
                                                if (DateTime.TryParse(stringValue, out dateTimeValue))
                                                {
                                                    propertyInfo.SetValue(remoteIncident, dateTimeValue.ToUniversalTime(), null);
                                                }
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                        {
                                            if (dataValue.GetType() == typeof(int))
                                            {
                                                propertyInfo.SetValue(remoteIncident, dataValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                int intValue;
                                                if (Int32.TryParse(stringValue, out intValue))
                                                {
                                                    propertyInfo.SetValue(remoteIncident, intValue, null);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ExportCustomProperties(remoteIncident, customProperties, dataValues, customPropertyMapping);

                    if (lastRecord)
                    {
                        break;
                    }

                    //Now either insert or update the incident
                    if (remoteIncident.IncidentId.HasValue)
                    {
                        spiraImportExport.Incident_Update(remoteIncident);
                    }
                    else
                    {
                        //Insert case
                        remoteIncident = spiraImportExport.Incident_Create(remoteIncident);

                        //Update the cell with the incident ID to prevent multiple-appends
                        Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                        newKeyCell.Value2 = remoteIncident.IncidentId;
                    }
                    //Add a resolution if necessary
                    if (incidentResolution != "")
                    {
                        SpiraImportExport.RemoteComment remoteComment = new SpiraImportExport.RemoteComment();
                        remoteComment.ArtifactId = remoteIncident.IncidentId.Value;
                        remoteComment.Text = incidentResolution;
                        remoteComment.CreationDate = DateTime.UtcNow;
                        SpiraImportExport.RemoteComment[] remoteComments = new SpiraExcelAddIn.SpiraImportExport.RemoteComment[] { remoteComment };
                        spiraImportExport.Incident_AddComments(remoteComments);
                    }

                    //Move to the next row and update progress bar
                    exportCount++;
                    this.UpdateProgress(exportCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Gets the detailed validation messages from Spira
        /// </summary>
        /// <param name="validationException"></param>
        /// <returns></returns>
        private string GetValidationFaultDetail(FaultException<ValidationFaultMessage> validationException)
        {
            string message = "";
            ValidationFaultMessage validationFaultMessage = validationException.Detail;
            message = validationFaultMessage.Summary + ": \n";
            {
                foreach (ValidationFaultMessageItem messageItem in validationFaultMessage.Messages)
                {
                    message += messageItem.FieldName + "=" + messageItem.Message + " \n";
                }
            }

            return message;
        }

        /// <summary>
        /// Exports test sets
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportTestSets(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "TestSet_Status");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.TestSet, false);

            //Get the list of releases currently in this project
            RemoteRelease[] releases = spiraImportExport.Release_Retrieve(true);

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "TX #", "TestSetId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Folder", "Folder", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Version", "ReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "TestSetStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Creator", "CreatorId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Planned Date", "PlannedDate", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TestSetId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'TX #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Set Name'");
            }
            if (!fieldColumnMapping.ContainsKey("Folder"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Folder'");
            }
            if (!fieldColumnMapping.ContainsKey("TestSetStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //Now iterate through the rows in the sheet that have data
            Dictionary<int, int> parentPrimaryKeys = new Dictionary<int, int>();
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //First see if we have a test set or folder
                    int folderColumnIndex = fieldColumnMapping["Folder"];
                    object dataValue2 = dataValues[1, folderColumnIndex];
                    string folderValue = (dataValue2 == null) ? "" : dataValue2.ToString();
                    bool isTestSetFolder = (folderValue == "Y");

                    if (isTestSetFolder)
                    {
                        SpiraImportExport.RemoteTestSetFolder remoteTestSetFolder = null;

                        //See if we are inserting a new test case or updating an existing one
                        int primaryKeyColumn = fieldColumnMapping["TestSetId"];
                        if (dataValues[1, primaryKeyColumn] == null)
                        {
                            //We have insert case
                            remoteTestSetFolder = new SpiraExcelAddIn.SpiraImportExport.RemoteTestSetFolder();
                            remoteTestSetFolder.LastUpdateDate = DateTime.UtcNow;
                            remoteTestSetFolder.CreationDate = DateTime.UtcNow;
                        }
                        else
                        {
                            //We have update case
                            string testSetFolderIdString = dataValues[1, primaryKeyColumn].ToString();
                            int testSetFolderId;
                            if (!Int32.TryParse(testSetFolderIdString, out testSetFolderId))
                            {
                                throw new ApplicationException("Test Set Folder ID '" + testSetFolderIdString + "' is not valid. It needs to be a purely integer value.");
                            }
                            remoteTestSetFolder = spiraImportExport.TestSet_RetrieveFolderById(testSetFolderId);
                        }

                        //Loop through the fields though only the name and description field are used
                        int indentLevel = 0;
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //If we have the name field, need to use that to determine the indent
                            //and also to know when we've reached the end of the import
                            if (fieldName == "Name")
                            {
                                Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                                if (nameCell == null || nameCell.Value2 == null)
                                {
                                    lastRecord = true;
                                    break;
                                }
                                else
                                {
                                    indentLevel = (int)nameCell.IndentLevel;
                                }
                            }

                            //See if this field exists on the remote object (except some specific fields which are internal to the sheet)
                            Type remoteObjectType = remoteTestSetFolder.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];

                                //Make sure that we do any necessary type conversion
                                //Make sure the field handles nullable types
                                if (dataValue == null)
                                {
                                    if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                    {
                                        propertyInfo.SetValue(remoteTestSetFolder, null, null);
                                    }
                                }
                                else
                                {
                                    //Test case folders only have string properties (name/description)
                                    if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        if (dataValue.GetType() == typeof(string))
                                        {
                                            //Need to handle large string issue
                                            SafeSetStringValue(propertyInfo, remoteTestSetFolder, MakeXmlSafe(dataValue));
                                        }
                                        else
                                        {
                                            propertyInfo.SetValue(remoteTestSetFolder, dataValue.ToString(), null);
                                        }
                                    }
                                }
                            }
                        }

                        //Now either insert or update the test case folder
                        if (remoteTestSetFolder.TestSetFolderId.HasValue)
                        {
                            //Update case
                            spiraImportExport.TestSet_UpdateFolder(remoteTestSetFolder);
                        }
                        else
                        {
                            //Insert case

                            //If we have an item already loaded that is a parent of this, then set the parent folder id
                            if (parentPrimaryKeys.ContainsKey(indentLevel - 1))
                            {
                                remoteTestSetFolder.ParentTestSetFolderId = parentPrimaryKeys[indentLevel - 1];
                            }
                            else
                            {
                                remoteTestSetFolder.ParentTestSetFolderId = null;
                            }

                            remoteTestSetFolder = spiraImportExport.TestSet_CreateFolder(remoteTestSetFolder);
                            //Update the cell with the test case ID to prevent multiple-appends
                            Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                            newKeyCell.Value2 = remoteTestSetFolder.TestSetFolderId;
                        }

                        //Add to the parent indent dictionary (only folders)
                        if (remoteTestSetFolder.TestSetFolderId.HasValue)
                        {
                            if (!parentPrimaryKeys.ContainsKey(indentLevel))
                            {
                                parentPrimaryKeys.Add(indentLevel, remoteTestSetFolder.TestSetFolderId.Value);
                            }
                            else
                            {
                                parentPrimaryKeys[indentLevel] = remoteTestSetFolder.TestSetFolderId.Value;
                            }
                        }
                    }
                    else
                    {
                        //See if we are inserting a new test set or updating an existing one
                        int primaryKeyColumn = fieldColumnMapping["TestSetId"];
                        SpiraImportExport.RemoteTestSet remoteTestSet = null;
                        if (dataValues[1, primaryKeyColumn] == null)
                        {
                            //We have insert case
                            remoteTestSet = new SpiraExcelAddIn.SpiraImportExport.RemoteTestSet();
                            remoteTestSet.TestSetStatusId = 1;  //Not Started
                            remoteTestSet.TestRunTypeId = 1;    //Manual
                        }
                        else
                        {
                            //We have update case
                            string testSetIdString = dataValues[1, primaryKeyColumn].ToString();
                            int testSetId;
                            if (!Int32.TryParse(testSetIdString, out testSetId))
                            {
                                throw new ApplicationException("Test Set ID '" + testSetIdString + "' is not valid. It needs to be a purely integer value.");
                            }
                            remoteTestSet = spiraImportExport.TestSet_RetrieveById(testSetId);
                        }

                        //Iterate through the various mapped fields
                        int indentLevel = 0;
                        string newComment = "";
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //If we have the name field, need to use that to determine the indent
                            //and also to know when we've reached the end of the import
                            if (fieldName == "Name")
                            {
                                Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                                if (nameCell == null || nameCell.Value2 == null)
                                {
                                    lastRecord = true;
                                    break;
                                }
                                else
                                {
                                    indentLevel = (int)nameCell.IndentLevel;
                                }
                            }

                            //See if this field exists on the remote object (except Comment which is handled separately)
                            if (fieldName == "Comment")
                            {
                                object dataValue = dataValues[1, columnIndex];
                                if (dataValue != null)
                                {
                                    newComment = (MakeXmlSafe(dataValue)).Trim();
                                }
                            }
                            else
                            {
                                Type remoteObjectType = remoteTestSet.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanWrite)
                                {
                                    object dataValue = dataValues[1, columnIndex];

                                    //See if we have one of the special known lookups
                                    if (fieldName == "TestSetStatusId")
                                    {
                                        if (dataValue == null)
                                        {
                                            //This field is not nullable, so we need to pass 1 to default to 'Not Started'
                                            propertyInfo.SetValue(remoteTestSet, 1, null);
                                        }
                                        else
                                        {
                                            string lookupValue = MakeXmlSafe(dataValue);
                                            int fieldValue = 1; //Default to Not Started
                                            foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                            {
                                                if (mappingEntry.Value == lookupValue)
                                                {
                                                    fieldValue = mappingEntry.Key;
                                                    break;
                                                }
                                            }
                                            if (fieldValue != -1)
                                            {
                                                propertyInfo.SetValue(remoteTestSet, fieldValue, null);
                                            }
                                        }
                                    }
                                    else if (fieldName == "ReleaseVersionNumber")
                                    {
                                        if (dataValue == null)
                                        {
                                            propertyInfo.SetValue(remoteTestSet, null, null);
                                        }
                                        else if (dataValue is String)
                                        {
                                            //We need to get the version number and find the corresponding ReleaseId, if it exists
                                            string versionNumber = (string)dataValue;
                                            RemoteRelease release = releases.FirstOrDefault(r => r.VersionNumber.Trim() == versionNumber.Trim());
                                            if (release != null && release.ReleaseId.HasValue)
                                            {
                                                remoteTestSet.ReleaseId = release.ReleaseId.Value;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Make sure that we do any necessary type conversion
                                        //Make sure the field handles nullable types
                                        if (dataValue == null)
                                        {
                                            if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                            {
                                                propertyInfo.SetValue(remoteTestSet, null, null);
                                            }
                                        }
                                        else
                                        {
                                            if (propertyInfo.PropertyType == typeof(string))
                                            {
                                                if (dataValue.GetType() == typeof(string))
                                                {
                                                    //Need to handle large string issue
                                                    SafeSetStringValue(propertyInfo, remoteTestSet, MakeXmlSafe(dataValue));

                                                }
                                                else
                                                {
                                                    propertyInfo.SetValue(remoteTestSet, dataValue.ToString(), null);
                                                }
                                            }
                                            if (propertyInfo.PropertyType == typeof(bool))
                                            {
                                                if (dataValue.GetType() == typeof(string))
                                                {
                                                    bool flagValue = (MakeXmlSafe(dataValue) == "Y");
                                                    propertyInfo.SetValue(remoteTestSet, flagValue, null);
                                                }
                                            }
                                            if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(Nullable<DateTime>))
                                            {
                                                if (dataValue.GetType() == typeof(DateTime))
                                                {
                                                    DateTime dateTimeValue = (DateTime)dataValue;
                                                    propertyInfo.SetValue(remoteTestSet, dateTimeValue.ToUniversalTime(), null);
                                                }
                                                else if (dataValue.GetType() == typeof(double))
                                                {
                                                    DateTime dateTimeValue = DateTime.FromOADate((double)dataValue).ToUniversalTime();
                                                    propertyInfo.SetValue(remoteTestSet, dateTimeValue, null);
                                                }
                                                else
                                                {
                                                    string stringValue = dataValue.ToString();
                                                    DateTime dateTimeValue;
                                                    if (DateTime.TryParse(stringValue, out dateTimeValue))
                                                    {
                                                        propertyInfo.SetValue(remoteTestSet, dateTimeValue.ToUniversalTime(), null);
                                                    }
                                                }
                                            }
                                            if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                            {
                                                if (dataValue.GetType() == typeof(int))
                                                {
                                                    propertyInfo.SetValue(remoteTestSet, dataValue, null);
                                                }
                                                else
                                                {
                                                    string stringValue = dataValue.ToString();
                                                    int intValue;
                                                    if (Int32.TryParse(stringValue, out intValue))
                                                    {
                                                        propertyInfo.SetValue(remoteTestSet, intValue, null);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //Iterate through all the custom properties
                        ExportCustomProperties(remoteTestSet, customProperties, dataValues, customPropertyMapping);

                        if (lastRecord)
                        {
                            break;
                        }

                        //Now either insert or update the test set
                        if (remoteTestSet.TestSetId.HasValue)
                        {
                            spiraImportExport.TestSet_Update(remoteTestSet);
                        }
                        else
                        {
                            //Insert case

                            //If we have an item already loaded that is a parent of this, then specify the folder id
                            if (parentPrimaryKeys.ContainsKey(indentLevel - 1))
                            {
                                remoteTestSet.TestSetFolderId = parentPrimaryKeys[indentLevel - 1];
                            }
                            else
                            {
                                remoteTestSet.TestSetFolderId = null;
                            }
                            remoteTestSet = spiraImportExport.TestSet_Create(remoteTestSet);
                            
                            //Update the cell with the test set ID to prevent multiple-appends
                            Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                            newKeyCell.Value2 = remoteTestSet.TestSetId;
                        }

                        //Add a comment if necessary
                        if (newComment != "")
                        {
                            SpiraImportExport.RemoteComment remoteComment = new SpiraImportExport.RemoteComment();
                            remoteComment.ArtifactId = remoteTestSet.TestSetId.Value;
                            remoteComment.Text = newComment;
                            remoteComment.CreationDate = DateTime.UtcNow;
                            spiraImportExport.TestSet_CreateComment(remoteComment);
                        }
                    }

                    //Move to the next row and update progress bar
                    exportCount++;
                    this.UpdateProgress(exportCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Exports test cases
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportTestCases(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> priorityMapping = LoadLookup(importState.LookupWorksheet, "TestCase_Priority");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "TestCase_Type");
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "TestCase_Status");

            //Get the test case and test step custom property definitions for the current project
            RemoteCustomProperty[] testCaseCustomProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.TestCase, false);
            RemoteCustomProperty[] testStepCustomProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.TestStep, false);

            //Get the list of releases currently in this project
            RemoteRelease[] releases = spiraImportExport.Release_Retrieve(true);

            //Get the list of components currently in this project
            RemoteComponent[] components = spiraImportExport.Component_Retrieve(true, false);

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> testCaseCustomPropertyMapping = new Dictionary<int, int>();
            Dictionary<int, int> testStepCustomPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Test #", "TestCaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Case Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Case Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Priority", "TestCasePriorityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Row Type", "Type", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release", "ReleaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement", "RequirementId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set", "TestSetId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Components", "ComponentIds", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "TestCaseTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "TestCaseStatusId", columnIndex);
                //Test Step Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Step #", "TestStepId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Step Description", "TestStepDescription", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Expected Result", "ExpectedResult", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Sample Data", "SampleData", columnIndex);

                //TestCase Custom Properties
                CheckForCustomPropHeaderCells(testCaseCustomPropertyMapping, testCaseCustomProperties, cell, columnIndex);
                //TestStep Custom Properties
                CheckForCustomPropHeaderCells(testStepCustomPropertyMapping, testStepCustomProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TestCaseId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Case Name'");
            }
            if (!fieldColumnMapping.ContainsKey("Type"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Row Type'");
            }
            if (!fieldColumnMapping.ContainsKey("TestStepId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Step #'");
            }
            if (!fieldColumnMapping.ContainsKey("TestStepDescription"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Step Description'");
            }
            if (!fieldColumnMapping.ContainsKey("ExpectedResult"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Expected Result'");
            }

            //The error column is the column after the last data column
            int errorColumn1 = GetErrorColumn(fieldColumnMapping, testCaseCustomPropertyMapping);
            int errorColumn2 = GetErrorColumn(fieldColumnMapping, testStepCustomPropertyMapping);
            int errorColumn = (errorColumn1 > errorColumn2) ? errorColumn1 : errorColumn2;

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //Now iterate through the rows in the sheet that have data
            Dictionary<int, int> parentPrimaryKeys = new Dictionary<int, int>();
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            SpiraImportExport.RemoteTestCase remoteTestCase = null;
            SpiraImportExport.RemoteTestCaseFolder remoteTestFolder = null;
            int lastPosition = 1;
            bool lastTestCaseWasInsert = true;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //First see if we have a test case, test folder or test step
                    int typeColumnIndex = fieldColumnMapping["Type"];
                    object dataValue2 = dataValues[1, typeColumnIndex];
                    if (dataValue2 == null || dataValue2.ToString() == "")
                    {
                        //End the import if the type field is not set
                        lastRecord = true;
                        break;
                    }
                    string typeValue = dataValue2.ToString();
                    bool isTestStep = (typeValue == ">TestStep");
                    bool isTestFolder = (typeValue == "FOLDER");

                    if (isTestStep)
                    {
                        //We need to have a parent test case (not a folder)
                        if (remoteTestCase == null)
                        {
                            //Warn the user that they need a test case row before the test steps
                            throw new ApplicationException("You cannot have a Test Step row that is not associated with a Test Case");
                        }

                        //See if we are inserting a new test step or updating an existing one
                        int primaryKeyColumn = fieldColumnMapping["TestStepId"];
                        SpiraImportExport.RemoteTestStep remoteTestStep = null;
                        if (dataValues[1, primaryKeyColumn] == null)
                        {
                            //We have insert case
                            remoteTestStep = new SpiraExcelAddIn.SpiraImportExport.RemoteTestStep();
                            remoteTestStep.Position = lastPosition;
                        }
                        else
                        {
                            //We have update case
                            string testStepIdString = dataValues[1, primaryKeyColumn].ToString();
                            int testStepId;
                            if (!Int32.TryParse(testStepIdString, out testStepId))
                            {
                                throw new ApplicationException("Test Step ID '" + testStepIdString + "' is not valid. It needs to be a purely integer value.");
                            }
                            //Locate the remote test step
                            foreach (SpiraImportExport.RemoteTestStep testStep in remoteTestCase.TestSteps)
                            {
                                if (testStep.TestStepId == testStepId)
                                {
                                    remoteTestStep = testStep;
                                    break;
                                }
                            }
                            if (remoteTestStep == null)
                            {
                                throw new ApplicationException("Unable to find a Test Step with ID 'TS" + testStepId + "' in the system. Ignoring this row.");
                            }
                        }
                        lastPosition++;

                        //Iterate through the various mapped fields
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //The Test Step description name was changed to avoid conflicting with the test case description
                            if (fieldName == "Description")
                            {
                                fieldName = "TestCaseDescription";
                            }
                            if (fieldName == "TestStepDescription")
                            {
                                fieldName = "Description";
                            }

                            Type remoteObjectType = remoteTestStep.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];

                                //If we have an empty description, end the import
                                if (fieldName == "Description")
                                {
                                    Range descCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                                    if (descCell == null || descCell.Value2 == null || descCell.Value2.ToString() == "")
                                    {
                                        lastRecord = true;
                                        break;
                                    }
                                }

                                //Make sure that we do any necessary type conversion
                                //Make sure the field handles nullable types
                                if (dataValue == null)
                                {
                                    if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                    {
                                        propertyInfo.SetValue(remoteTestStep, null, null);
                                    }
                                }
                                else
                                {
                                    if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        if (dataValue.GetType() == typeof(string))
                                        {
                                            //Need to handle large string issue
                                            SafeSetStringValue(propertyInfo,remoteTestStep, MakeXmlSafe(dataValue));
                                        }
                                        else
                                        {
                                            propertyInfo.SetValue(remoteTestStep, dataValue.ToString(), null);
                                        }
                                    }
                                    if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                    {
                                        if (dataValue.GetType() == typeof(int))
                                        {
                                            propertyInfo.SetValue(remoteTestStep, dataValue, null);
                                        }
                                        else
                                        {
                                            string stringValue = dataValue.ToString();
                                            int intValue;
                                            if (Int32.TryParse(stringValue, out intValue))
                                            {
                                                propertyInfo.SetValue(remoteTestStep, intValue, null);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //Iterate through all the test step custom properties
                        ExportCustomProperties(remoteTestStep, testStepCustomProperties, dataValues, testStepCustomPropertyMapping);

                        //Now either insert or update the test steps
                        if (remoteTestStep.TestStepId.HasValue)
                        {
                            //For the Update Step case we don't do it now, we wait until we get to the next test case entry
                            //so that all steps are updated in one go
                        }
                        else
                        {
                            //Insert Step case
                            remoteTestStep = spiraImportExport.TestCase_AddStep(remoteTestStep, remoteTestCase.TestCaseId.Value);

                            //Update the cell with the test step ID to prevent multiple-appends
                            Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                            newKeyCell.Value2 = remoteTestStep.TestStepId;
                        }
                        exportCount++;
                    }
                    else if (isTestFolder)
                    {
                        //If we have an existing test case object from a previous row, need to first update that
                        if (remoteTestCase != null && remoteTestCase.TestCaseId.HasValue && !lastTestCaseWasInsert)
                        {
                            spiraImportExport.TestCase_Update(remoteTestCase);
                            remoteTestCase = null;
                        }

                        //See if we are inserting a new test case or updating an existing one
                        int primaryKeyColumn = fieldColumnMapping["TestCaseId"];
                        if (dataValues[1, primaryKeyColumn] == null)
                        {
                            //We have insert case
                            remoteTestFolder = new SpiraExcelAddIn.SpiraImportExport.RemoteTestCaseFolder();
                            remoteTestFolder.LastUpdateDate = DateTime.UtcNow;
                        }
                        else
                        {
                            //We have update case
                            string testCaseFolderIdString = dataValues[1, primaryKeyColumn].ToString();
                            int testCaseFolderId;
                            if (!Int32.TryParse(testCaseFolderIdString, out testCaseFolderId))
                            {
                                throw new ApplicationException("Test Case Folder ID '" + testCaseFolderIdString + "' is not valid. It needs to be a purely integer value.");
                            }
                            remoteTestFolder = spiraImportExport.TestCase_RetrieveFolderById(testCaseFolderId);
                        }
                        //Restart test step positions
                        lastPosition = 1;

                        //Loop through the fields though only the name and description field are used
                        int indentLevel = 0;
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //If we have the name field, need to use that to determine the indent
                            //and also to know when we've reached the end of the import
                            if (fieldName == "Name")
                            {
                                Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                                if (nameCell == null || nameCell.Value2 == null)
                                {
                                    lastRecord = true;
                                    break;
                                }
                                else
                                {
                                    indentLevel = (int)nameCell.IndentLevel;
                                }
                            }

                            //See if this field exists on the remote object (except some specific fields which are internal to the sheet)
                            if (fieldName == "Type")
                            {
                                object dataValue = dataValues[1, columnIndex];
                                if (dataValue == null || dataValue.ToString() == "")
                                {
                                    lastRecord = true;
                                    break;
                                }
                            }
                            else
                            {
                                Type remoteObjectType = remoteTestFolder.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanWrite)
                                {
                                    object dataValue = dataValues[1, columnIndex];

                                    //Make sure that we do any necessary type conversion
                                    //Make sure the field handles nullable types
                                    if (dataValue == null)
                                    {
                                        if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                        {
                                            propertyInfo.SetValue(remoteTestFolder, null, null);
                                        }
                                    }
                                    else
                                    {
                                        //Test case folders only have string properties (name/description)
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                //Need to handle large string issue
                                                SafeSetStringValue(propertyInfo, remoteTestFolder, MakeXmlSafe(dataValue));
                                            }
                                            else
                                            {
                                                propertyInfo.SetValue(remoteTestFolder, dataValue.ToString(), null);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //Now either insert or update the test case folde
                        if (remoteTestFolder.TestCaseFolderId.HasValue)
                        {
                            //Update case
                            spiraImportExport.TestCase_UpdateFolder(remoteTestFolder);
                        }
                        else
                        {
                            //Insert case

                            //If we have an item already loaded that is a parent of this, then set the parent folder id
                            if (parentPrimaryKeys.ContainsKey(indentLevel - 1))
                            {
                                remoteTestFolder.ParentTestCaseFolderId = parentPrimaryKeys[indentLevel - 1];
                            }
                            else
                            {
                                remoteTestFolder.ParentTestCaseFolderId = null;
                            }

                            remoteTestFolder = spiraImportExport.TestCase_CreateFolder(remoteTestFolder);
                            //Update the cell with the test case ID to prevent multiple-appends
                            Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                            newKeyCell.Value2 = remoteTestFolder.TestCaseFolderId;
                        }

                        //Add to the parent indent dictionary (only folders)
                        if (remoteTestFolder.TestCaseFolderId.HasValue)
                        {
                            if (!parentPrimaryKeys.ContainsKey(indentLevel))
                            {
                                parentPrimaryKeys.Add(indentLevel, remoteTestFolder.TestCaseFolderId.Value);
                            }
                            else
                            {
                                parentPrimaryKeys[indentLevel] = remoteTestFolder.TestCaseFolderId.Value;
                            }
                        }
                    }
                    else
                    {
                        //If we have an existing test case object from a previous row, need to first update that
                        if (remoteTestCase != null && remoteTestCase.TestCaseId.HasValue && !lastTestCaseWasInsert)
                        {
                            spiraImportExport.TestCase_Update(remoteTestCase);
                            remoteTestCase = null;
                        }

                        //See if we are inserting a new test case or updating an existing one
                        int primaryKeyColumn = fieldColumnMapping["TestCaseId"];
                        if (dataValues[1, primaryKeyColumn] == null)
                        {
                            //We have insert case, set the default values for Type and Status in case those fields are not mapped
                            lastTestCaseWasInsert = true;
                            remoteTestCase = new SpiraExcelAddIn.SpiraImportExport.RemoteTestCase();
                            remoteTestCase.TestCaseStatusId = 1;    //Draft
                            remoteTestCase.TestCaseTypeId = 3;  //Functional
                        }
                        else
                        {
                            //We have update case
                            lastTestCaseWasInsert = false;
                            string testCaseIdString = dataValues[1, primaryKeyColumn].ToString();
                            int testCaseId;
                            if (!Int32.TryParse(testCaseIdString, out testCaseId))
                            {
                                throw new ApplicationException("Test Case ID '" + testCaseIdString + "' is not valid. It needs to be a purely integer value.");
                            }
                            remoteTestCase = spiraImportExport.TestCase_RetrieveById(testCaseId);
                        }
                        //Restart test step positions
                        lastPosition = 1;

                        //Iterate through the various mapped fields
                        int indentLevel = 0;
                        List<int> requirementIds = new List<int>();
                        List<int> releaseIds = new List<int>();
                        List<int> testSetIds = new List<int>();
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //If we have the name field, need to use that to determine the indent
                            //and also to know when we've reached the end of the import
                            if (fieldName == "Name")
                            {
                                Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                                if (nameCell == null || nameCell.Value2 == null)
                                {
                                    lastRecord = true;
                                    break;
                                }
                                else
                                {
                                    indentLevel = (int)nameCell.IndentLevel;
                                }
                            }

                            //See if this field exists on the remote object (except some specific fields which are internal to the sheet)
                            if (fieldName == "Type")
                            {
                                object dataValue = dataValues[1, columnIndex];
                                if (dataValue == null || dataValue.ToString() == "")
                                {
                                    lastRecord = true;
                                    break;
                                }
                            }
                            else if (fieldName == "RequirementId")
                            {
                                object dataValue = dataValues[1, columnIndex];
                                if (dataValue != null)
                                {
                                    string lookupIdString = dataValue.ToString();
                                    string[] lookupIdComponents = lookupIdString.Split(',');
                                    foreach (string lookupIdComponent in lookupIdComponents)
                                    {
                                        int lookupId;
                                        if (Int32.TryParse(lookupIdComponent, out lookupId))
                                        {
                                            requirementIds.Add(lookupId);
                                        }
                                    }
                                }
                            }
                            else if (fieldName == "TestSetId")
                            {
                                object dataValue = dataValues[1, columnIndex];
                                if (dataValue != null)
                                {
                                    string lookupIdString = dataValue.ToString();
                                    string[] lookupIdComponents = lookupIdString.Split(',');
                                    foreach (string lookupIdComponent in lookupIdComponents)
                                    {
                                        int lookupId;
                                        if (Int32.TryParse(lookupIdComponent, out lookupId))
                                        {
                                            testSetIds.Add(lookupId);
                                        }
                                    }
                                }
                            }
                            else if (fieldName == "ReleaseId")
                            {
                                object dataValue = dataValues[1, columnIndex];
                                if (dataValue != null)
                                {
                                    string lookupIdString = dataValue.ToString();
                                    string[] lookupIdComponents = lookupIdString.Split(',');
                                    foreach (string lookupIdComponent in lookupIdComponents)
                                    {
                                        //See if we have a release that has this version number
                                        RemoteRelease matchedRelease = releases.FirstOrDefault(r => r.VersionNumber.Trim() == lookupIdComponent.Trim());
                                        if (matchedRelease != null && matchedRelease.ReleaseId.HasValue)
                                        {
                                            releaseIds.Add(matchedRelease.ReleaseId.Value);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Type remoteObjectType = remoteTestCase.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanWrite)
                                {
                                    object dataValue = dataValues[1, columnIndex];

                                    //See if we have one of the special known lookups
                                    if (fieldName == "TestCasePriorityId")
                                    {
                                        if (dataValue == null)
                                        {
                                            //This field is nullable so just set to null
                                            propertyInfo.SetValue(remoteTestCase, null, null);
                                        }
                                        else
                                        {
                                            string lookupValue = MakeXmlSafe(dataValue);
                                            int fieldValue = -1;
                                            foreach (KeyValuePair<int, string> mappingEntry in priorityMapping)
                                            {
                                                if (mappingEntry.Value == lookupValue)
                                                {
                                                    fieldValue = mappingEntry.Key;
                                                    break;
                                                }
                                            }
                                            if (fieldValue != -1)
                                            {
                                                propertyInfo.SetValue(remoteTestCase, fieldValue, null);
                                            }
                                        }
                                    }
                                    else if (fieldName == "TestCaseStatusId")
                                    {
                                        if (dataValue != null)
                                        {
                                            //Field is not nullable so ignore nulls
                                            string lookupValue = MakeXmlSafe(dataValue);
                                            int fieldValue = -1;
                                            foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                            {
                                                if (mappingEntry.Value == lookupValue)
                                                {
                                                    fieldValue = mappingEntry.Key;
                                                    break;
                                                }
                                            }
                                            if (fieldValue != -1)
                                            {
                                                propertyInfo.SetValue(remoteTestCase, fieldValue, null);
                                            }
                                        }
                                    }
                                    else if (fieldName == "TestCaseTypeId")
                                    {
                                        if (dataValue != null)
                                        {
                                            //Field is not nullable so ignore nulls
                                            string lookupValue = MakeXmlSafe(dataValue);
                                            int fieldValue = -1;
                                            foreach (KeyValuePair<int, string> mappingEntry in typeMapping)
                                            {
                                                if (mappingEntry.Value == lookupValue)
                                                {
                                                    fieldValue = mappingEntry.Key;
                                                    break;
                                                }
                                            }
                                            if (fieldValue != -1)
                                            {
                                                propertyInfo.SetValue(remoteTestCase, fieldValue, null);
                                            }
                                        }
                                    }
                                    else if (fieldName == "ComponentIds")
                                    {
                                        if (dataValue == null)
                                        {
                                            propertyInfo.SetValue(remoteTestCase, new int[0], null);
                                        }
                                        else if (dataValue is String)
                                        {
                                            //We need to get the list of component names
                                            string[] componentNames = ((string)dataValue).Split(',');

                                            //Convert to component IDs
                                            List<int> componentIds = new List<int>();
                                            foreach (string componentName in componentNames)
                                            {
                                                RemoteComponent component = components.FirstOrDefault(c => c.Name.Trim() == componentName.Trim());
                                                if (component != null && component.ComponentId.HasValue)
                                                {
                                                    componentIds.Add(component.ComponentId.Value);
                                                }
                                            }
                                            if (componentIds.Count > 0)
                                            {
                                                remoteTestCase.ComponentIds = componentIds.ToArray();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Make sure that we do any necessary type conversion
                                        //Make sure the field handles nullable types
                                        if (dataValue == null)
                                        {
                                            if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                            {
                                                propertyInfo.SetValue(remoteTestCase, null, null);
                                            }
                                        }
                                        else
                                        {
                                            if (propertyInfo.PropertyType == typeof(string))
                                            {
                                                if (dataValue.GetType() == typeof(string))
                                                {
                                                    //Need to handle large string issue
                                                    SafeSetStringValue(propertyInfo, remoteTestCase, MakeXmlSafe(dataValue));
                                                }
                                                else
                                                {
                                                    propertyInfo.SetValue(remoteTestCase, dataValue.ToString(), null);
                                                }
                                            }
                                            if (propertyInfo.PropertyType == typeof(bool))
                                            {
                                                if (dataValue.GetType() == typeof(string))
                                                {
                                                    bool flagValue = (MakeXmlSafe(dataValue) == "Y");
                                                    propertyInfo.SetValue(remoteTestCase, flagValue, null);
                                                }
                                            }
                                            if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(Nullable<DateTime>))
                                            {
                                                if (dataValue.GetType() == typeof(DateTime))
                                                {
                                                    propertyInfo.SetValue(remoteTestCase, dataValue, null);
                                                }
                                                else if (dataValue.GetType() == typeof(double))
                                                {
                                                    DateTime dateTimeValue = DateTime.FromOADate((double)dataValue);
                                                    propertyInfo.SetValue(remoteTestCase, dateTimeValue, null);
                                                }
                                                else
                                                {
                                                    string stringValue = dataValue.ToString();
                                                    DateTime dateTimeValue;
                                                    if (DateTime.TryParse(stringValue, out dateTimeValue))
                                                    {
                                                        propertyInfo.SetValue(remoteTestCase, dateTimeValue, null);
                                                    }
                                                }
                                            }
                                            if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                            {
                                                if (dataValue.GetType() == typeof(int))
                                                {
                                                    propertyInfo.SetValue(remoteTestCase, dataValue, null);
                                                }
                                                else
                                                {
                                                    string stringValue = dataValue.ToString();
                                                    int intValue;
                                                    if (Int32.TryParse(stringValue, out intValue))
                                                    {
                                                        propertyInfo.SetValue(remoteTestCase, intValue, null);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //Iterate through all the test case custom properties
                        ExportCustomProperties(remoteTestCase, testCaseCustomProperties, dataValues, testCaseCustomPropertyMapping);

                        if (lastRecord)
                        {
                            break;
                        }

                        //Now insert the test case, the update case is handled after all test steps are captured
                        if (!remoteTestCase.TestCaseId.HasValue)
                        {
                            //Insert case

                            //If we have an item already loaded that is a parent of this, then set the parent folder id
                            if (parentPrimaryKeys.ContainsKey(indentLevel - 1))
                            {
                                remoteTestCase.TestCaseFolderId = parentPrimaryKeys[indentLevel - 1];
                            }
                            else
                            {
                                remoteTestCase.TestCaseFolderId = null;
                            }
                            remoteTestCase = spiraImportExport.TestCase_Create(remoteTestCase);
                            //Update the cell with the test case ID to prevent multiple-appends
                            Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                            newKeyCell.Value2 = remoteTestCase.TestCaseId;
                        }

                        //Add to the requirement, release or test set if specified
                        if (requirementIds.Count > 0)
                        {
                            foreach (int requirementId in requirementIds)
                            {
                                SpiraImportExport.RemoteRequirementTestCaseMapping remoteMapping = new SpiraExcelAddIn.SpiraImportExport.RemoteRequirementTestCaseMapping();
                                remoteMapping.RequirementId = requirementId;
                                remoteMapping.TestCaseId = remoteTestCase.TestCaseId.Value;
                                spiraImportExport.Requirement_AddTestCoverage(remoteMapping);
                            }
                        }
                        if (releaseIds.Count > 0)
                        {
                            foreach (int releaseId in releaseIds)
                            {
                                SpiraImportExport.RemoteReleaseTestCaseMapping remoteMapping = new SpiraExcelAddIn.SpiraImportExport.RemoteReleaseTestCaseMapping();
                                remoteMapping.ReleaseId = releaseId;
                                remoteMapping.TestCaseId = remoteTestCase.TestCaseId.Value;
                                spiraImportExport.Release_AddTestMapping(remoteMapping);
                            }
                        }
                        if (testSetIds.Count > 0)
                        {
                            foreach (int testSetId in testSetIds)
                            {

                                SpiraImportExport.RemoteTestSetTestCaseMapping remoteMapping = new SpiraExcelAddIn.SpiraImportExport.RemoteTestSetTestCaseMapping();
                                remoteMapping.TestSetId = testSetId;
                                remoteMapping.TestCaseId = remoteTestCase.TestCaseId.Value;
                                spiraImportExport.TestSet_AddTestMapping(remoteMapping, null, null);
                            }
                        }

                        //Move to the next row and update progress bar
                        exportCount++;
                        this.UpdateProgress(exportCount, null);
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }

            //If the last test case was an existing one, do the final update
            if (remoteTestCase != null && remoteTestCase.TestCaseId.HasValue && !lastTestCaseWasInsert)
            {
                //Update the last test case
                spiraImportExport.TestCase_Update(remoteTestCase);
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Imports test sets
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportTestSets(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "TestSet_Status");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.TestSet, false);

            int totalNumberOfTestSets = (int)spiraImportExport.TestSet_Count(null, null);
            this.UpdateProgress(0, totalNumberOfTestSets);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "TX #", "TestSetId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Folder", "Folder", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Version", "ReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "TestSetStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Creator", "CreatorId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Planned Date", "PlannedDate", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TestSetId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'TX #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Set Name'");
            }
            if (!fieldColumnMapping.ContainsKey("Folder"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Folder'");
            }
            if (!fieldColumnMapping.ContainsKey("TestSetStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Now iterate through the test sets and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            int artifactCount = 0;

            //First retrieve all the test set folders in the project
            SpiraImportExport.RemoteTestSetFolder[] remoteTestSetFolders = spiraImportExport.TestSet_RetrieveFolders();
            if (remoteTestSetFolders != null && remoteTestSetFolders.Length > 0)
            {
                //Add to the overall count total
                totalNumberOfTestSets += remoteTestSetFolders.Length;
                this.UpdateProgress(0, totalNumberOfTestSets);

                //Loop through the folders
                foreach (SpiraImportExport.RemoteTestSetFolder remoteTestSetFolder in remoteTestSetFolders)
                {
                    //Import the test folder
                    try
                    {
                        //For performance using VSTO Interop we need to update all the fields in the row in one go
                        Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                        object[,] dataValues = (object[,])dataRange.Value2;

                        //Iterate through the various mapped fields
                        bool oldTruncated = truncated;
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //See if this field exists on the remote object (except type which is internal to the sheet)
                            if (fieldName == "Folder")
                            {
                                dataValues[1, columnIndex] = "Y";
                            }
                            else if (fieldName == "TestSetId")
                            {
                                //This is really the test folder
                                dataValues[1, columnIndex] = remoteTestSetFolder.TestSetFolderId.Value;
                            }
                            else
                            {
                                Type remoteObjectType = remoteTestSetFolder.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanRead)
                                {
                                    //If we have the name field, need to update the indent level
                                    if (fieldName == "Name")
                                    {
                                        Range nameCell = (Range)worksheet.Cells[rowIndex + 2, columnIndex];
                                        nameCell.IndentLevel = (remoteTestSetFolder.IndentLevel.Length / 3) - 1;
                                    }

                                    //See if we have one of the special columns to handle differently
                                    object propertyValue = propertyInfo.GetValue(remoteTestSetFolder, null);
                                    if (fieldName == "Description")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else
                                    {
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            //For strings we need to verify length and truncate if necessary
                                            dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                        }
                                        else
                                        {
                                            dataValues[1, columnIndex] = propertyValue;
                                        }
                                    }
                                }
                            }
                        }

                        //Now commit the data
                        dataRange.Value2 = dataValues;

                        //Since the test set is a folder one, mark field as Bold.
                        dataRange.Font.Bold = true;

                        //If it was truncated on this row, display a message in the right-most column
                        if (truncated && !oldTruncated)
                        {
                            Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                            errorCell.Value2 = "This row had data truncated.";
                        }

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Import aborted by user.");
                        }

                        //Now we need to import any Test Sets
                        rowIndex++;
                        ImportTestSetsInFolder(spiraImportExport, remoteTestSetFolder.TestSetFolderId, worksheet, customProperties, fieldColumnMapping, customPropertyMapping, ref rowIndex, maxColumnIndex, ref truncated, errorColumn, ref errorCount, ref artifactCount, ref importCount, statusMapping, (remoteTestSetFolder.IndentLevel.Length / 3) - 1);

                        //Move to the next row and update progress bar
                        importCount++;
                        this.UpdateProgress(importCount, null);
                    }
                    catch (Exception exception)
                    {
                        //Record the error on the sheet and add to the error count, then continue
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = exception.Message;
                        errorCount++;
                    }
                }
            }

            //Import all the test sets in the root folder
            ImportTestSetsInFolder(spiraImportExport, null, worksheet, customProperties, fieldColumnMapping, customPropertyMapping, ref rowIndex, maxColumnIndex, ref truncated, errorColumn, ref errorCount, ref artifactCount, ref importCount, statusMapping, 0);

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Imports the test cases in the specified folder
        /// </summary>
        private void ImportTestSetsInFolder(SpiraImportExport.SoapServiceClient spiraImportExport, int? testSetFolderId, Worksheet worksheet, RemoteCustomProperty[] customProperties, Dictionary<string, int> fieldColumnMapping, Dictionary<int, int> customPropertyMapping, ref int rowIndex, int maxColumnIndex, ref bool truncated, int errorColumn, ref int errorCount, ref int artifactCount, ref int importCount, Dictionary<int, string> statusMapping, int folderIndentlevel)
        {
            RemoteSort remoteSort = new RemoteSort();
            remoteSort.PropertyName = "Name";
            remoteSort.SortAscending = true;
            SpiraImportExport.RemoteTestSet[] remoteTestSets = spiraImportExport.TestSet_RetrieveByFolder(testSetFolderId, null, remoteSort, 1, Int32.MaxValue, null);
            artifactCount += remoteTestSets.Length;

            //Iterate through the results
            foreach (SpiraImportExport.RemoteTestSet remoteTestSet in remoteTestSets)
            {
                try
                {
                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated;
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //See if this field exists on the remote object (except Folder which is internal to the sheet)
                        if (fieldName == "Folder")
                        {
                            dataValues[1, columnIndex] = "N";
                        }
                        else
                        {
                            //See if this field exists on the remote object
                            Type remoteObjectType = remoteTestSet.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanRead)
                            {
                                //If we have the name field, need to update the indent-level
                                if (fieldName == "Name")
                                {
                                    Range nameCell = (Range)worksheet.Cells[rowIndex + 2, columnIndex];
                                    nameCell.IndentLevel = folderIndentlevel + 1;
                                }
                                else if (fieldName == "Folder")
                                {
                                    dataValues[1, columnIndex] = "N";
                                }

                                //See if we have one of the special known lookups
                                object propertyValue = propertyInfo.GetValue(remoteTestSet, null);
                                if (fieldName == "TestSetStatusId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (statusMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = statusMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "Description")
                                {
                                    //Need to strip off any formatting and make sure it's not too long
                                    dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                }
                                else
                                {
                                    //or if we have to convert data-types
                                    if (propertyInfo.PropertyType == typeof(bool))
                                    {
                                        bool flagValue = (bool)propertyValue;
                                        dataValues[1, columnIndex] = (flagValue) ? "Y" : "N";
                                    }
                                    else if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        //For strings we need to verify length and truncate if necessary
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
                                    {
                                        //For dates we need to convert to local-time
                                        if (propertyValue is DateTime)
                                        {
                                            DateTime dateTimeValue = (DateTime)propertyValue;
                                            dataValues[1, columnIndex] = dateTimeValue.ToLocalTime();
                                        }
                                        else if (propertyValue is DateTime?)
                                        {
                                            DateTime? dateTimeValue = (DateTime?)propertyValue;
                                            if (dateTimeValue.HasValue)
                                            {
                                                dataValues[1, columnIndex] = dateTimeValue.Value.ToLocalTime();
                                            }
                                            else
                                            {
                                                dataValues[1, columnIndex] = null;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dataValues[1, columnIndex] = propertyValue;
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ImportCustomProperties(remoteTestSet, customProperties, dataValues, customPropertyMapping);

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //The test set is not a folder
                    dataRange.Font.Bold = false;

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                    //Move to the next row and update progress bar
                    rowIndex++;
                    importCount++;
                    this.UpdateProgress(importCount, null);
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
        }

        /// <summary>
        /// Imports the test cases in the specified folder
        /// </summary>
        private void ImportTestCasesInFolder(SpiraImportExport.SoapServiceClient spiraImportExport, int? testCaseFolderId, Worksheet worksheet, RemoteCustomProperty[] testCaseCustomProperties, RemoteCustomProperty[] testStepCustomProperties, Dictionary<string, int> fieldColumnMapping, Dictionary<int, int> testCaseCustomPropertyMapping, Dictionary<int, int> testStepCustomPropertyMapping, ref int rowIndex, int maxColumnIndex, ref bool truncated, int errorColumn, ref int errorCount, ref int artifactCount, ref int importCount, Dictionary<int, string> priorityMapping, Dictionary<int, string> statusMapping, Dictionary<int, string> typeMapping, int folderIndentlevel, RemoteComponent[] components)
        {
            RemoteSort remoteSort = new RemoteSort();
            remoteSort.PropertyName = "Name";
            remoteSort.SortAscending = true;
            SpiraImportExport.RemoteTestCase[] remoteTestCases = spiraImportExport.TestCase_RetrieveByFolder(testCaseFolderId, null, remoteSort, 1, Int32.MaxValue, null);
            artifactCount += remoteTestCases.Length;

            //Iterate through the batch
            foreach (SpiraImportExport.RemoteTestCase remoteTestCase in remoteTestCases)
            {
                try
                {
                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated;
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //See if this field exists on the remote object (except type which is internal to the sheet)
                        if (fieldName == "Type")
                        {
                            dataValues[1, columnIndex] = "TestCase";
                        }
                        else
                        {
                            Type remoteObjectType = remoteTestCase.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanRead)
                            {
                                //If we have the name field, need to update the indent level
                                if (fieldName == "Name")
                                {
                                    Range nameCell = (Range)worksheet.Cells[rowIndex + 2, columnIndex];
                                    nameCell.IndentLevel = folderIndentlevel + 1;
                                }

                                //See if we have one of the special known lookups
                                object propertyValue = propertyInfo.GetValue(remoteTestCase, null);
                                if (fieldName == "TestCasePriorityId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (priorityMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = priorityMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "TestCaseStatusId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (statusMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = statusMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "TestCaseTypeId")
                                {
                                    if (propertyValue != null)
                                    {
                                        int fieldValue = (int)propertyValue;
                                        if (typeMapping.ContainsKey(fieldValue))
                                        {
                                            string lookupValue = typeMapping[fieldValue];
                                            dataValues[1, columnIndex] = lookupValue;
                                        }
                                    }
                                }
                                else if (fieldName == "ComponentIds")
                                {
                                    if (propertyValue == null)
                                    {
                                        dataValues[1, columnIndex] = "";
                                    }
                                    else
                                    {
                                        int[] componentIds = (int[])propertyValue;
                                        if (componentIds.Length > 0)
                                        {
                                            string componentNames = "";
                                            foreach (int componentId in componentIds)
                                            {
                                                RemoteComponent component = components.FirstOrDefault(c => c.ComponentId == componentId);
                                                if (component != null)
                                                {
                                                    if (componentNames == "")
                                                    {
                                                        componentNames = component.Name;
                                                    }
                                                    else
                                                    {
                                                        componentNames += "," + component.Name;
                                                    }
                                                }
                                            }
                                            dataValues[1, columnIndex] = componentNames;
                                        }
                                        else
                                        {
                                            dataValues[1, columnIndex] = "";
                                        }
                                    }
                                }
                                else if (fieldName == "Description")
                                {
                                    //Need to strip off any formatting and make sure it's not too long
                                    dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                }
                                else
                                {
                                    //or if we have to convert data-types
                                    if (propertyInfo.PropertyType == typeof(bool))
                                    {
                                        bool flagValue = (bool)propertyValue;
                                        dataValues[1, columnIndex] = (flagValue) ? "Y" : "N";
                                    }
                                    else if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        //For strings we need to verify length and truncate if necessary
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else
                                    {
                                        dataValues[1, columnIndex] = propertyValue;
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ImportCustomProperties(remoteTestCase, testCaseCustomProperties, dataValues, testCaseCustomPropertyMapping);

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //If the test case is a folder one, mark field as Bold.
                    dataRange.Font.Bold = false;

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                    //Now we need to import any Test Steps
                    rowIndex++;
                    ImportTestSteps(spiraImportExport, remoteTestCase.TestCaseId.Value, worksheet, testStepCustomProperties, fieldColumnMapping, testStepCustomPropertyMapping, ref rowIndex, maxColumnIndex, ref truncated, errorColumn, ref errorCount);

                    //Move to the next row and update progress bar
                    importCount++;
                    this.UpdateProgress(importCount, null);
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
        }

        /// <summary>
        /// Imports test cases
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportTestCases(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> priorityMapping = LoadLookup(importState.LookupWorksheet, "TestCase_Priority");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "TestCase_Type");
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "TestCase_Status");

            //Get the test case and test step custom property definitions for the current project
            RemoteCustomProperty[] testCaseCustomProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.TestCase, false);
            RemoteCustomProperty[] testStepCustomProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.TestStep, false);

            //Get the list of components currently in this project
            RemoteComponent[] components = spiraImportExport.Component_Retrieve(true, false);

            //Set the progress bar accordingly
            int totalNumberOfTestCases = (int)spiraImportExport.TestCase_Count(null, null);
            this.UpdateProgress(0, totalNumberOfTestCases);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> testCaseCustomPropertyMapping = new Dictionary<int, int>();
            Dictionary<int, int> testStepCustomPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Test #", "TestCaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Case Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Case Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Priority", "TestCasePriorityId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Row Type", "Type", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Components", "ComponentIds", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "TestCaseTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "TestCaseStatusId", columnIndex);
                //Test Step Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Step #", "TestStepId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Step Description", "TestStepDescription", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Expected Result", "ExpectedResult", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Sample Data", "SampleData", columnIndex);

                //TestCase Custom Properties
                CheckForCustomPropHeaderCells(testCaseCustomPropertyMapping, testCaseCustomProperties, cell, columnIndex);
                //TestStep Custom Properties
                CheckForCustomPropHeaderCells(testStepCustomPropertyMapping, testStepCustomProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TestCaseId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Case Name'");
            }
            if (!fieldColumnMapping.ContainsKey("Type"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Row Type'");
            }
            if (!fieldColumnMapping.ContainsKey("TestStepId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Step #'");
            }
            if (!fieldColumnMapping.ContainsKey("TestStepDescription"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Step Description'");
            }
            if (!fieldColumnMapping.ContainsKey("ExpectedResult"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Expected Result'");
            }

            //The error column is the column after the last data column
            int errorColumn1 = GetErrorColumn(fieldColumnMapping, testCaseCustomPropertyMapping);
            int errorColumn2 = GetErrorColumn(fieldColumnMapping, testStepCustomPropertyMapping);
            int errorColumn = (errorColumn1 > errorColumn2) ? errorColumn1 : errorColumn2;

            //Now iterate through the test cases and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            int artifactCount = 0;

            //First retrieve all the folders in the project
            SpiraImportExport.RemoteTestCaseFolder[] remoteTestFolders = spiraImportExport.TestCase_RetrieveFolders();
            if (remoteTestFolders != null && remoteTestFolders.Length > 0)
            {
                //Add to the overall count total
                totalNumberOfTestCases += remoteTestFolders.Length;
                this.UpdateProgress(0, totalNumberOfTestCases);

                //Loop through the folders
                foreach (SpiraImportExport.RemoteTestCaseFolder remoteTestFolder in remoteTestFolders)
                {
                    //Import the test folder
                    try
                    {
                        //For performance using VSTO Interop we need to update all the fields in the row in one go
                        Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                        object[,] dataValues = (object[,])dataRange.Value2;

                        //Iterate through the various mapped fields
                        bool oldTruncated = truncated;
                        foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                        {
                            int columnIndex = fieldColumnPair.Value;
                            string fieldName = fieldColumnPair.Key;

                            //See if this field exists on the remote object (except type which is internal to the sheet)
                            if (fieldName == "Type")
                            {
                                dataValues[1, columnIndex] = "FOLDER";
                            }
                            else if (fieldName == "TestCaseId")
                            {
                                //This is really the test folder
                                dataValues[1, columnIndex] = remoteTestFolder.TestCaseFolderId.Value;
                            }
                            else
                            {
                                Type remoteObjectType = remoteTestFolder.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanRead)
                                {
                                    //If we have the name field, need to update the indent level
                                    if (fieldName == "Name")
                                    {
                                        Range nameCell = (Range)worksheet.Cells[rowIndex + 2, columnIndex];
                                        nameCell.IndentLevel = (remoteTestFolder.IndentLevel.Length / 3) - 1;
                                    }

                                    //See if we have one of the special columns to handle differently
                                    object propertyValue = propertyInfo.GetValue(remoteTestFolder, null);
                                    if (fieldName == "Description")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else
                                    {
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            //For strings we need to verify length and truncate if necessary
                                            dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                        }
                                        else
                                        {
                                            dataValues[1, columnIndex] = propertyValue;
                                        }
                                    }
                                }
                            }
                        }

                        //Now commit the data
                        dataRange.Value2 = dataValues;

                        //Since the test case is a folder one, mark field as Bold.
                        dataRange.Font.Bold = true;

                        //If it was truncated on this row, display a message in the right-most column
                        if (truncated && !oldTruncated)
                        {
                            Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                            errorCell.Value2 = "This row had data truncated.";
                        }

                        //Check for abort condition
                        if (this.IsAborted)
                        {
                            throw new ApplicationException("Import aborted by user.");
                        }

                        //Now we need to import any Test Cases
                        rowIndex++;
                        ImportTestCasesInFolder(spiraImportExport, remoteTestFolder.TestCaseFolderId, worksheet, testCaseCustomProperties, testStepCustomProperties, fieldColumnMapping, testCaseCustomPropertyMapping, testStepCustomPropertyMapping, ref rowIndex, maxColumnIndex, ref truncated, errorColumn, ref errorCount, ref artifactCount, ref importCount, priorityMapping, statusMapping, typeMapping, (remoteTestFolder.IndentLevel.Length / 3) - 1, components);

                        //Move to the next row and update progress bar
                        importCount++;
                        this.UpdateProgress(importCount, null);
                    }
                    catch (Exception exception)
                    {
                        //Record the error on the sheet and add to the error count, then continue
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = exception.Message;
                        errorCount++;
                    }
                }
            }

            //Import all the test cases in the root folder
            ImportTestCasesInFolder(spiraImportExport, null, worksheet, testCaseCustomProperties, testStepCustomProperties, fieldColumnMapping, testCaseCustomPropertyMapping, testStepCustomPropertyMapping, ref rowIndex, maxColumnIndex, ref truncated, errorColumn, ref errorCount, ref artifactCount, ref importCount, priorityMapping, statusMapping, typeMapping, 0, components);

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import. Please look in the column to the right of the data to see which rows were affected.");
            }

            return importCount;
        }

        /// <summary>
        /// Exports the custom properties from the Excel sheet into Spira
        /// </summary>
        /// <param name="remoteArtifact">The current artifact entity</param>
        /// <param name="customProperties">The custom property definitions</param>
        /// <param name="dataValues">The Excel sheet data values</param>
        /// <param name="customPropertyMapping">The custom property/column mapping</param>
        private void ExportCustomProperties(RemoteArtifact remoteArtifact, RemoteCustomProperty[] customProperties, object[,] dataValues, Dictionary<int, int> customPropertyMapping)
        {
            //Iterate through the mapped custom properties
            foreach (KeyValuePair<int, int> kvp in customPropertyMapping)
            {
                int columnIndex = kvp.Value;
                int propertyNumber = kvp.Key;

                //Get the custom properties
                List<RemoteArtifactCustomProperty> artifactCustomProperties;
                if (remoteArtifact.CustomProperties == null || remoteArtifact.CustomProperties.Length < 1)
                {
                    artifactCustomProperties = new List<RemoteArtifactCustomProperty>();
                }
                else
                {
                    artifactCustomProperties = remoteArtifact.CustomProperties.ToList();
                }

                //See if we have this custom property defined
                RemoteCustomProperty customPropertyDefinition = customProperties.FirstOrDefault(c => c.PropertyNumber == propertyNumber);
                if (customPropertyDefinition != null)
                {
                    //See if we already have this custom property on the artifact, create otherwise
                    RemoteArtifactCustomProperty artifactCustomProperty = artifactCustomProperties.FirstOrDefault(a => a.PropertyNumber == propertyNumber);
                    if (artifactCustomProperty == null)
                    {
                        artifactCustomProperty = new RemoteArtifactCustomProperty();
                        artifactCustomProperty.PropertyNumber = propertyNumber;
                        artifactCustomProperties.Add(artifactCustomProperty);
                    }

                    //See what type of CP we have and handle accordingly
                    object dataValue = dataValues[1, columnIndex];
                    if (dataValue != null)
                    {
                        string cellValue = dataValue.ToString();
                        switch ((CustomPropertyTypeEnum)customPropertyDefinition.CustomPropertyTypeId)
                        {
                            case CustomPropertyTypeEnum.Boolean:
                                {
                                    if (cellValue.Trim().ToLowerInvariant() == "true")
                                    {
                                        artifactCustomProperty.BooleanValue = true;
                                    }
                                    if (cellValue.Trim().ToLowerInvariant() == "false")
                                    {
                                        artifactCustomProperty.BooleanValue = false;
                                    }
                                }
                                break;

                            case CustomPropertyTypeEnum.Date:
                                {
                                    if (dataValue.GetType() == typeof(DateTime))
                                    {
                                        artifactCustomProperty.DateTimeValue = ((DateTime)dataValue).ToUniversalTime();
                                    }
                                    else if (dataValue.GetType() == typeof(double))
                                    {
                                        DateTime dateTimeValue = DateTime.FromOADate((double)dataValue);
                                        artifactCustomProperty.DateTimeValue = dateTimeValue.ToUniversalTime();
                                    }
                                    else
                                    {
                                        string stringValue = dataValue.ToString();
                                        DateTime dateTimeValue;
                                        if (DateTime.TryParse(stringValue, out dateTimeValue))
                                        {
                                            artifactCustomProperty.DateTimeValue = dateTimeValue.ToUniversalTime();
                                        }
                                    }
                                }
                                break;

                            case CustomPropertyTypeEnum.Decimal:
                                {
                                    Decimal decimalValue;
                                    if (Decimal.TryParse(cellValue, out decimalValue))
                                    {
                                        artifactCustomProperty.DecimalValue = decimalValue;
                                    }
                                }
                                break;

                            case CustomPropertyTypeEnum.Integer:
                            case CustomPropertyTypeEnum.List:
                            case CustomPropertyTypeEnum.User:
                                {
                                    int intValue;
                                    if (Int32.TryParse(cellValue, out intValue))
                                    {
                                        artifactCustomProperty.IntegerValue = intValue;
                                    }
                                }
                                break;

                            case CustomPropertyTypeEnum.Text:
                                {
                                    artifactCustomProperty.StringValue = cellValue;
                                }
                                break;

                            case CustomPropertyTypeEnum.MultiList:
                                {
                                    string[] components = cellValue.Split(',');
                                    List<int> listIds = new List<int>();
                                    foreach (string component in components)
                                    {
                                        int intValue;
                                        if (Int32.TryParse(component, out intValue))
                                        {
                                            listIds.Add(intValue);
                                        }
                                    }
                                    artifactCustomProperty.IntegerListValue = listIds.ToArray();
                                }
                                break;
                        }
                    }
                }

                //Set the list of custom properties on the artifact object
                remoteArtifact.CustomProperties = artifactCustomProperties.ToArray();
            }
        }

        /// <summary>
        /// Imports the custom properties from Spira into the Excel sheet
        /// </summary>
        /// <param name="remoteArtifact">The current artifact entity</param>
        /// <param name="customProperties">The custom property definitions</param>
        /// <param name="dataValues">The Excel sheet data values</param>
        /// <param name="customPropertyMapping">The custom property/column mapping</param>
        private void ImportCustomProperties(RemoteArtifact remoteArtifact, RemoteCustomProperty[] customProperties, object[,] dataValues, Dictionary<int, int> customPropertyMapping)
        {
            //Iterate through the mapped custom properties
            foreach (KeyValuePair<int, int> kvp in customPropertyMapping)
            {
                int columnIndex = kvp.Value;
                int propertyNumber = kvp.Key;

                //See if we have this custom property
                RemoteArtifactCustomProperty artifactCustomProperty = remoteArtifact.CustomProperties.FirstOrDefault(a => a.PropertyNumber == propertyNumber);
                RemoteCustomProperty customPropertyDefinition = customProperties.FirstOrDefault(c => c.PropertyNumber == propertyNumber);
                if (artifactCustomProperty != null && customPropertyDefinition != null)
                {
                    //See what type of CP we have and handle accordingly
                    string cellValue = "";
                    switch ((CustomPropertyTypeEnum)customPropertyDefinition.CustomPropertyTypeId)
                    {
                        case CustomPropertyTypeEnum.Boolean:
                            {
                                if (artifactCustomProperty.BooleanValue.HasValue)
                                {
                                    cellValue = (artifactCustomProperty.BooleanValue.Value) ? "True" : "False";
                                }
                            }
                            break;

                        case CustomPropertyTypeEnum.Date:
                            {
                                if (artifactCustomProperty.DateTimeValue.HasValue)
                                {
                                    cellValue = artifactCustomProperty.DateTimeValue.Value.ToLocalTime().ToShortDateString();
                                }
                            }
                            break;

                        case CustomPropertyTypeEnum.Decimal:
                            {
                                if (artifactCustomProperty.DecimalValue.HasValue)
                                {
                                    cellValue = artifactCustomProperty.DecimalValue.Value.ToString("0.00");
                                }
                            }
                            break;

                        case CustomPropertyTypeEnum.Integer:
                        case CustomPropertyTypeEnum.List:
                        case CustomPropertyTypeEnum.User:
                            {
                                if (artifactCustomProperty.IntegerValue.HasValue)
                                {
                                    cellValue = artifactCustomProperty.IntegerValue.Value.ToString();
                                }
                            }
                            break;

                        case CustomPropertyTypeEnum.Text:
                            {
                                if (artifactCustomProperty.StringValue != null)
                                {
                                    cellValue = artifactCustomProperty.StringValue;
                                }
                            }
                            break;

                        case CustomPropertyTypeEnum.MultiList:
                            {
                                if (artifactCustomProperty.IntegerListValue != null && artifactCustomProperty.IntegerListValue.Length > 0)
                                {
                                    cellValue = artifactCustomProperty.IntegerListValue.ToFormattedString();
                                }
                            }
                            break;
                    }
                    dataValues[1, columnIndex] = cellValue;
                }
            }
        }

        /// <summary>
        /// Imports test steps belonging to a test case
        /// </summary>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <param name="fieldColumnMapping">The fields to be imported</param>
        /// <param name="worksheet">The worksheet to import the data into</param>
        /// <param name="testCaseId">The test case we're importing steps for</param>
        /// <param name="rowIndex">The row index we're adding the data at</param>
        /// <param name="maxColumnIndex">The last column on the sheet</param>
        /// <param name="truncated">Has any data truncation occurred</param>
        /// <param name="errorCount">The count of errors raised</param>
        /// <param name="errorColumn">The error column</param>
        private void ImportTestSteps(SoapServiceClient spiraImportExport, int testCaseId, Worksheet worksheet, RemoteCustomProperty[] testStepCustomProperties, Dictionary<string, int> fieldColumnMapping, Dictionary<int, int> testStepCustomPropertyMapping, ref int rowIndex, int maxColumnIndex, ref bool truncated, int errorColumn, ref int errorCount)
        {
            //Retrieve all the test steps in the test case
            RemoteTestCase remoteTestCase = spiraImportExport.TestCase_RetrieveById(testCaseId);

            //Now iterate through the test steps and populate the fields
            foreach (SpiraImportExport.RemoteTestStep remoteTestStep in remoteTestCase.TestSteps)
            {
                try
                {
                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated;
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //The Test Step description name was changed to avoid conflicting with the test case description
                        if (fieldName == "Description")
                        {
                            fieldName = "TestCaseDescription";
                        }
                        if (fieldName == "TestStepDescription")
                        {
                            fieldName = "Description";
                        }

                        //See if this field exists on the remote object (except type which is internal to the sheet)
                        if (fieldName == "Type")
                        {
                            dataValues[1, columnIndex] = ">TestStep";
                        }
                        else
                        {
                            //Need to handle the case of test step links separately
                            if (remoteTestStep.LinkedTestCaseId.HasValue && fieldName == "ExpectedResult")
                            {
                                dataValues[1, columnIndex] = remoteTestStep.LinkedTestCaseId.Value.ToString();
                            }
                            else
                            {
                                Type remoteObjectType = remoteTestStep.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanRead)
                                {
                                    //Populate the field values
                                    object propertyValue = propertyInfo.GetValue(remoteTestStep, null);

                                    //See if we have one of the long text fields to safely handle
                                    if (fieldName == "Description")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else if (fieldName == "ExpectedResult")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else if (fieldName == "SampleData")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else if (propertyInfo.PropertyType == typeof(string))
                                    {
                                        //For strings we need to verify length and truncate if necessary
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else
                                    {
                                        dataValues[1, columnIndex] = propertyValue;
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ImportCustomProperties(remoteTestStep, testStepCustomProperties, dataValues, testStepCustomPropertyMapping);

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }

                    //Move to the next row
                    rowIndex++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
        }

        /// <summary>
        /// Imports test runs
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of imported artifacts</returns>
        private int ImportTestRuns(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "TestRun_Status");

            //Retrieve all the test cases and test sets assigned to the current user
            SpiraImportExport.RemoteTestCase[] remoteTestCases = spiraImportExport.TestCase_RetrieveForOwner();
            SpiraImportExport.RemoteTestSet[] remoteTestSets = spiraImportExport.TestSet_RetrieveForOwner();

            //Now we need to create a master shell Test Run for all test sets and test cases
            List<RemoteManualTestRun> testRuns = new List<RemoteManualTestRun>();

            //First the test cases
            try
            {
                List<int> testCaseIds = new List<int>();
                if (remoteTestCases != null)
                {
                    foreach (SpiraImportExport.RemoteTestCase remoteTestCase in remoteTestCases)
                    {
                        //We need to make sure that we're only considering test cases with steps in the current project
                        if (remoteTestCase.ProjectId == importState.ProjectId && remoteTestCase.IsTestSteps)
                        {
                            testCaseIds.Add(remoteTestCase.TestCaseId.Value);
                        }
                    }
                }

                //Now create the test runs for these test cases and add to the master list
                if (testCaseIds.Count > 0)
                {
                    foreach (int testcaseId in testCaseIds)
                    {
                        RemoteManualTestRun[] testCaseTestRuns = spiraImportExport.TestRun_CreateFromTestCases(new int[] { testcaseId }, null);
                        if (testCaseTestRuns != null && testCaseTestRuns.Length > 0)
                        {
                            foreach (RemoteManualTestRun testCaseTestRun in testCaseTestRuns)
                            {
                                testRuns.Add(testCaseTestRun);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                //Rethrow as an application exception
                throw new ApplicationException(exception.Message);
            }

            try
            {
                //Next the test sets
                if (remoteTestSets != null)
                {
                    foreach (RemoteTestSet remoteTestSet in remoteTestSets)
                    {
                        //We need to make sure that we're only considering automated test sets in the current project
                        if (remoteTestSet.ProjectId == importState.ProjectId && remoteTestSet.TestRunTypeId == 1/*Manual*/)
                        {
                            RemoteManualTestRun[] testSetTestRuns = spiraImportExport.TestRun_CreateFromTestSet(remoteTestSet.TestSetId.Value);
                            if (testSetTestRuns != null && testSetTestRuns.Length > 0)
                            {
                                foreach (RemoteManualTestRun testSetTestRun in testSetTestRuns)
                                {
                                    testRuns.Add(testSetTestRun);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                //Rethrow as an application exception
                throw new ApplicationException(exception.Message);
            }

            int artifactCount = testRuns.Count;

            if (artifactCount == 0)
            {
                throw new ApplicationException("You don't have any test cases or test sets assigned to your user in the selected project. Make sure that you have test cases or test sets assigned to you in 'My Page'.");
            }

            //Set the progress bar accordingly
            this.UpdateProgress(0, artifactCount);

            //Now populate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Test #", "TestCaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Case Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release #", "ReleaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set #", "TestSetId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "TXTC #", "TestSetTestCaseId", columnIndex);
                //Test Step Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Step #", "TestStepId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Step Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Expected Result", "ExpectedResult", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Sample Data", "SampleData", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "ExecutionStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Actual Result", "ActualResult", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Incident Name", "IncidentName", columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TestCaseId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Case Name'");
            }
            if (!fieldColumnMapping.ContainsKey("TestStepId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Step #'");
            }
            if (!fieldColumnMapping.ContainsKey("Description"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Step Description'");
            }
            if (!fieldColumnMapping.ContainsKey("ExpectedResult"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Expected Result'");
            }
            if (!fieldColumnMapping.ContainsKey("ExecutionStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("ActualResult"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Actual Result'");
            }
            if (!fieldColumnMapping.ContainsKey("IncidentName"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident Name'");
            }

            //The error column is the column after the last data column
            int errorColumn = 1;
            foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
            {
                int columnIndex = fieldColumnPair.Value;
                if (columnIndex > errorColumn)
                {
                    errorColumn = columnIndex;
                }
            }
            errorColumn++;

            //Now iterate through the test cases and populate the fields
            int rowIndex = 1;
            int importCount = 0;
            bool truncated = false;
            int errorCount = 0;
            foreach (RemoteManualTestRun remoteTestRun in testRuns)
            {
                try
                {
                    //Ignore any test runs that don't have test run steps
                    if (remoteTestRun.TestRunSteps == null || remoteTestRun.TestRunSteps.Length == 0)
                    {
                        continue;
                    }

                    //For performance using VSTO Interop we need to update all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //Iterate through the various mapped fields
                    bool oldTruncated = truncated; 
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //See if this field exists on the remote object
                        Type remoteObjectType = remoteTestRun.GetType();
                        PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                        if (propertyInfo != null && propertyInfo.CanRead)
                        {
                            //Set the values
                            object propertyValue = propertyInfo.GetValue(remoteTestRun, null);
                            //See if we have one of the special known lookups
                            if (fieldName == "ExecutionStatusId")
                            {
                                if (propertyValue != null)
                                {
                                    int fieldValue = (int)propertyValue;
                                    if (statusMapping.ContainsKey(fieldValue))
                                    {
                                        string lookupValue = statusMapping[fieldValue];
                                        dataValues[1, columnIndex] = lookupValue;
                                    }
                                }
                            }
                            else if (fieldName == "Description")
                            {
                                //Need to strip off any formatting and make sure it's not too long
                                dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                            }
                            else if (propertyInfo.PropertyType == typeof(string))
                            {
                                //For strings we need to verify length and truncate if necessary
                                dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                            }
                            else
                            {
                                dataValues[1, columnIndex] = propertyValue;
                            }
                        }
                    }

                    //Now commit the data
                    dataRange.Value2 = dataValues;

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Now we need to import any Test Run Steps
                    rowIndex++;
                    foreach (SpiraImportExport.RemoteTestRunStep remoteTestRunStep in remoteTestRun.TestRunSteps)
                    {
                        try
                        {
                            //For performance using VSTO Interop we need to update all the fields in the row in one go
                            dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex + 2, 1], worksheet.Cells[rowIndex + 2, maxColumnIndex]];
                            dataValues = (object[,])dataRange.Value2;

                            //Iterate through the various mapped fields
                            bool oldTruncated2 = truncated;
                            foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                            {
                                int columnIndex = fieldColumnPair.Value;
                                string fieldName = fieldColumnPair.Key;

                                //See if this field exists on the remote object
                                Type remoteObjectType = remoteTestRunStep.GetType();
                                PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                if (propertyInfo != null && propertyInfo.CanRead)
                                {
                                    //Set the values
                                    object propertyValue = propertyInfo.GetValue(remoteTestRunStep, null);
                                    //See if we have one of the special known lookups
                                    if (fieldName == "ExecutionStatusId")
                                    {
                                        if (propertyValue != null)
                                        {
                                            int fieldValue = (int)propertyValue;
                                            if (statusMapping.ContainsKey(fieldValue))
                                            {
                                                string lookupValue = statusMapping[fieldValue];
                                                dataValues[1, columnIndex] = lookupValue;
                                            }
                                        }
                                    }
                                    else if (fieldName == "ExpectedResult")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else if (fieldName == "SampleData")
                                    {
                                        //Need to strip off any formatting and make sure it's not too long
                                        dataValues[1, columnIndex] = CleanTruncateLongText(propertyValue, ref truncated);
                                    }
                                    else
                                    {
                                        dataValues[1, columnIndex] = propertyValue;
                                    }
                                }
                            }

                            //If it was truncated on this row, display a message in the right-most column
                            if (truncated && !oldTruncated2)
                            {
                                Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                                errorCell.Value2 = "This row had data truncated.";
                            }

                            //Check for abort condition
                            if (this.IsAborted)
                            {
                                throw new ApplicationException("Import aborted by user.");
                            }

                            //Now commit the data
                            dataRange.Value2 = dataValues;
                            rowIndex++;
                        }
                        catch (Exception exception)
                        {
                            //Record the error on the sheet and add to the error count, then continue
                            Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                            errorCell.Value2 = exception.Message;
                            errorCount++;
                        }
                    }

                    //Move to the next row and update progress bar
                    importCount++;
                    this.UpdateProgress(importCount, null);

                    //If it was truncated on this row, display a message in the right-most column
                    if (truncated && !oldTruncated)
                    {
                        Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                        errorCell.Value2 = "This row had data truncated.";
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex + 2, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }
            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Import failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            //Check for truncation
            if (truncated)
            {
                throw new ApplicationException("Some of the long text fields were truncated during import, please check them before Exporting back to Spira.");
            }

            return importCount;
        }

        /// <summary>
        /// Exports test runs
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportTestRuns(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "TestRun_Status");

            //Retrieve all the test cases and test sets assigned to the current user
            SpiraImportExport.RemoteTestCase[] remoteTestCases = spiraImportExport.TestCase_RetrieveForOwner();
            SpiraImportExport.RemoteTestSet[] remoteTestSets = spiraImportExport.TestSet_RetrieveForOwner();

            //Now we need to create a master shell Test Run for all test sets and test cases
            List<RemoteManualTestRun> testRuns = new List<RemoteManualTestRun>();

            //First the test cases
            List<int> testCaseIds = new List<int>();
            if (remoteTestCases != null)
            {
                foreach (SpiraImportExport.RemoteTestCase remoteTestCase in remoteTestCases)
                {
                    //We need to make sure that we're only considering test cases in the current project that have steps
                    if (remoteTestCase.ProjectId == importState.ProjectId && remoteTestCase.IsTestSteps)
                    {
                        testCaseIds.Add(remoteTestCase.TestCaseId.Value);
                    }
                }
            }

            //Now create the test runs for these test cases and add to the master list
            if (testCaseIds.Count > 0)
            {
                RemoteManualTestRun[] testCaseTestRuns = spiraImportExport.TestRun_CreateFromTestCases(testCaseIds.ToArray(), null);
                foreach (RemoteManualTestRun testCaseTestRun in testCaseTestRuns)
                {
                    testRuns.Add(testCaseTestRun);
                }
            }

            //Next the test sets
            if (remoteTestSets != null)
            {
                foreach (RemoteTestSet remoteTestSet in remoteTestSets)
                {
                    //We need to make sure that we're only considering manual test sets in the current project
                    if (remoteTestSet.ProjectId == importState.ProjectId && remoteTestSet.TestRunTypeId == 1/*Manual*/)
                    {
                        RemoteManualTestRun[] testSetTestRuns = spiraImportExport.TestRun_CreateFromTestSet(remoteTestSet.TestSetId.Value);
                        foreach (RemoteManualTestRun testSetTestRun in testSetTestRuns)
                        {
                            testRuns.Add(testSetTestRun);
                        }
                    }
                }
            }

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Test #", "TestCaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Case Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release #", "ReleaseId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Set #", "TestSetId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "TXTC #", "TestSetTestCaseId", columnIndex);
                //Test Step Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Step #", "TestStepId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Test Step Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Expected Result", "ExpectedResult", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Sample Data", "SampleData", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "ExecutionStatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Actual Result", "ActualResult", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Incident Name", "IncidentName", columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Export aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("TestCaseId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Case Name'");
            }
            if (!fieldColumnMapping.ContainsKey("TestStepId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Step #'");
            }
            if (!fieldColumnMapping.ContainsKey("Description"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Test Step Description'");
            }
            if (!fieldColumnMapping.ContainsKey("ExpectedResult"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Expected Result'");
            }
            if (!fieldColumnMapping.ContainsKey("ExecutionStatusId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Status'");
            }
            if (!fieldColumnMapping.ContainsKey("ActualResult"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Actual Result'");
            }
            if (!fieldColumnMapping.ContainsKey("IncidentName"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Incident Name'");
            }

            //The error column is the column after the last data column
            int errorColumn = 1;
            foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
            {
                int columnIndex = fieldColumnPair.Value;
                if (columnIndex > errorColumn)
                {
                    errorColumn = columnIndex;
                }
            }
            errorColumn++;

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count - 2);

            //See if we need to use a specific date for this execution
            DateTime execDate = DateTime.UtcNow;
            if (Configuration.Default.TestRunDate.HasValue)
            {
                execDate = Configuration.Default.TestRunDate.Value;
            }

            //Now iterate through the rows in the sheet that have data and update the test run accordingly
            int exportCount = 0;
            int errorCount = 0;
            bool lastRecord = false;
            int previousTestCaseId = -1;
            Nullable<int> releaseId = null;
            Nullable<int> testSetId = null;
            Nullable<int> testSetTestCaseId = null;
            Dictionary<string, string> linkedIncidents = new Dictionary<string, string>();
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //First see if we have a test case or test step row
                    int testStepId = -1;
                    int testCaseId = -1; 
                    int testCaseIdColumnIndex = fieldColumnMapping["TestCaseId"];
                    object dataValue2 = dataValues[1, testCaseIdColumnIndex];
                    if (dataValue2 != null && dataValue2.ToString().Trim() != "")
                    {
                        string testCaseIdString = dataValue2.ToString();
                        int tempId;
                        if (Int32.TryParse(testCaseIdString, out tempId))
                        {
                            testCaseId = tempId;
                        }
                    }
                    int testStepIdColumnIndex = fieldColumnMapping["TestStepId"];
                    dataValue2 = dataValues[1, testStepIdColumnIndex];
                    if (dataValue2 != null && dataValue2.ToString().Trim() != "")
                    {
                        string testStepIdString = dataValue2.ToString();
                        int tempId;
                        if (Int32.TryParse(testStepIdString, out tempId))
                        {
                            testStepId = tempId;
                        }
                    }
                    if (testCaseId == -1 && testStepId == -1)
                    {
                        //End the import if we have neither case
                        lastRecord = true;
                        break;
                    }

                    //If we have a test case, we just need to note down the id for later
                    //Note that test steps may also have test case id values of their own (since could be linked steps)
                    if (testStepId == -1)
                    {
                        previousTestCaseId = testCaseId;

                        //Also need to capture the test set, release and test set-test case ids
                        if (fieldColumnMapping.ContainsKey("ReleaseId"))
                        {
                            object value = dataValues[1, fieldColumnMapping["ReleaseId"]];
                            if (value != null)
                            {
                                releaseId = Int32.Parse(value.ToString());
                            }
                            else
                            {
                                releaseId = null;
                            }
                        }
                        else
                        {
                            releaseId = null;
                        }
                        if (fieldColumnMapping.ContainsKey("TestSetId"))
                        {
                            object value = dataValues[1, fieldColumnMapping["TestSetId"]];
                            if (value != null)
                            {
                                testSetId = Int32.Parse(value.ToString());
                            }
                            else
                            {
                                testSetId = null;
                            }
                        }
                        else
                        {
                            testSetId = null;
                        }
                        if (fieldColumnMapping.ContainsKey("TestSetTestCaseId"))
                        {
                            object value = dataValues[1, fieldColumnMapping["TestSetTestCaseId"]];
                            if (value != null)
                            {
                                testSetTestCaseId = Int32.Parse(value.ToString());
                            }
                            else
                            {
                                testSetTestCaseId = null;
                            }
                        }
                        else
                        {
                            testSetTestCaseId = null;
                        }
                    }

                    //If we have a test step, need to update the master Test Run data object appropriately
                    if (testStepId != -1 && previousTestCaseId != -1)
                    {
                        RemoteManualTestRun matchedTestRun = null;
                        RemoteTestRunStep matchedTestRunStep = null;
                        foreach (RemoteManualTestRun testRun in testRuns)
                        {
                            if (testRun.TestCaseId == previousTestCaseId)
                            {
                                matchedTestRun = testRun;
                                matchedTestRun.ReleaseId = releaseId;
                                matchedTestRun.TestSetId = testSetId;
                                matchedTestRun.TestSetTestCaseId = testSetTestCaseId;
                                matchedTestRun.StartDate = execDate;
                                matchedTestRun.EndDate = execDate;
                                break;
                            }
                        }
                        foreach (SpiraImportExport.RemoteTestRunStep testRunStep in matchedTestRun.TestRunSteps)
                        {
                            if (testRunStep.TestStepId == testStepId)
                            {
                                matchedTestRunStep = testRunStep;
                                break;
                            }
                        }
                        
                        //Make sure we have a matching test run step row
                        if (matchedTestRun != null && matchedTestRunStep != null)
                        {
                            //Iterate through the various mapped fields
                            foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                            {
                                int columnIndex = fieldColumnPair.Value;
                                string fieldName = fieldColumnPair.Key;

                                //Incident name is not part of the data-object, but is just stored in the local dictionary
                                object dataValue = dataValues[1, columnIndex];
                                if (fieldName == "IncidentName")
                                {
                                    //We need to get the incident name and associate it with the appropriate test step and test case
                                    if (dataValue != null)
                                    {
                                        string key = previousTestCaseId + "+" + testStepId;
                                        string incidentName = MakeXmlSafe(dataValue);
                                        if (!linkedIncidents.ContainsKey(key))
                                        {
                                            linkedIncidents.Add(key, incidentName);
                                        }
                                    }
                                }
                                else
                                {
                                    Type remoteObjectType = matchedTestRunStep.GetType();
                                    PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                                    if (propertyInfo != null && propertyInfo.CanWrite)
                                    {
                                        //See if we have one of the special known lookups
                                        if (fieldName == "ExecutionStatusId")
                                        {
                                            if (dataValue == null)
                                            {
                                                //This field is not nullable, so we need to pass 1 to default to 'Not Started'
                                                propertyInfo.SetValue(matchedTestRunStep, 1, null);
                                            }
                                            else
                                            {
                                                string lookupValue = MakeXmlSafe(dataValue);
                                                int fieldValue = -1;
                                                foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                                {
                                                    if (mappingEntry.Value == lookupValue)
                                                    {
                                                        fieldValue = mappingEntry.Key;
                                                        break;
                                                    }
                                                }
                                                if (fieldValue != -1)
                                                {
                                                    propertyInfo.SetValue(matchedTestRunStep, fieldValue, null);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //Make sure that we do any necessary type conversion
                                            //Make sure the field handles nullable types
                                            if (dataValue == null)
                                            {
                                                if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                                {
                                                    propertyInfo.SetValue(matchedTestRunStep, null, null);
                                                }
                                            }
                                            else
                                            {
                                                if (propertyInfo.PropertyType == typeof(string))
                                                {
                                                    if (dataValue.GetType() == typeof(string))
                                                    {
                                                        //Need to handle large string issue
                                                        SafeSetStringValue(propertyInfo, matchedTestRunStep, MakeXmlSafe(dataValue));
                                                    }
                                                    else
                                                    {
                                                        propertyInfo.SetValue(matchedTestRunStep, dataValue.ToString(), null);
                                                    }
                                                }
                                                if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                                {
                                                    if (dataValue.GetType() == typeof(int))
                                                    {
                                                        propertyInfo.SetValue(matchedTestRunStep, dataValue, null);
                                                    }
                                                    else
                                                    {
                                                        string stringValue = dataValue.ToString();
                                                        int intValue;
                                                        if (Int32.TryParse(stringValue, out intValue))
                                                        {
                                                            propertyInfo.SetValue(matchedTestRunStep, intValue, null);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }

                //Move to the next row and update progress bar
                exportCount++;
                this.UpdateProgress(exportCount, null);
            }

            //Now we need to sent the updated test results back to the Spira web service
            //Now we need to sent the updated test results back to the Spira web service
            List<RemoteManualTestRun> savedTestRuns = new List<RemoteManualTestRun>();
            foreach (RemoteManualTestRun run in testRuns.ToArray())
            {
                savedTestRuns.AddRange(spiraImportExport.TestRun_Save(new RemoteManualTestRun[] { run }, execDate));
            }
            RemoteManualTestRun[] savedRemoteTestRuns = savedTestRuns.ToArray();

            //See if we have rich text enabled
            bool richTextEnabled = true;    //Default in Spira is true
            RemoteSetting[] remoteSettings = spiraImportExport.System_GetSettings();
            if (remoteSettings != null)
            {
                foreach (RemoteSetting remoteSetting in remoteSettings)
                {
                    if (remoteSetting.Name == "SpiraTest.RichTextArtifactDesc" && remoteSetting.Value == "N")
                    {
                        richTextEnabled = false;
                    }
                }
            }

            //Now add the incident records, linked to the test runs
            if (savedRemoteTestRuns != null)
            {
                foreach (RemoteManualTestRun remoteTestRun in savedRemoteTestRuns)
                {
                    if (remoteTestRun.TestRunSteps != null)
                    {
                        foreach (SpiraImportExport.RemoteTestRunStep savedTestRunStep in remoteTestRun.TestRunSteps)
                        {
                            string key = remoteTestRun.TestCaseId + "+" + savedTestRunStep.TestStepId;
                            if (linkedIncidents.ContainsKey(key))
                            {
                                //Depending on whether SpiraTest is configured for rich-text editing
                                //create the incident description accordingly
                                string description;
                                if (richTextEnabled)
                                {
                                    description =
                                        "<b><u>Description:</u></b><br />\n" +
                                        savedTestRunStep.Description + "<br /><br />\n<b><u>Expected Result:</u></b><br />\n" +
                                        savedTestRunStep.ExpectedResult + "<br /><br />\n<b><u>Actual Result:</u></b><br />\n" + savedTestRunStep.ActualResult + "<br />";
                                }
                                else
                                {
                                    description =
                                        savedTestRunStep.Description + " - " +
                                        savedTestRunStep.ExpectedResult + " - " +
                                        savedTestRunStep.ActualResult;
                                }

                                string incidentName = linkedIncidents[key];
                                SpiraImportExport.RemoteIncident remoteIncident = new SpiraImportExport.RemoteIncident();
                                remoteIncident.Name = incidentName;
                                remoteIncident.IncidentStatusId = null;   //Default
                                remoteIncident.IncidentTypeId = null;   //Default
                                remoteIncident.Description = description;

                                //Link to the test run steps
                                if (savedTestRunStep.TestRunStepId.HasValue)
                                {
                                    remoteIncident.TestRunStepIds = new int[1] { savedTestRunStep.TestRunStepId.Value };
                                }
                                spiraImportExport.Incident_Create(remoteIncident);
                            }
                        }
                    }
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Checks for the presence of a known single header cell and adds it to the mapping collection
        /// </summary>
        /// <param name="fieldColumnMapping">The mapping collection</param>
        /// <param name="cell">The Excel cell</param>
        /// <param name="label">The label who's name we're looking for</param>
        /// <param name="fieldName">The name of the object field the data relates to</param>
        /// <param name="columnIndex">The 1-based column index</param>
        protected void CheckForHeaderCell(Dictionary<string, int> fieldColumnMapping, Range cell, string label, string fieldName, int columnIndex)
        {
            if (cell.Value2 != null && cell.Value2.GetType() == typeof(string))
            {
                string value = (string)cell.Value2;
                if (value.Trim().ToLowerInvariant() == label.ToLowerInvariant())
                {
                    if (!fieldColumnMapping.ContainsKey(fieldName))
                    {
                        fieldColumnMapping.Add(fieldName, columnIndex);
                    }
                }
            }
        }

        /// <summary>
        /// Checks for the presence of the various custom property header cells and adds them to the mapping collection
        /// </summary>
        /// <param name="customProperties">The project's custom property definitions</param>
        /// <param name="fieldColumnMapping">The mapping collection</param>
        /// <param name="cell">The Excel cell</param>
        /// <param name="columnIndex">The 1-based column index</param>
        protected void CheckForCustomPropHeaderCells(Dictionary<int, int> customPropertyMapping, RemoteCustomProperty[] customProperties, Range cell, int columnIndex)
        {
            if (cell.Value2 != null && cell.Value2.GetType() == typeof(string))
            {
                string value = (string)cell.Value2;

                //Loop through the custom properties
                foreach (RemoteCustomProperty customProperty in customProperties)
                {
                    //See if the column exists
                    string label = String.Format("CUS-{0:00}", customProperty.PropertyNumber);

                    if (value.Trim().ToUpperInvariant() == label.ToUpperInvariant())
                    {
                        if (!customPropertyMapping.ContainsKey(customProperty.PropertyNumber))
                        {
                            customPropertyMapping.Add(customProperty.PropertyNumber, columnIndex);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Exports data to Spira from the excel sheet
        /// </summary>
        /// <param name="artifactTypeName">The type of data we're exporting</param>
        /// <param name="projectId">The id of the project to export to</param>
        public void Export(int projectId, string artifactTypeName)
        {
            //By default, it's not aborted
            this.isAborted = false;

            //Make sure that the handles to the progress dialog and application are available
            if (this.ProgressForm == null)
            {
                throw new ApplicationException("Unable to get handle to progress form. Aborting Export");
            }
            if (this.ExcelApplication == null)
            {
                throw new ApplicationException("Unable to get handle to Excel application instance. Aborting Export");
            }

            //Make sure we have a workbook loaded
            if (this.ExcelApplication.ActiveWorkbook == null || this.ExcelApplication.Worksheets == null || this.ExcelApplication.Worksheets.Count == 0)
            {
                throw new ApplicationException("No Excel worksheet is currently loaded. Please open the Excel import template");
            }

            //Make sure that the required worksheets exist
            Worksheet importWorksheet = null;
            foreach (Worksheet worksheet in this.ExcelApplication.Worksheets)
            {
                if (worksheet.Name.Trim().ToLowerInvariant() == artifactTypeName.Trim().ToLowerInvariant())
                {
                    importWorksheet = worksheet;
                    break;
                }
            }
            if (importWorksheet == null)
            {
                throw new ApplicationException("Unable to locate a worksheet with name '" + artifactTypeName + "'. Aborting Export");
            }

            //Worksheet containing lookups
            Worksheet lookupWorksheet = null;
            foreach (Worksheet worksheet in this.ExcelApplication.Worksheets)
            {
                if (worksheet.Name.Trim().ToLowerInvariant() == "lookups")
                {
                    lookupWorksheet = worksheet;
                    break;
                }
            }
            if (lookupWorksheet == null)
            {
                throw new ApplicationException("Unable to locate a worksheet with name 'Lookups'. Aborting Export");
            }

            //Start the background thread that performs the export
            ImportState importState = new ImportState();
            importState.ProjectId = projectId;
            importState.ArtifactTypeName = artifactTypeName;
            importState.ExcelWorksheet = importWorksheet;
            importState.LookupWorksheet = lookupWorksheet;
            ThreadPool.QueueUserWorkItem(new WaitCallback(this.Export_Process), importState);
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

                //Now see which data is being exported and handle accordingly
                int exportCount = 0;
                switch (importState.ArtifactTypeName)
                {
                    case "Requirements":
                        exportCount = ExportRequirements(spiraImportExport, importState);
                        break;
                    case "Releases":
                        exportCount = ExportReleases(spiraImportExport, importState);
                        break;
                    case "Test Sets":
                        exportCount = ExportTestSets(spiraImportExport, importState);
                        break;
                    case "Test Cases":
                        exportCount = ExportTestCases(spiraImportExport, importState);
                        break;
                    case "Test Runs":
                        exportCount = ExportTestRuns(spiraImportExport, importState);
                        break;
                    case "Incidents":
                        exportCount = ExportIncidents(spiraImportExport, importState);
                        break;
                    case "Tasks":
                        exportCount = ExportTasks(spiraImportExport, importState);
                        break;
                    case "Custom Values":
                        exportCount = ExportCustomValues(spiraImportExport, importState);
                        break;
                }

                //Set the progress indicator to 100%
                UpdateProgress(exportCount, exportCount);

                //Raise the success event
                OnOperationCompleted(exportCount);
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
            catch (Exception exception)
            {
                //If we get an exception need to raise an error event that the form displays
                OnErrorOccurred(exception.Message);
            }
        }

        /// <summary>
        /// Exports requirements
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportRequirements(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //Get any lookup ranges that we need to transform the data
            Dictionary<int, string> statusMapping = LoadLookup(importState.LookupWorksheet, "Req_Status");
            Dictionary<int, string> typeMapping = LoadLookup(importState.LookupWorksheet, "Req_Type");
            Dictionary<int, string> importanceMapping = LoadLookup(importState.LookupWorksheet, "Req_Importance");

            //Get the custom property definitions for the current project
            RemoteCustomProperty[] customProperties = spiraImportExport.CustomProperty_RetrieveForArtifactType((int)ArtifactTypeEnum.Requirement, false);

            //Get the list of components currently in this project
            RemoteComponent[] components = spiraImportExport.Component_Retrieve(true, false);

            //Get the list of releases currently in this project
            RemoteRelease[] releases = spiraImportExport.Release_Retrieve(true);

            //Now validate the Excel Sheet. All headers need to be in the first two rows of the worksheet
            //Find the needed cells and populate a lookup dictionary
            Worksheet worksheet = importState.ExcelWorksheet;
            Dictionary<string, int> fieldColumnMapping = new Dictionary<string, int>();
            Dictionary<int, int> customPropertyMapping = new Dictionary<int, int>();
            int headerRowIndex = 2;
            int maxColumnIndex = worksheet.Columns.Count;
            if (maxColumnIndex > worksheet.UsedRange.Columns.Count)
            {
                maxColumnIndex = worksheet.UsedRange.Columns.Count;
            }
            for (int columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                Range cell = (Range)worksheet.Cells[headerRowIndex, columnIndex];
                //Standard Fields
                CheckForHeaderCell(fieldColumnMapping, cell, "Req #", "RequirementId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement Name", "Name", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Requirement Description", "Description", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Release Version", "ReleaseVersionNumber", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Importance", "ImportanceId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Status", "StatusId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Type", "RequirementTypeId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Author", "AuthorId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Owner", "OwnerId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Component", "ComponentId", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Estimate", "EstimatePoints", columnIndex);

                //These are not actually fields but child collections
                CheckForHeaderCell(fieldColumnMapping, cell, "Comment", "Comment", columnIndex);
                CheckForHeaderCell(fieldColumnMapping, cell, "Linked Requirements", "LinkedRequirements", columnIndex);

                //Custom Properties
                CheckForCustomPropHeaderCells(customPropertyMapping, customProperties, cell, columnIndex);

                //Check for abort condition
                if (this.IsAborted)
                {
                    throw new ApplicationException("Import aborted by user.");
                }
            }

            //Make sure all the required fields are populated
            if (!fieldColumnMapping.ContainsKey("RequirementId"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Req #'");
            }
            if (!fieldColumnMapping.ContainsKey("Name"))
            {
                throw new ApplicationException("Unable to find a column heading with name 'Requirement Name'");
            }

            //The error column is the column after the last data column
            int errorColumn = GetErrorColumn(fieldColumnMapping, customPropertyMapping);

            //Set the progress max-value
            this.UpdateProgress(0, worksheet.UsedRange.Rows.Count-2);

            //Now iterate through the rows in the sheet that have data
            Dictionary<int, int> parentPrimaryKeys = new Dictionary<int, int>();
            int exportCount = 0;
            int currentIndent = 0;
            int errorCount = 0;
            bool lastRecord = false;
            for (int rowIndex = 3; rowIndex < worksheet.UsedRange.Rows.Count + 1 && !lastRecord; rowIndex++)
            {
                try
                {
                    //For performance using VSTO Interop we need to read all the fields in the row in one go
                    Range dataRange = (Range)worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, maxColumnIndex]];
                    object[,] dataValues = (object[,])dataRange.Value2;

                    //See if we are inserting a new requirement or updating an existing one
                    int primaryKeyColumn = fieldColumnMapping["RequirementId"];
                    SpiraImportExport.RemoteRequirement remoteRequirement = null;
                    List<int> linkedRequirementIds = new List<int>();
                    if (dataValues[1, primaryKeyColumn] == null)
                    {
                        //We have insert case
                        remoteRequirement = new SpiraExcelAddIn.SpiraImportExport.RemoteRequirement();

                        //Default to type 'feature' and status 'requested' unless otherwise specified
                        remoteRequirement.RequirementTypeId = 2;
                        remoteRequirement.StatusId = 1;
                    }
                    else
                    {
                        //We have update case
                        string requirementIdString = dataValues[1, primaryKeyColumn].ToString();
                        int requirementId;
                        if (!Int32.TryParse(requirementIdString, out requirementId))
                        {
                            throw new ApplicationException("Requirement ID '" + requirementIdString + "' is not valid. It needs to be a purely integer value.");
                        }
                        remoteRequirement = spiraImportExport.Requirement_RetrieveById(requirementId);
                    }

                    //Iterate through the various mapped fields
                    int indentLevel = 0;
                    string newComment = "";
                    foreach (KeyValuePair<string, int> fieldColumnPair in fieldColumnMapping)
                    {
                        int columnIndex = fieldColumnPair.Value;
                        string fieldName = fieldColumnPair.Key;

                        //If we have the name field, need to use that to determine the indent
                        //and also to know when we've reached the end of the import
                        if (fieldName == "Name")
                        {
                            Range nameCell = (Range)worksheet.Cells[rowIndex, columnIndex];
                            if (nameCell == null || nameCell.Value2 == null)
                            {
                                lastRecord = true;
                                break;
                            }
                            else
                            {
                                indentLevel = (int)nameCell.IndentLevel;
                                //Add to the dictionary if not a new item
                                if (remoteRequirement.RequirementId.HasValue)
                                {
                                    if (!parentPrimaryKeys.ContainsKey(indentLevel))
                                    {
                                        parentPrimaryKeys.Add(indentLevel,remoteRequirement.RequirementId.Value);
                                    }
                                    else
                                    {
                                        parentPrimaryKeys[indentLevel] = remoteRequirement.RequirementId.Value;
                                    }
                                }
                            }
                        }

                        //See if this field exists on the remote object (except Comment and Linked Requirements which are handled separately)
                        if (fieldName == "Comment")
                        {
                            object dataValue = dataValues[1, columnIndex];
                            if (dataValue != null)
                            {
                                newComment = (MakeXmlSafe(dataValue)).Trim();
                            }
                        }
                        else if (fieldName == "LinkedRequirements")
                        {
                            object dataValue = dataValues[1, columnIndex];
                            if (dataValue != null && dataValue is String)
                            {
                                string[] ids = ((string)dataValue).Split(',');
                                foreach (string id in ids)
                                {
                                    int requirementId;
                                    if (Int32.TryParse(id, out requirementId))
                                    {
                                        linkedRequirementIds.Add(requirementId);
                                    }
                                }
                            }
                        }
                        else
                        {
                            Type remoteObjectType = remoteRequirement.GetType();
                            PropertyInfo propertyInfo = remoteObjectType.GetProperty(fieldName);
                            if (propertyInfo != null && propertyInfo.CanWrite)
                            {
                                object dataValue = dataValues[1, columnIndex];

                                //See if we have one of the special known lookups
                                if (fieldName == "StatusId")
                                {
                                    if (dataValue == null)
                                    {
                                        //This field is not nullable, so we need to pass 1 to default to 'Requested'
                                        propertyInfo.SetValue(remoteRequirement, 1, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = 1; //Default to requested
                                        foreach (KeyValuePair<int, string> mappingEntry in statusMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteRequirement, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "ImportanceId")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteRequirement, null, null);
                                    }
                                    else
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = -1;
                                        foreach (KeyValuePair<int, string> mappingEntry in importanceMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteRequirement, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "RequirementTypeId")
                                {
                                    if (dataValue != null)
                                    {
                                        string lookupValue = MakeXmlSafe(dataValue);
                                        int fieldValue = -1;
                                        foreach (KeyValuePair<int, string> mappingEntry in typeMapping)
                                        {
                                            if (mappingEntry.Value == lookupValue)
                                            {
                                                fieldValue = mappingEntry.Key;
                                                break;
                                            }
                                        }
                                        if (fieldValue != -1)
                                        {
                                            propertyInfo.SetValue(remoteRequirement, fieldValue, null);
                                        }
                                    }
                                }
                                else if (fieldName == "ReleaseVersionNumber")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteRequirement, null, null);
                                    }
                                    else if (dataValue is String)
                                    {
                                        //We need to get the version number and find the corresponding ReleaseId, if it exists
                                        string versionNumber = (string)dataValue;
                                        RemoteRelease release = releases.FirstOrDefault(r => r.VersionNumber.Trim() == versionNumber.Trim());
                                        if (release != null && release.ReleaseId.HasValue)
                                        {
                                            remoteRequirement.ReleaseId = release.ReleaseId.Value;
                                        }
                                    }
                                }
                                else if (fieldName == "ComponentId")
                                {
                                    if (dataValue == null)
                                    {
                                        propertyInfo.SetValue(remoteRequirement, null, null);
                                    }
                                    else if (dataValue is String)
                                    {
                                        //We need to get the component name and find the corresponding ComponentId, if it exists
                                        string componentName = (string)dataValue;
                                        RemoteComponent component = components.FirstOrDefault(c => c.Name.Trim() == componentName.Trim());
                                        if (component != null && component.ComponentId.HasValue)
                                        {
                                            remoteRequirement.ComponentId = component.ComponentId.Value;
                                        }
                                    }
                                }
                                else
                                {
                                    //Make sure that we do any necessary type conversion
                                    //Make sure the field handles nullable types
                                    if (dataValue == null)
                                    {
                                        if (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                        {
                                            propertyInfo.SetValue(remoteRequirement, null, null);
                                        }
                                    }
                                    else
                                    {
                                        if (propertyInfo.PropertyType == typeof(string))
                                        {
                                            if (dataValue.GetType() == typeof(string))
                                            {
                                                //Need to handle large string issue
                                                SafeSetStringValue(propertyInfo, remoteRequirement, MakeXmlSafe(dataValue));

                                            }
                                            else
                                            {
                                                propertyInfo.SetValue(remoteRequirement, dataValue.ToString(), null);
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(Nullable<int>))
                                        {
                                            if (dataValue.GetType() == typeof(int))
                                            {
                                                propertyInfo.SetValue(remoteRequirement, dataValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                int intValue;
                                                if (Int32.TryParse(stringValue, out intValue))
                                                {
                                                    propertyInfo.SetValue(remoteRequirement, intValue, null);
                                                }
                                            }
                                        }
                                        if (propertyInfo.PropertyType == typeof(decimal) || propertyInfo.PropertyType == typeof(Nullable<decimal>))
                                        {
                                            if (dataValue.GetType() == typeof(decimal))
                                            {
                                                propertyInfo.SetValue(remoteRequirement, dataValue, null);
                                            }
                                            else
                                            {
                                                string stringValue = dataValue.ToString();
                                                decimal decimalValue;
                                                if (Decimal.TryParse(stringValue, out decimalValue))
                                                {
                                                    propertyInfo.SetValue(remoteRequirement, decimalValue, null);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Iterate through all the custom properties
                    ExportCustomProperties(remoteRequirement, customProperties, dataValues, customPropertyMapping);

                    if (lastRecord)
                    {
                        break;
                    }

                    //Now either insert or update the requirement
                    if (remoteRequirement.RequirementId.HasValue)
                    {
                        spiraImportExport.Requirement_Update(remoteRequirement);
                    }
                    else
                    {
                        //Insert case

                        //If we have an item already loaded that is a parent of this, then use the insert child API method
                        if (parentPrimaryKeys.ContainsKey(indentLevel-1))
                        {
                            remoteRequirement = spiraImportExport.Requirement_Create2(remoteRequirement, parentPrimaryKeys[indentLevel - 1]);
                        }
                        else
                        {
                            //Need to determine the relative indent offset
                            int indentOffset = indentLevel - currentIndent;
                            remoteRequirement = spiraImportExport.Requirement_Create1(remoteRequirement, indentOffset);
                            currentIndent = indentLevel;
                        }
                        //Update the cell with the requirement ID to prevent multiple-appends
                        Range newKeyCell = (Range)worksheet.Cells[rowIndex, primaryKeyColumn];
                        newKeyCell.Value2 = remoteRequirement.RequirementId;
                    }

                    //Add a comment if necessary
                    if (newComment != "")
                    {
                        SpiraImportExport.RemoteComment remoteComment = new SpiraImportExport.RemoteComment();
                        remoteComment.ArtifactId = remoteRequirement.RequirementId.Value;
                        remoteComment.Text = newComment;
                        remoteComment.CreationDate = DateTime.UtcNow;
                        spiraImportExport.Requirement_CreateComment(remoteComment);
                    }

                    //Now we need to add any associations
                    if (linkedRequirementIds != null && linkedRequirementIds.Count > 0)
                    {
                        foreach (int requirementId in linkedRequirementIds)
                        {
                            try
                            {
                                RemoteAssociation remoteAssociation = new RemoteAssociation();
                                remoteAssociation.ArtifactLinkTypeId = 1;   /* Related-To */
                                remoteAssociation.SourceArtifactTypeId = 1; /* Requirement */
                                remoteAssociation.SourceArtifactId = remoteRequirement.RequirementId.Value;
                                remoteAssociation.DestArtifactTypeId = 1; /* Requirement */
                                remoteAssociation.DestArtifactId = requirementId;
                                remoteAssociation.CreationDate = DateTime.UtcNow;
                                spiraImportExport.Association_Create(remoteAssociation);
                            }
                            catch (Exception exception)
                            {
                                //Simply let the outer exception handler deal with it.
                                throw;
                            }
                        }
                    }

                    //Move to the next row and update progress bar
                    exportCount++;
                    this.UpdateProgress(exportCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (FaultException<ValidationFaultMessage> validationException)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = GetValidationFaultDetail(validationException);
                    errorCount++;
                }
                catch (Exception exception)
                {
                    //Record the error on the sheet and add to the error count, then continue
                    Range errorCell = (Range)worksheet.Cells[rowIndex, errorColumn];
                    errorCell.Value2 = exception.Message;
                    errorCount++;
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please look in the column to the right of the data for the appropriate error message.");
            }

            return exportCount;
        }

        /// <summary>
        /// Aborts the current operation
        /// </summary>
        public void AbortOperation()
        {
            //Set the abort flag, which the background thread will see
            this.isAborted = true;
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
