using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Xml;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace SpiraWordAddIn
{
    /// <summary>
    /// Contains the logic to import/export data from SpiraTeam to/from the Word document
    /// </summary>
    public class Importer
    {
        public const string SOAP_RELATIVE_URL = "Services/v5_0/SoapService.svc";
        public const int NAVIGATION_ID_ATTACHMENTS = -14;

        private const int COM_TRUE = -1;
        private const int COM_FALSE = 0;

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
        /// The handle to the word document
        /// </summary>
        public Microsoft.Office.Interop.Word._Application WordApplication
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

        /// <summary>
        /// The list of styles mapped
        /// </summary>
        public Dictionary<SpiraRibbon.MappedStyleKeys, string> MappedStyles
        {
            get;
            set;
        }

        #endregion

        /// <summary>
        /// Aborts the current operation
        /// </summary>
        public void AbortOperation()
        {
            //Set the abort flag, which the background thread will see
            this.isAborted = true;
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
        /// Is the provided style one of the specially mapped ones for this artifact type
        /// </summary>
        /// <param name="styleName">The name of the style</param>
        /// <param name="artifactType"></param>
        /// <param name="mappedStyles">The list of mapped styles</param>
        /// <returns>True if this is a reserved style</returns>
        protected bool IsReservedStyleForArtifact(string styleName, int artifactType, Dictionary<SpiraRibbon.MappedStyleKeys, string> mappedStyles)
        {
            bool isReserved = false;
            //First see if we have the style name at all
            if (mappedStyles.ContainsValue(styleName))
            {
                //Next make sure it matches the artifact type
                if (artifactType == 1)
                {
                    //Make sure we have a requirements style
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent1] == styleName)
                    {
                        isReserved = true;
                    }
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent2] == styleName)
                    {
                        isReserved = true;
                    }
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent3] == styleName)
                    {
                        isReserved = true;
                    }
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent4] == styleName)
                    {
                        isReserved = true;
                    }
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent5] == styleName)
                    {
                        isReserved = true;
                    }
                }
                if (artifactType == 2)
                {
                    //Make sure we have a test case style
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.TestCase_Folder] == styleName)
                    {
                        isReserved = true;
                    }
                    if (mappedStyles[SpiraRibbon.MappedStyleKeys.TestCase_TestCase] == styleName)
                    {
                        isReserved = true;
                    }
                }
            }
            return isReserved;
        }

        /// <summary>
        /// Returns the numeric indent position for a given style name
        /// </summary>
        /// <param name="styleName">The name of the style</param>
        /// <returns></returns>
        public int GetIndentLevelForStyle (string styleName)
        {
            //Get the key that matches the style name
            SpiraRibbon.MappedStyleKeys key = SpiraRibbon.MappedStyleKeys.Requirement_Indent1;
            foreach (KeyValuePair<SpiraRibbon.MappedStyleKeys, string> kvp in this.MappedStyles)
            {
                if (kvp.Value == styleName)
                {
                    key = kvp.Key;
                    switch (key)
                    {
                        case SpiraRibbon.MappedStyleKeys.Requirement_Indent1:
                            return 0;
                        case SpiraRibbon.MappedStyleKeys.Requirement_Indent2:
                            return 1;
                        case SpiraRibbon.MappedStyleKeys.Requirement_Indent3:
                            return 2;
                        case SpiraRibbon.MappedStyleKeys.Requirement_Indent4:
                            return 3;
                        case SpiraRibbon.MappedStyleKeys.Requirement_Indent5:
                            return 4;
                    }
                }
            }
            return 0;
        }

        /// <summary>
        /// Determines if we have a test folder or test case based on its style
        /// </summary>
        /// <param name="styleName">The name of the style</param>
        /// <returns>True if a test folder</returns>
        public bool IsTestFolder(string styleName)
        {
            //Get the key that matches the style name
            SpiraRibbon.MappedStyleKeys key = SpiraRibbon.MappedStyleKeys.TestCase_TestCase;
            foreach (KeyValuePair<SpiraRibbon.MappedStyleKeys, string> kvp in this.MappedStyles)
            {
                if (kvp.Value == styleName)
                {
                    key = kvp.Key;
                    switch (key)
                    {
                        case SpiraRibbon.MappedStyleKeys.TestCase_Folder:
                            return true;
                        case SpiraRibbon.MappedStyleKeys.TestCase_TestCase:
                            return false;
                    }
                }
            }
            return false;
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
            output = output.Replace("&#x7", "");
            return output;
        }

        /// <summary>
        /// Exports data to Spira from the word document
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
            if (this.WordApplication == null)
            {
                throw new ApplicationException("Unable to get handle to Word application instance. Aborting Export");
            }

            
            //Make sure we have a document loaded
            try
            {
                if (this.WordApplication.ActiveDocument == null || this.WordApplication.Documents == null || this.WordApplication.Documents.Count == 0)
                {
                    throw new ApplicationException("No Word document is currently loaded. Please open the Word document you wish to export from.");
                }
            }
            catch (Exception)
            {
                throw new ApplicationException("No Word document is currently loaded. Please open the Word document you wish to export from.");
            }

            //Make sure that we have a text selection made, since we only import what's selected
            if (this.WordApplication.Selection == null || this.WordApplication.Selection.Paragraphs == null || this.WordApplication.Selection.Paragraphs.Count == 0)
            {
                throw new ApplicationException("No text in the Word document has been selected. You need to select the text from the document to be exported.");
            }

            //Start the background thread that performs the export
            ImportState importState = new ImportState();
            importState.ProjectId = projectId;
            importState.ArtifactTypeName = artifactTypeName;
            importState.WordSelection = this.WordApplication.Selection;
            importState.MappedStyles = this.MappedStyles;
            ParameterizedThreadStart threadStart = new ParameterizedThreadStart(this.Export_Process);
            Thread thread = new Thread(threadStart);
            //Needs to be an STA thread to access the Windows clipboard
            thread.TrySetApartmentState(ApartmentState.STA);
            thread.Start(importState);
        }

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
                    case "Test Cases":
                        exportCount = ExportTestCases(spiraImportExport, importState);
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
        /// Extracts the text and images out of a cell and returns as HTML markup
        /// </summary>
        /// <param name="cell">The cell being parsed</param>
        /// <returns>The equivalent HTML Markup</returns>
        /// <param name="foundImages">The dictionary of found images</param>
        /// <param name="imageId">The notional id given to each image</param>
        /// <param name="stepIndex">The index of the test step</param>
        private string ParseCell(Cell cell, Dictionary<SpiraImportExport.RemoteDocument, byte[]> foundImages, ref int imageId, int stepIndex)
        {
            string markup = "";
            
            //Need to loop through each paragraph in the cell
            foreach (Paragraph paragraph in cell.Range.Paragraphs)
            {
                //First add the text for the paragraph
                markup += paragraph.Range.Text;

                //Now loop through each image in the cell's paragraphs
                foreach (InlineShape inlineShape in paragraph.Range.InlineShapes)
                {
                    if (inlineShape != null)
                    {
                        string altText = inlineShape.AlternativeText;

                        //Need to copy into the clipboard
                        inlineShape.Select();
                        WordApplication.Selection.CopyAsPicture();
                        // get the object data from the clipboard
                        IDataObject ido = Clipboard.GetDataObject();
                        if (ido != null)
                        {
                            // can convert to bitmap?
                            if (ido.GetDataPresent(DataFormats.Bitmap))
                            {
                                // cast the data into a bitmap object
                                Bitmap bmp = (Bitmap)ido.GetData(DataFormats.Bitmap);
                                // validate that we got the data
                                if (bmp != null)
                                {
                                    //See which image format we have
                                    ImageFormat imageFormat = bmp.RawFormat;
                                    string fileExtension = "";
                                    if (imageFormat == ImageFormat.Bmp)
                                    {
                                        fileExtension = "bmp";
                                    }
                                    if (imageFormat == ImageFormat.Gif)
                                    {
                                        fileExtension = "gif";
                                    }
                                    if (imageFormat == ImageFormat.Jpeg)
                                    {
                                        fileExtension = "jpg";
                                    }
                                    if (imageFormat == ImageFormat.Png || imageFormat.Guid.ToString() == "b96b3caa-0728-11d3-9d7b-0000f81ef32e")
                                    {
                                        fileExtension = "png";
                                    }
                                    if (imageFormat == ImageFormat.Wmf)
                                    {
                                        fileExtension = "wmf";
                                    }
                                    if (imageFormat == ImageFormat.Emf)
                                    {
                                        fileExtension = "emf";
                                    }
                                    if (imageFormat == ImageFormat.Tiff)
                                    {
                                        fileExtension = "tiff";
                                    }
                                    //See if we have a known type and add as an attachment
                                    if (fileExtension != "")
                                    {
                                        byte[] rawData = (byte[])System.ComponentModel.TypeDescriptor.GetConverter(bmp).ConvertTo(bmp, typeof(byte[]));
                                        SpiraImportExport.RemoteDocument remoteDoc = new SpiraImportExport.RemoteDocument();
                                        remoteDoc.AuthorId = null; // Default
                                        remoteDoc.AttachedArtifacts = new SpiraImportExport.RemoteLinkedArtifact[1] { new SpiraImportExport.RemoteLinkedArtifact() };
                                        remoteDoc.AttachedArtifacts[0].ArtifactTypeId = 7;   //Test Step
                                        remoteDoc.AttachedArtifacts[0].ArtifactId = stepIndex; //Use the step index until we get the real id back from the database
                                        remoteDoc.FilenameOrUrl = "Inline" + imageId + "." + fileExtension;
                                        if (!String.IsNullOrEmpty(altText))
                                        {
                                            remoteDoc.Description = altText;
                                        }
                                        foundImages.Add(remoteDoc, rawData);
                                        imageId++;

                                        //Also add an img tag
                                        //For now we use a temporary URL, will replace with attachment id once we have it
                                        markup += "<img src=\"" + remoteDoc.FilenameOrUrl + "\" alt=\"" + altText + "\" />";
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return markup;
        }

        /// <summary>
        /// Adds a single Word table to the list of test steps
        /// </summary>
        /// <param name="table">The MS-Word table</param>
        /// <param name="foundImages">The collection of images</param>
        /// <param name="testSteps">The list of test steps</param>
        /// <param name="imageId">The image index</param>
        private void ExportTestStepsTable(Dictionary<SpiraImportExport.RemoteDocument, byte[]> foundImages, ref int imageId, List<SpiraImportExport.RemoteTestStep> testSteps, Table table)
        {
            //Now iterate through each of the table rows
            int stepIndex = 0;
            foreach (Row row in table.Rows)
            {
                //Ignore the first row, which will contain the headings
                if (!row.IsFirst)
                {
                    //See which columns have been specified (we get the index from the name of the column - last character)
                    int descriptionColumnIndex = -1;
                    int expectedResultColumnIndex = -1;
                    int sampleDataColumnIndex = -1;
                    int parsedColumnIndex;
                    
                    //Description
                    string columnName = this.MappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_Description];
                    if (columnName.Length > 0 && Int32.TryParse(columnName.Substring(columnName.Length-1,1), out parsedColumnIndex))
                    {
                        descriptionColumnIndex = parsedColumnIndex;
                    }
                    //Expected Result
                    columnName = this.MappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_ExpectedResult];
                    if (columnName.Length > 0 && Int32.TryParse(columnName.Substring(columnName.Length-1,1), out parsedColumnIndex))
                    {
                        expectedResultColumnIndex = parsedColumnIndex;
                    }
                    //Sample Data
                    columnName = this.MappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_SampleData];
                    if (columnName.Length > 0 && Int32.TryParse(columnName.Substring(columnName.Length-1,1), out parsedColumnIndex))
                    {
                        sampleDataColumnIndex = parsedColumnIndex;
                    }

                    //Now get the data from each cell
                    string description = "Not Specified";
                    if (row.Cells.Count >= descriptionColumnIndex)
                    {
                        description = ParseCell(row.Cells[descriptionColumnIndex], foundImages, ref imageId, stepIndex);
                    }

                    string expectedResult = "";
                    if (row.Cells.Count >= descriptionColumnIndex)
                    {
                        expectedResult = ParseCell(row.Cells[expectedResultColumnIndex], foundImages, ref imageId, stepIndex);
                    }

                    string sampleData = "";
                    if (row.Cells.Count >= descriptionColumnIndex)
                    {
                        sampleData = ParseCell(row.Cells[sampleDataColumnIndex], foundImages, ref imageId, stepIndex);
                    }

                    //Actually add the test step
                    SpiraImportExport.RemoteTestStep testStep = new SpiraWordAddIn.SpiraImportExport.RemoteTestStep();
                    testStep.Description = MakeXmlSafe(description);
                    testStep.Position = -1; //Insert at end
                    testStep.ExpectedResult = MakeXmlSafe(expectedResult);
                    testStep.SampleData = MakeXmlSafe(sampleData);
                    testSteps.Add(testStep);

                    //Increment the step index
                    stepIndex++;
                }
            }
        }

        /// <summary>
        /// Exports a single Word table into an XHTML table
        /// </summary>
        /// <param name="table">The MS-Word table</param>
        /// <param name="xmlDoc">The XML document</param>
        /// <returns>The XHTML table</returns>
        private ImportRequirementAdditionalInfo ExportTable(XmlDocument xmlDoc, Table table, XmlElement xhtmlRootNode, Dictionary<SpiraImportExport.RemoteDocument, byte[]> foundImages, SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState, string baseUrl, Selection selection, StreamWriter streamWriter)
        {

            const int COM_TRUE = -1;

            int exportCount = 0;
            int errorCount = 0;

            ImportRequirementAdditionalInfo expInfo = null;

            //First create the table
            XmlElement xmlTable = xmlDoc.CreateElement("table");

            //See if we should include a border
            string style = "border-collapse: collapse;";
            if (table.Borders.HasHorizontal)
            {
                style += "border-top: black 1px solid;";
                style += "border-bottom: black 1px solid;";
            }
            if (table.Borders.HasHorizontal)
            {
                style += "border-left: black 1px solid;";
                style += "border-right: black 1px solid;";
            }
            xmlTable.Attributes.Append(xmlDoc.CreateAttribute("style"));
            xmlTable.Attributes["style"].Value = style;

            // The table column widths are retained
            int maxAmountOfColums = table.Range.Columns.Count;
            float fullColumnsWidth = 0;
            List<float> basicColumnsWidths = new List<float>();
            List<float> basicColumnsLenghts = new List<float>();
            foreach (Row row in table.Rows)
            {
                if (row.Range.Columns.Count == maxAmountOfColums)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        basicColumnsWidths.Add(cell.Width);
                        fullColumnsWidth += cell.Width;
                        basicColumnsLenghts.Add(fullColumnsWidth);
                    }
                    break;
                }
            }

            //Now add the rows
            int rowCount = 1;
            foreach (Row row in table.Rows)
            {
                XmlElement xmlTableRow = xmlDoc.CreateElement("tr");
                xmlTable.AppendChild(xmlTableRow);

                int amountOfColumnsInRow = table.Rows[rowCount].Cells.Count;
                int columsUsed = 0;
                float widthUsed = 0;

                //Now add the cells
                foreach (Cell cell in row.Cells)
                {
                    XmlElement xmlTableCell;
                    if (row.IsFirst)
                    {
                        xmlTableCell = xmlDoc.CreateElement("th");
                    }
                    else
                    {
                        xmlTableCell = xmlDoc.CreateElement("td");
                    }
                    xmlTableRow.AppendChild(xmlTableCell);

                    // adding colspans if necessary
                    if (amountOfColumnsInRow == 1)
                    {
                        xmlTableCell.Attributes.Append(xmlDoc.CreateAttribute("colspan"));
                        xmlTableCell.Attributes["colspan"].Value = (maxAmountOfColums - columsUsed).ToString("G");
                    }
                    else if (amountOfColumnsInRow == maxAmountOfColums)
                    {

                    }
                    else
                    {
                        int colspan = 1;
                        widthUsed += cell.Width;
                        int block = GetNearestBlock(basicColumnsLenghts, widthUsed);
                        xmlTableCell.Attributes.Append(xmlDoc.CreateAttribute("colspan"));
                        if (block + (amountOfColumnsInRow - cell.ColumnIndex) > maxAmountOfColums)
                        {
                            colspan = maxAmountOfColums - (amountOfColumnsInRow - cell.ColumnIndex);
                        }
                        else
                        {
                            colspan = block - columsUsed;
                        }

                        if (colspan > 1)
                        {
                            xmlTableCell.Attributes.Append(xmlDoc.CreateAttribute("colspan"));
                            xmlTableCell.Attributes["colspan"].Value = colspan.ToString();
                        }
                        columsUsed += colspan;
                    }


                    //See if we should include a border
                    style = "";
                    if (cell.Borders.InsideLineStyle != WdLineStyle.wdLineStyleNone)
                    {
                        style += "border: black 1px solid;";
                    }
                    else if (cell.Borders.OutsideLineStyle != WdLineStyle.wdLineStyleNone)
                    {
                        style += "border: black 1px solid;";
                    }


                    // it is not necessary to set style for the bold/italic/underline - it simply doesn't work
                    //See if we have any styling to apply
                    if (cell.Range.Bold == COM_TRUE)
                    {
                        //style += "font-weight:bold;";
                    }
                    if (cell.Range.Italic == COM_TRUE)
                    {
                        //style += "font-style:italic;";
                    }
                    if (cell.Range.Underline != WdUnderline.wdUnderlineNone)
                    {
                        //style += "text-decoration:underline;";
                    }

                    xmlTableCell.Attributes.Append(xmlDoc.CreateAttribute("style"));
                    xmlTableCell.Attributes["style"].Value = style;

                    // the table internal part is processed using same steps as the normal text
                    expInfo = ExportRequirementInnerProcessing(cell.Range, xmlDoc, ref xmlTableCell, foundImages, spiraImportExport, importState, baseUrl, selection, streamWriter, table.ID);
                    exportCount = expInfo.ExportCount;
                    errorCount = expInfo.ErrorCount;
                }
                rowCount++;
            }

            expInfo.XmlTable = xmlTable;
            return expInfo;
        }

        /// <summary>
        /// Exports requirements (internal processing)
        /// this is moved out as a separate steps because normal text and text inside the table 
        /// shall be processed identically
        /// </summary>
        private ImportRequirementAdditionalInfo ExportRequirementInnerProcessing(Range range, XmlDocument xhtmlDoc, ref XmlElement xhtmlRootNode, Dictionary<SpiraImportExport.RemoteDocument, byte[]> foundImages, SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState, string baseUrl, Selection selection, StreamWriter streamWriter, string tableID)
        {
            int exportCount = 0;
            int progressCount = 0;
            int currentIndent = 0;
            int indentOffset = 0;
            int errorCount = 0;
            string requirementName = "";
            string markup = null;
            int imageId = 1;
            int listLevel = -1;
            XmlElement xmlPara = null;
            XmlElement xmlList = null;
            int tableId = 1;

            int imho = 0;
            bool skippable = false;

            foreach (Paragraph paragraph in range.Paragraphs)
            {
                try
                {
                    //See if we have a table
                    if (paragraph.Range.Tables.Count > 0)
                    {
                        int tablesProcessed = 0;
                        bool sameTable = false;
                        foreach (Table table in paragraph.Range.Tables)
                        {
                            //See if we've already imported the specified table
                            if (String.IsNullOrEmpty(table.ID))
                            {
                                tablesProcessed++;
                                table.ID = "added_" + tableId;
                                ImportRequirementAdditionalInfo callback = ExportTable(xhtmlDoc, table, xhtmlRootNode, foundImages, spiraImportExport, importState, baseUrl, selection, streamWriter);
                                XmlElement xmlTable = callback.XmlTable;
                                exportCount = callback.ExportCount;
                                errorCount = callback.ErrorCount;

                                xhtmlRootNode.AppendChild(xmlTable);
                                listLevel = -1;
                                xmlPara = null;
                                tableId++;
                            }
                            else
                            {
                                sameTable |= table.ID.Equals(tableID);
                            }
                        }
                        //Don't import the paragraph content as text since we already added the table
                        if (tablesProcessed > 0 || !sameTable) continue;
                    }
                    //See if we have a list style or not
                    ListFormat listFormat = paragraph.Range.ListFormat;
                    if (listFormat != null && listFormat.ListLevelNumber > 0)
                    {
                        if (listFormat.ListType == WdListType.wdListBullet || listFormat.ListType == WdListType.wdListPictureBullet)
                        {
                            //See what our existing list level was
                            if (listLevel == -1)
                            {
                                //Create a new UL with nested LI since no list element before
                                xmlList = xhtmlDoc.CreateElement("ul");
                                xmlPara = xhtmlDoc.CreateElement("li");
                                xhtmlRootNode.AppendChild(xmlList);
                                xmlList.AppendChild(xmlPara);
                            }
                            else if (xmlPara != null)
                            {
                                int currentListLevel = listFormat.ListLevelNumber;
                                if (currentListLevel > listLevel)
                                {
                                    //Create a new UL with nested LI under the old list
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        XmlElement xmlList2 = xhtmlDoc.CreateElement("ul");
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlList2);
                                        xmlList2.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel == listLevel)
                                {
                                    //Create just a new LI tag under the existing LI
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && xmlList.Name == "ul")
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel < listLevel)
                                {
                                    //The new list item is above the last one, so we need to traverse the
                                    //tree upwards the appropriate number of times
                                    int listOffset = listLevel - listFormat.ListLevelNumber;
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    for (int i = 0; i < listOffset; i++)
                                    {
                                        if (xmlList.ParentNode != null)
                                        {
                                            xmlList = (XmlElement)xmlList.ParentNode;
                                        }
                                    }
                                    //Create just a new LI tag under the existing LI
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                            }
                            listLevel = listFormat.ListLevelNumber;
                        }
                        else if (listFormat.ListType == WdListType.wdListListNumOnly || listFormat.ListType == WdListType.wdListMixedNumbering
                                    || listFormat.ListType == WdListType.wdListOutlineNumbering || listFormat.ListType == WdListType.wdListSimpleNumbering)
                        {
                            //
                            XmlElement foo = xhtmlDoc.CreateElement("p");
                            string paraText = null;
                            try
                            {

                                paraText = paragraph.Range.Text;
                                if (paraText.Equals("\r\a"))
                                {
                                    foo.InnerText = paragraph.Range.ListFormat.ListString;
                                }
                                else
                                {
                                    skippable = true;
                                    foo.InnerText = paragraph.Range.ListFormat.ListString + paraText;
                                }
                            }
                            catch (Exception e)
                            {
                                foo.InnerText = paragraph.Range.ListFormat.ListString;
                            }
                            xhtmlRootNode.AppendChild(foo);
                            xmlPara = foo;

                        }
                        else
                        {
                            //Just use a paragraph and reset list level
                            listLevel = -1;
                            xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                        }

                        //If we didn't create an LI, just create a paragraph instead (fail-safe)
                        if (xmlPara == null)
                        {
                            //Just use a paragraph and reset list level
                            listLevel = -1;
                            xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                        }
                    }
                    else
                    {
                        listLevel = -1;
                        xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                    }

                    //See if we have a matching style name
                    Style style = (Style)paragraph.get_Style();
                    if (style != null && IsReservedStyleForArtifact(style.NameLocal, 1, importState.MappedStyles))
                    {
                        //Export the requirement
                        if (requirementName != "")
                        {
                            markup = xhtmlRootNode.InnerXml;
                            SpiraImportExport.RemoteRequirement remoteRequirement = new SpiraImportExport.RemoteRequirement();
                            remoteRequirement.Name = MakeXmlSafe(requirementName).Trim();
                            remoteRequirement.Description = MakeXmlSafe(markup);
                            remoteRequirement.StatusId = 1; //Requested
                            remoteRequirement.RequirementTypeId = 2; //Feature
                            remoteRequirement = spiraImportExport.Requirement_Create1(remoteRequirement, indentOffset);

                            //Now any images
                            foreach (KeyValuePair<SpiraImportExport.RemoteDocument, byte[]> kvp in foundImages)
                            {
                                SpiraImportExport.RemoteDocument remoteDoc = kvp.Key;
                                remoteDoc.AttachedArtifacts[0].ArtifactId = remoteRequirement.RequirementId.Value;
                                remoteDoc = spiraImportExport.Document_AddFile(remoteDoc, kvp.Value);

                                //Now we need to update the temporary URLs with the real attachment id
                                if (remoteDoc.AttachmentId.HasValue)
                                {
                                    int attachmentId = remoteDoc.AttachmentId.Value;
                                    XmlNode xmlNode = xhtmlRootNode.SelectSingleNode(".//img[@src='" + remoteDoc.FilenameOrUrl + "']");
                                    if (xmlNode != null)
                                    {
                                        string attachmentUrl = spiraImportExport.System_GetArtifactUrl(NAVIGATION_ID_ATTACHMENTS, importState.ProjectId, attachmentId, "");
                                        xmlNode.Attributes["src"].Value = attachmentUrl.Replace("~", baseUrl);

                                        //Now update the requirement
                                        remoteRequirement = spiraImportExport.Requirement_RetrieveById(remoteRequirement.RequirementId.Value);
                                        markup = xhtmlRootNode.InnerXml;
                                        remoteRequirement.Description = MakeXmlSafe(markup);
                                        spiraImportExport.Requirement_Update(remoteRequirement);
                                    }
                                }
                            }
                            exportCount++;
                        }
                        //Reset the XML document and attachments
                        xhtmlDoc = new XmlDocument();
                        xhtmlRootNode = xhtmlDoc.CreateElement("html");
                        xhtmlDoc.AppendChild(xhtmlRootNode);
                        foundImages.Clear();
                        imageId = 1;

                        //Get the name of the next requirement
                        requirementName = paragraph.Range.Text;

                        //Determine the indent level
                        int indentLevel = GetIndentLevelForStyle(style.NameLocal);

                        //Need to determine the relative indent offset
                        indentOffset = indentLevel - currentIndent;
                        currentIndent = indentLevel;
                    }
                    else
                    {
                        //Add to the body of the requirement

                        //Set the paragraph-level styles
                        string styleText = "";
                        if (paragraph.Range.Bold == COM_TRUE)
                        {
                            styleText += "font-weight: bold;";
                        }
                        if (paragraph.Range.Italic == COM_TRUE)
                        {
                            styleText += "font-style: italic;";
                        }
                        if (paragraph.Range.Underline != WdUnderline.wdUnderlineNone && (int)paragraph.Range.Underline < 100)
                        {
                            styleText += "text-decoration: underline;";
                        }
                        xmlPara.Attributes.Append(xhtmlDoc.CreateAttribute("style"));
                        xmlPara.Attributes["style"].Value = styleText;

                        //Loop through each word in the paragraph
                        string paraText = "";
                        foreach (Range word in paragraph.Range.Words)
                        {
                            //Handle the various word styles (not paragraph ones)
                            if (word.Bold == COM_TRUE && paragraph.Range.Bold != COM_TRUE)
                            {
                                paraText += "<b>";
                            }
                            if (word.Italic == COM_TRUE && paragraph.Range.Italic != COM_TRUE)
                            {
                                paraText += "<i>";
                            }
                            if (word.Underline != WdUnderline.wdUnderlineNone && paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                            {
                                paraText += "<u>";
                            }
                            if (word.Font.Name != paragraph.Range.Font.Name)
                            {
                                //paraText += "<span style=\"font-family:" + word.Font.Name + "\">";
                            }

                            paraText += System.Security.SecurityElement.Escape(word.Text);

                            //Handle the various styles
                            if (word.Font.Name != paragraph.Range.Font.Name)
                            {
                                paraText += "</span>";
                            }
                            if (word.Underline != WdUnderline.wdUnderlineNone && paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                            {
                                paraText += "</u>";
                            }
                            if (word.Italic == COM_TRUE && paragraph.Range.Italic != COM_TRUE)
                            {
                                paraText += "</i>";
                            }
                            if (word.Bold == COM_TRUE && paragraph.Range.Bold != COM_TRUE)
                            {
                                paraText += "</b>";
                            }
                        }

                        if (skippable != true)
                        {
                            xmlPara.InnerXml = MakeXmlSafe(paraText);
                            if (xmlPara.InnerXml.Equals(""))
                            {
                                xmlPara.InnerXml = "<br />";
                            }
                        }
                        else
                        {
                            skippable = false;
                        }


                        //Now loop through each image in the paragraph
                        foreach (InlineShape inlineShape in paragraph.Range.InlineShapes)
                        {
                            if (inlineShape != null)
                            {
                                string altText = inlineShape.AlternativeText;

                                //Need to copy into the clipboard
                                inlineShape.Select();
                                WordApplication.Selection.CopyAsPicture();
                                // get the object data from the clipboard
                                IDataObject ido = Clipboard.GetDataObject();
                                if (ido != null)
                                {
                                    // can convert to bitmap?
                                    if (ido.GetDataPresent(DataFormats.Bitmap))
                                    {
                                        // cast the data into a bitmap object
                                        Bitmap bmp = (Bitmap)ido.GetData(DataFormats.Bitmap);
                                        // validate that we got the data
                                        if (bmp != null)
                                        {
                                            //See which image format we have
                                            ImageFormat imageFormat = bmp.RawFormat;
                                            string fileExtension = "";
                                            if (imageFormat == ImageFormat.Bmp)
                                            {
                                                fileExtension = "bmp";
                                            }
                                            if (imageFormat == ImageFormat.Gif)
                                            {
                                                fileExtension = "gif";
                                            }
                                            if (imageFormat == ImageFormat.Jpeg)
                                            {
                                                fileExtension = "jpg";
                                            }
                                            if (imageFormat == ImageFormat.Png || imageFormat.Guid.ToString() == "b96b3caa-0728-11d3-9d7b-0000f81ef32e")
                                            {
                                                fileExtension = "png";
                                            }
                                            if (imageFormat == ImageFormat.Wmf)
                                            {
                                                fileExtension = "wmf";
                                            }
                                            if (imageFormat == ImageFormat.Emf)
                                            {
                                                fileExtension = "emf";
                                            }
                                            if (imageFormat == ImageFormat.Tiff)
                                            {
                                                fileExtension = "tiff";
                                            }
                                            //See if we have a known type and add as an attachment
                                            if (fileExtension != "")
                                            {
                                                byte[] rawData = (byte[])System.ComponentModel.TypeDescriptor.GetConverter(bmp).ConvertTo(bmp, typeof(byte[]));
                                                SpiraImportExport.RemoteDocument remoteDoc = new SpiraImportExport.RemoteDocument();
                                                remoteDoc.AttachedArtifacts = new SpiraImportExport.RemoteLinkedArtifact[1] { new SpiraImportExport.RemoteLinkedArtifact() };
                                                remoteDoc.AttachedArtifacts[0].ArtifactTypeId = 1;   //Requirement
                                                remoteDoc.AuthorId = null; // Default
                                                remoteDoc.FilenameOrUrl = "Inline" + imageId + "." + fileExtension;
                                                if (!String.IsNullOrEmpty(altText))
                                                {
                                                    remoteDoc.Description = altText;
                                                }
                                                foundImages.Add(remoteDoc, rawData);
                                                imageId++;

                                                //Also add an img tag
                                                //For now we use a temporary URL, will replace with attachment id once we have it
                                                XmlElement imgElement = xhtmlDoc.CreateElement("img");
                                                imgElement.Attributes.Append(xhtmlDoc.CreateAttribute("src"));
                                                imgElement.Attributes.Append(xhtmlDoc.CreateAttribute("alt"));
                                                imgElement.Attributes["src"].Value = remoteDoc.FilenameOrUrl;
                                                imgElement.Attributes["alt"].Value = altText;
                                                xmlPara.AppendChild(imgElement);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Move to the next row and update progress bar
                    progressCount++;
                    this.UpdateProgress(progressCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (Exception exception)
                {
                    //Record the error in the log and add to the error count, then continue
                    streamWriter.WriteLine("Error During Export from Word > Spira: " + exception.Message + " (" + exception.StackTrace + ")");
                    streamWriter.Flush();
                    errorCount++;

                    //Reset the XML document and attachments
                    xhtmlDoc = new XmlDocument();
                    xhtmlRootNode = xhtmlDoc.CreateElement("html");
                    xhtmlDoc.AppendChild(xhtmlRootNode);
                    foundImages.Clear();
                    imageId = 1;
                }
            }

            return new ImportRequirementAdditionalInfo(exportCount, errorCount, requirementName);
        }

        /// <summary>
        /// Exports requirements
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportRequirements(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //First open up the textfile that we will log information to (used for debugging purposes)
            string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_WordImport.log";
            StreamWriter streamWriter = File.CreateText(debugFile);

            //Get a handle to the current Word selection
            Selection selection = importState.WordSelection;

            //The following XHTML document is used to store the parsed MS-Word content
            XmlDocument xhtmlDoc = new XmlDocument();
            XmlElement xhtmlRootNode = xhtmlDoc.CreateElement("html");
            xhtmlDoc.AppendChild(xhtmlRootNode);

            //The following collection stores any inline images
            Dictionary<SpiraImportExport.RemoteDocument, byte[]> foundImages = new Dictionary<SpiraImportExport.RemoteDocument, byte[]>();

            //Get the application base URL
            string baseUrl = spiraImportExport.System_GetWebServerUrl();

            //Set the progress max-value
            Range range = selection.FormattedText;
            this.UpdateProgress(0, range.Paragraphs.Count);

            //Find all the textual elements in the selection and iterate through the paragraphs
            int exportCount = 0;
            int progressCount = 0;
            int currentIndent = 0;
            int indentOffset = 0;
            int errorCount = 0;
            string requirementName = "";
            string markup;

            foreach (Table table in range.Tables)
            {
                table.ID = "";
            }

            ImportRequirementAdditionalInfo additionalInfo = ExportRequirementInnerProcessing(range, xhtmlDoc, ref xhtmlRootNode, foundImages, spiraImportExport, importState, baseUrl, selection, streamWriter, "");
            exportCount = additionalInfo.ExportCount;
            errorCount = additionalInfo.ErrorCount;
            requirementName = additionalInfo.RequirementName;

            /*
            int imageId = 1;
            int listLevel = -1;
            XmlElement xmlPara = null;
            XmlElement xmlList = null;
            int tableId = 1;
            foreach (Paragraph paragraph in range.Paragraphs)
            {
                try
                {
                    //See if we have a table
                    if (paragraph.Range.Tables.Count > 0)
                    {
                        foreach (Table table in paragraph.Range.Tables)
                        {
                            //See if we've already imported the specified table
                            if (String.IsNullOrEmpty(table.ID))
                            {
                                XmlElement xmlTable = ExportTable(xhtmlDoc, table);
                                xhtmlRootNode.AppendChild(xmlTable);
                                listLevel = -1;
                                xmlPara = null;
                                table.ID = "added_" + tableId;
                                tableId++;
                            }
                        }
                        //Don't import the paragraph content as text since we already added the table
                        continue;
                    }
                    //See if we have a list style or not
                    ListFormat listFormat = paragraph.Range.ListFormat;
                    if (listFormat != null && listFormat.ListLevelNumber > 0)
                    {
                        if (listFormat.ListType == WdListType.wdListBullet || listFormat.ListType == WdListType.wdListPictureBullet)
                        {
                            //See what our existing list level was
                            if (listLevel == -1)
                            {
                                //Create a new UL with nested LI since no list element before
                                xmlList = xhtmlDoc.CreateElement("ul");
                                xmlPara = xhtmlDoc.CreateElement("li");
                                xhtmlRootNode.AppendChild(xmlList);
                                xmlList.AppendChild(xmlPara);
                            }
                            else if (xmlPara != null)
                            {
                                int currentListLevel = listFormat.ListLevelNumber;
                                if (currentListLevel > listLevel)
                                {
                                    //Create a new UL with nested LI under the old list
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        XmlElement xmlList2 = xhtmlDoc.CreateElement("ul");
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlList2);
                                        xmlList2.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel == listLevel)
                                {
                                    //Create just a new LI tag under the existing LI
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && xmlList.Name == "ul")
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel < listLevel)
                                {
                                    //The new list item is above the last one, so we need to traverse the
                                    //tree upwards the appropriate number of times
                                    int listOffset = listLevel - listFormat.ListLevelNumber;
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    for (int i = 0; i < listOffset; i++)
                                    {
                                        if (xmlList.ParentNode != null)
                                        {
                                            xmlList = (XmlElement)xmlList.ParentNode;
                                        }
                                    }
                                    //Create just a new LI tag under the existing LI
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                            }
                            listLevel = listFormat.ListLevelNumber;
                        }
                        else if (listFormat.ListType == WdListType.wdListListNumOnly || listFormat.ListType == WdListType.wdListMixedNumbering
                                    || listFormat.ListType == WdListType.wdListOutlineNumbering || listFormat.ListType == WdListType.wdListSimpleNumbering)
                        {
                            //See what our existing list level was
                            if (listLevel == -1)
                            {
                                //Create a new OL with nested LI since no list element before
                                xmlList = xhtmlDoc.CreateElement("ol");
                                xmlPara = xhtmlDoc.CreateElement("li");
                                xhtmlRootNode.AppendChild(xmlList);
                                xmlList.AppendChild(xmlPara);
                            }
                            else if (xmlPara != null)
                            {
                                int currentListLevel = listFormat.ListLevelNumber;
                                if (currentListLevel > listLevel)
                                {
                                    //Create a new OL with nested LI under the old list
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        XmlElement xmlList2 = xhtmlDoc.CreateElement("ol");
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlList2);
                                        xmlList2.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel == listLevel)
                                {
                                    //Create just a new LI tag under the existing LI
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && xmlList.Name == "ol")
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel < listLevel)
                                {
                                    //The new list item is above the last one, so we need to traverse the
                                    //tree upwards the appropriate number of times
                                    int listOffset = listLevel - listFormat.ListLevelNumber;
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    for (int i = 0; i < listOffset; i++)
                                    {
                                        if (xmlList.ParentNode != null)
                                        {
                                            xmlList = (XmlElement)xmlList.ParentNode;
                                        }
                                    }
                                    //Create just a new LI tag under the existing LI
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                            }
                            listLevel = listFormat.ListLevelNumber;
                        }
                        else
                        {
                            //Just use a paragraph and reset list level
                            listLevel = -1;
                            xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                        }

                        //If we didn't create an LI, just create a paragraph instead (fail-safe)
                        if (xmlPara == null)
                        {
                            //Just use a paragraph and reset list level
                            listLevel = -1;
                            xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                        }
                    }
                    else
                    {
                        listLevel = -1;
                        xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                    }

                    //See if we have a matching style name
                    Style style = (Style)paragraph.get_Style();
                    if (style != null && IsReservedStyleForArtifact(style.NameLocal,1, importState.MappedStyles))
                    {
                        //Export the requirement
                        if (requirementName != "")
                        {
                            markup = xhtmlRootNode.InnerXml;
                            SpiraImportExport.RemoteRequirement remoteRequirement = new SpiraImportExport.RemoteRequirement();
                            remoteRequirement.Name = MakeXmlSafe(requirementName).Trim();
                            remoteRequirement.Description = MakeXmlSafe(markup);
                            remoteRequirement.StatusId = 1; //Requested
                            remoteRequirement.RequirementTypeId = 2; //Feature
                            remoteRequirement = spiraImportExport.Requirement_Create1(remoteRequirement, indentOffset);

                            //Now any images
                            foreach (KeyValuePair<SpiraImportExport.RemoteDocument, byte[]> kvp in foundImages)
                            {
                                SpiraImportExport.RemoteDocument remoteDoc = kvp.Key;
                                remoteDoc.AttachedArtifacts[0].ArtifactId = remoteRequirement.RequirementId.Value;
                                remoteDoc = spiraImportExport.Document_AddFile(remoteDoc, kvp.Value);

                                //Now we need to update the temporary URLs with the real attachment id
                                if (remoteDoc.AttachmentId.HasValue)
                                {
                                    int attachmentId = remoteDoc.AttachmentId.Value;
                                    XmlNode xmlNode = xhtmlRootNode.SelectSingleNode(".//img[@src='" + remoteDoc.FilenameOrUrl + "']");
                                    if (xmlNode != null)
                                    {
                                        string attachmentUrl = spiraImportExport.System_GetArtifactUrl(NAVIGATION_ID_ATTACHMENTS, importState.ProjectId, attachmentId, "");
                                        xmlNode.Attributes["src"].Value = attachmentUrl.Replace("~", baseUrl);

                                        //Now update the requirement
                                        remoteRequirement = spiraImportExport.Requirement_RetrieveById(remoteRequirement.RequirementId.Value);
                                        markup = xhtmlRootNode.InnerXml;
                                        remoteRequirement.Description = MakeXmlSafe(markup);
                                        spiraImportExport.Requirement_Update(remoteRequirement);
                                    }
                                }
                            }
                            exportCount++;
                        }
                        //Reset the XML document and attachments
                        xhtmlDoc = new XmlDocument();
                        xhtmlRootNode = xhtmlDoc.CreateElement("html");
                        xhtmlDoc.AppendChild(xhtmlRootNode);
                        foundImages.Clear();
                        imageId = 1;

                        //Get the name of the next requirement
                        requirementName = paragraph.Range.Text;

                        //Determine the indent level
                        int indentLevel = GetIndentLevelForStyle(style.NameLocal);

                        //Need to determine the relative indent offset
                        indentOffset = indentLevel - currentIndent;
                        currentIndent = indentLevel;
                    }
                    else
                    {
                        //Add to the body of the requirement
                        
                        //Set the paragraph-level styles
                        string styleText = "";
                        if (paragraph.Range.Bold == COM_TRUE)
                        {
                            styleText += "font-weight:bold;";
                        }
                        if (paragraph.Range.Italic == COM_TRUE)
                        {
                            styleText += "font-style:italic;";
                        }
                        if (paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                        {
                            styleText += "text-decoration:bold;";
                        }
                        if (paragraph.Range.Font.Name != selection.Range.Font.Name)
                        {
                            styleText += "font-family:" + paragraph.Range.Font.Name + ";";
                        }
                        xmlPara.Attributes.Append(xhtmlDoc.CreateAttribute("style"));
                        xmlPara.Attributes["style"].Value = styleText;

                        //Loop through each word in the paragraph
                        string paraText = "";
                        foreach (Range word in paragraph.Range.Words)
                        {
                            //Handle the various word styles (not paragraph ones)
                            if (word.Bold == COM_TRUE && paragraph.Range.Bold == COM_FALSE)
                            {
                                paraText += "<b>";
                            }
                            if (word.Italic == COM_TRUE && paragraph.Range.Italic == COM_FALSE)
                            {
                                paraText += "<i>";
                            }
                            if (word.Underline != WdUnderline.wdUnderlineNone && paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                            {
                                paraText += "<u>";
                            }
                            if (word.Font.Name != paragraph.Range.Font.Name)
                            {
                                paraText += "<span style=\"font-family:" + word.Font.Name + "\">";
                            }

                            paraText += System.Security.SecurityElement.Escape(word.Text);

                            //Handle the various styles
                            if (word.Font.Name != paragraph.Range.Font.Name)
                            {
                                paraText += "</span>";
                            }
                            if (word.Underline != WdUnderline.wdUnderlineNone && paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                            {
                                paraText += "</u>";
                            }
                            if (word.Italic == COM_TRUE && paragraph.Range.Italic == COM_FALSE)
                            {
                                paraText += "</i>";
                            }
                            if (word.Bold == COM_TRUE && paragraph.Range.Bold == COM_FALSE)
                            {
                                paraText += "</b>";
                            }
                        }
                        xmlPara.InnerXml = MakeXmlSafe(paraText);

                        //Now loop through each image in the paragraph
                        foreach (InlineShape inlineShape in paragraph.Range.InlineShapes)
                        {
                            if (inlineShape != null)
                            {
                                string altText = inlineShape.AlternativeText;

                                //Need to copy into the clipboard
                                inlineShape.Select();
                                WordApplication.Selection.CopyAsPicture();
                                // get the object data from the clipboard
                                IDataObject ido = Clipboard.GetDataObject();
                                if (ido != null)
                                {
                                    // can convert to bitmap?
                                    if (ido.GetDataPresent(DataFormats.Bitmap))
                                    {
                                        // cast the data into a bitmap object
                                        Bitmap bmp = (Bitmap)ido.GetData(DataFormats.Bitmap);
                                        // validate that we got the data
                                        if (bmp != null)
                                        {
                                            //See which image format we have
                                            ImageFormat imageFormat = bmp.RawFormat;
                                            string fileExtension = "";
                                            if (imageFormat == ImageFormat.Bmp)
                                            {
                                                fileExtension = "bmp";
                                            }
                                            if (imageFormat == ImageFormat.Gif)
                                            {
                                                fileExtension = "gif";
                                            }
                                            if (imageFormat == ImageFormat.Jpeg)
                                            {
                                                fileExtension = "jpg";
                                            }
                                            if (imageFormat == ImageFormat.Png || imageFormat.Guid.ToString() == "b96b3caa-0728-11d3-9d7b-0000f81ef32e")
                                            {
                                                fileExtension = "png";
                                            }
                                            if (imageFormat == ImageFormat.Wmf)
                                            {
                                                fileExtension = "wmf";
                                            }
                                            if (imageFormat == ImageFormat.Emf)
                                            {
                                                fileExtension = "emf";
                                            }
                                            if (imageFormat == ImageFormat.Tiff)
                                            {
                                                fileExtension = "tiff";
                                            }
                                            //See if we have a known type and add as an attachment
                                            if (fileExtension != "")
                                            {
                                                byte[] rawData = (byte[])System.ComponentModel.TypeDescriptor.GetConverter(bmp).ConvertTo(bmp, typeof(byte[]));
                                                SpiraImportExport.RemoteDocument remoteDoc = new SpiraImportExport.RemoteDocument();
                                                remoteDoc.AttachedArtifacts = new SpiraImportExport.RemoteLinkedArtifact[1] { new SpiraImportExport.RemoteLinkedArtifact() };
                                                remoteDoc.AttachedArtifacts[0].ArtifactTypeId = 1;   //Requirement
                                                remoteDoc.AuthorId = null; // Default
                                                remoteDoc.FilenameOrUrl = "Inline" + imageId + "." + fileExtension;
                                                if (!String.IsNullOrEmpty(altText))
                                                {
                                                    remoteDoc.Description = altText;
                                                }
                                                foundImages.Add(remoteDoc, rawData);
                                                imageId++;

                                                //Also add an img tag
                                                //For now we use a temporary URL, will replace with attachment id once we have it
                                                XmlElement imgElement = xhtmlDoc.CreateElement("img");
                                                imgElement.Attributes.Append(xhtmlDoc.CreateAttribute("src"));
                                                imgElement.Attributes.Append(xhtmlDoc.CreateAttribute("alt"));
                                                imgElement.Attributes["src"].Value = remoteDoc.FilenameOrUrl;
                                                imgElement.Attributes["alt"].Value = altText;
                                                xmlPara.AppendChild(imgElement);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Move to the next row and update progress bar
                    progressCount++;
                    this.UpdateProgress(progressCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (Exception exception)
                {
                    //Record the error in the log and add to the error count, then continue
                    streamWriter.WriteLine("Error During Export from Word > Spira: " + exception.Message + " (" + exception.StackTrace + ")");
                    streamWriter.Flush();
                    errorCount++;

                    //Reset the XML document and attachments
                    xhtmlDoc = new XmlDocument();
                    xhtmlRootNode = xhtmlDoc.CreateElement("html");
                    xhtmlDoc.AppendChild(xhtmlRootNode);
                    foundImages.Clear();
                    imageId = 1;
                }
            }*/

            //Insert the final requirement
            if (requirementName != "")
            {               
                markup = xhtmlRootNode.InnerXml;
                SpiraImportExport.RemoteRequirement remoteRequirement = new SpiraImportExport.RemoteRequirement();
                remoteRequirement.Name = MakeXmlSafe(requirementName).Trim();
                remoteRequirement.Description = MakeXmlSafe(markup);
                remoteRequirement.StatusId = 1; //Requested
                remoteRequirement = spiraImportExport.Requirement_Create1(remoteRequirement, indentOffset);
                exportCount++;

                //Now any images
                foreach (KeyValuePair<SpiraImportExport.RemoteDocument, byte[]> kvp in foundImages)
                {
                    SpiraImportExport.RemoteDocument remoteDoc = kvp.Key;
                    remoteDoc.AttachedArtifacts[0].ArtifactId = remoteRequirement.RequirementId.Value;
                    remoteDoc = spiraImportExport.Document_AddFile(remoteDoc, kvp.Value);

                    //Now we need to update the temporary URLs with the real attachment id
                    if (remoteDoc.AttachmentId.HasValue)
                    {
                        int attachmentId = remoteDoc.AttachmentId.Value;
                        XmlNode xmlNode = xhtmlRootNode.SelectSingleNode(".//img[@src='" + remoteDoc.FilenameOrUrl + "']");
                        if (xmlNode != null)
                        {
                            string attachmentUrl = spiraImportExport.System_GetArtifactUrl(NAVIGATION_ID_ATTACHMENTS, importState.ProjectId, attachmentId, "");
                            xmlNode.Attributes["src"].Value = attachmentUrl.Replace("~", baseUrl);

                            //Now update the requirement
                            remoteRequirement = spiraImportExport.Requirement_RetrieveById(remoteRequirement.RequirementId.Value);
                            markup = xhtmlRootNode.InnerXml;
                            remoteRequirement.Description = MakeXmlSafe(markup);
                            spiraImportExport.Requirement_Update(remoteRequirement);
                        }
                    }
                }
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                streamWriter.Close();
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please check the export error log to view the details.");
            }
            streamWriter.Close();

            return exportCount;
        }

        public int GetNearestBlock(List<float> basicColumnsWidths, float usedWidth)
        {
            float nearestRatio = float.MaxValue;
            int blockNumber = -1;
            for (int i = 0; i < basicColumnsWidths.Count; i++)
            {
                float currentRatio = Math.Abs(basicColumnsWidths[i] - usedWidth);
                if (currentRatio < nearestRatio)
                {
                    blockNumber = i + 1;
                    nearestRatio = currentRatio;
                }

            }
            return blockNumber;
        }

        /// <summary>
        /// Exports test cases
        /// </summary>
        /// <param name="importState">The state info</param>
        /// <param name="spiraImportExport">The web service proxy class</param>
        /// <returns>The number of exported artifacts</returns>
        private int ExportTestCases(SpiraImportExport.SoapServiceClient spiraImportExport, ImportState importState)
        {
            //First open up the textfile that we will log information to (used for debugging purposes)
            string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_WordImport.log";
            StreamWriter streamWriter = File.CreateText(debugFile);

            //Get a handle to the current Word selection
            Selection selection = importState.WordSelection;

            //Get the application base URL
            string baseUrl = spiraImportExport.System_GetWebServerUrl();

            //The following XHTML document is used to store the parsed MS-Word content
            XmlDocument xhtmlDoc = new XmlDocument();
            XmlElement xhtmlRootNode = xhtmlDoc.CreateElement("html");
            xhtmlDoc.AppendChild(xhtmlRootNode);

            //The following collection stores any inline images
            Dictionary<SpiraImportExport.RemoteDocument, byte[]> foundImages = new Dictionary<SpiraImportExport.RemoteDocument, byte[]>();

            //Set the progress max-value
            Range range = selection.FormattedText;
            this.UpdateProgress(0, range.Paragraphs.Count);

            //Find all the textual elements in the selection and iterate through the paragraphs
            int exportCount = 0;
            int progressCount = 0;
            int lastFolderId = -1;
            int errorCount = 0;
            string testCaseOrFolderName = "";
            string markup;
            int imageId = 1;
            int listLevel = -1;
            XmlElement xmlPara = null;
            XmlElement xmlList = null;
            int tableId = 1;
            bool isFolder = false;
            List<SpiraImportExport.RemoteTestStep> testSteps = new List<SpiraWordAddIn.SpiraImportExport.RemoteTestStep>();
            foreach (Paragraph paragraph in range.Paragraphs)
            {
                try
                {
                    //See if we have a table, if so then it should be imported as the test steps for the test case
                    if (paragraph.Range.Tables.Count > 0)
                    {
                        foreach (Table table in paragraph.Range.Tables)
                        {
                            //See if we've already imported the specified table
                            if (String.IsNullOrEmpty(table.ID))
                            {
                                ExportTestStepsTable(foundImages, ref imageId, testSteps, table);
                                listLevel = -1;
                                xmlPara = null;
                                table.ID = "added_" + tableId;
                                tableId++;
                            }
                        }
                        //Don't import the paragraph content as text since we already handled the table
                        continue;
                    }
                    //See if we have a list style or not
                    ListFormat listFormat = paragraph.Range.ListFormat;
                    if (listFormat != null && listFormat.ListLevelNumber > 0)
                    {
                        if (listFormat.ListType == WdListType.wdListBullet || listFormat.ListType == WdListType.wdListPictureBullet)
                        {
                            //See what our existing list level was
                            if (listLevel == -1)
                            {
                                //Create a new UL with nested LI since no list element before
                                xmlList = xhtmlDoc.CreateElement("ul");
                                xmlPara = xhtmlDoc.CreateElement("li");
                                xhtmlRootNode.AppendChild(xmlList);
                                xmlList.AppendChild(xmlPara);
                            }
                            else if (xmlPara != null)
                            {
                                int currentListLevel = listFormat.ListLevelNumber;
                                if (currentListLevel > listLevel)
                                {
                                    //Create a new UL with nested LI under the old list
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        XmlElement xmlList2 = xhtmlDoc.CreateElement("ul");
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlList2);
                                        xmlList2.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel == listLevel)
                                {
                                    //Create just a new LI tag under the existing LI
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && xmlList.Name == "ul")
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel < listLevel)
                                {
                                    //The new list item is above the last one, so we need to traverse the
                                    //tree upwards the appropriate number of times
                                    int listOffset = listLevel - listFormat.ListLevelNumber;
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    for (int i = 0; i < listOffset; i++)
                                    {
                                        if (xmlList.ParentNode != null)
                                        {
                                            xmlList = (XmlElement)xmlList.ParentNode;
                                        }
                                    }
                                    //Create just a new LI tag under the existing LI
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                            }
                            listLevel = listFormat.ListLevelNumber;
                        }
                        else if (listFormat.ListType == WdListType.wdListListNumOnly || listFormat.ListType == WdListType.wdListMixedNumbering
                                    || listFormat.ListType == WdListType.wdListOutlineNumbering || listFormat.ListType == WdListType.wdListSimpleNumbering)
                        {
                            //See what our existing list level was
                            if (listLevel == -1)
                            {
                                //Create a new OL with nested LI since no list element before
                                xmlList = xhtmlDoc.CreateElement("ol");
                                xmlPara = xhtmlDoc.CreateElement("li");
                                xhtmlRootNode.AppendChild(xmlList);
                                xmlList.AppendChild(xmlPara);
                            }
                            else if (xmlPara != null)
                            {
                                int currentListLevel = listFormat.ListLevelNumber;
                                if (currentListLevel > listLevel)
                                {
                                    //Create a new OL with nested LI under the old list
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        XmlElement xmlList2 = xhtmlDoc.CreateElement("ol");
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlList2);
                                        xmlList2.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel == listLevel)
                                {
                                    //Create just a new LI tag under the existing LI
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    if (xmlList != null && xmlList.Name == "ol")
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                                if (currentListLevel < listLevel)
                                {
                                    //The new list item is above the last one, so we need to traverse the
                                    //tree upwards the appropriate number of times
                                    int listOffset = listLevel - listFormat.ListLevelNumber;
                                    xmlList = (XmlElement)xmlPara.ParentNode;
                                    for (int i = 0; i < listOffset; i++)
                                    {
                                        if (xmlList.ParentNode != null)
                                        {
                                            xmlList = (XmlElement)xmlList.ParentNode;
                                        }
                                    }
                                    //Create just a new LI tag under the existing LI
                                    if (xmlList != null && (xmlList.Name == "ul" || xmlList.Name == "ol"))
                                    {
                                        xmlPara = xhtmlDoc.CreateElement("li");
                                        xmlList.AppendChild(xmlPara);
                                    }
                                    else
                                    {
                                        xmlPara = null;
                                    }
                                }
                            }
                            listLevel = listFormat.ListLevelNumber;
                        }
                        else
                        {
                            //Just use a paragraph and reset list level
                            listLevel = -1;
                            xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                        }

                        //If we didn't create an LI, just create a paragraph instead (fail-safe)
                        if (xmlPara == null)
                        {
                            //Just use a paragraph and reset list level
                            listLevel = -1;
                            xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                        }
                    }
                    else
                    {
                        listLevel = -1;
                        xmlPara = (XmlElement)xhtmlRootNode.AppendChild(xhtmlDoc.CreateElement("p"));
                    }

                    //See if we have a matching style name
                    Style style = (Style)paragraph.get_Style();
                    if (style != null && IsReservedStyleForArtifact(style.NameLocal, 2, importState.MappedStyles))
                    {
                        //Export the test folder/case
                        string safeName = MakeXmlSafe(testCaseOrFolderName).Trim();
                        if (safeName != "")
                        {
                            markup = xhtmlRootNode.InnerXml;
                            
                            //We need to store the new test step ids against their index
                            Dictionary<int,int> testStepIdLookups = new Dictionary<int,int>();
                            
                            //Create either a folder or test case
                            if (isFolder)
                            {
                                SpiraImportExport.RemoteTestCaseFolder remoteTestCaseFolder = new SpiraImportExport.RemoteTestCaseFolder();
                                remoteTestCaseFolder.Name = safeName;
                                remoteTestCaseFolder.Description = MakeXmlSafe(markup);
                                remoteTestCaseFolder = spiraImportExport.TestCase_CreateFolder(remoteTestCaseFolder);

                                if (lastFolderId == -1)
                                {
                                    remoteTestCaseFolder.ParentTestCaseFolderId = null;
                                }
                                else
                                {
                                    remoteTestCaseFolder.ParentTestCaseFolderId = lastFolderId;
                                }
                                lastFolderId = remoteTestCaseFolder.TestCaseFolderId.Value;
                            }
                            else
                            {
                                SpiraImportExport.RemoteTestCase remoteTestCase = new SpiraImportExport.RemoteTestCase();
                                remoteTestCase.TestCaseStatusId = /*Draft*/1;
                                remoteTestCase.TestCaseTypeId = /*Functional Test*/3;
                                remoteTestCase.Name = safeName;
                                remoteTestCase.Description = MakeXmlSafe(markup);

                                if (lastFolderId == -1)
                                {
                                    remoteTestCase.TestCaseFolderId = null;
                                }
                                else
                                {
                                    remoteTestCase.TestCaseFolderId = lastFolderId;
                                }
                                remoteTestCase = spiraImportExport.TestCase_Create(remoteTestCase);

                                //Add any test steps
                                if (testSteps.Count > 0)
                                {
                                    for (int i = 0; i < testSteps.Count; i++)
                                    {
                                        SpiraImportExport.RemoteTestStep newTestStep = spiraImportExport.TestCase_AddStep(testSteps[i], remoteTestCase.TestCaseId.Value);
                                        testStepIdLookups.Add(i, newTestStep.TestStepId.Value);
                                    }
                                }
                                testSteps.Clear();

                                //Now any images
                                remoteTestCase = spiraImportExport.TestCase_RetrieveById(remoteTestCase.TestCaseId.Value);
                                foreach (KeyValuePair<SpiraImportExport.RemoteDocument, byte[]> kvp in foundImages)
                                {
                                    SpiraImportExport.RemoteDocument remoteDoc = kvp.Key;
                                    if (remoteDoc.ArtifactTypeId == 2)
                                    {
                                        remoteDoc.AttachedArtifacts[0].ArtifactId = remoteTestCase.TestCaseId.Value;
                                        remoteDoc = spiraImportExport.Document_AddFile(remoteDoc, kvp.Value);

                                        //Now we need to update the temporary URLs with the real attachment id
                                        if (remoteDoc.AttachmentId.HasValue && !String.IsNullOrEmpty(remoteTestCase.Description))
                                        {
                                            int attachmentId = remoteDoc.AttachmentId.Value;
                                            if (remoteTestCase.Description.Contains("src=\"" + remoteDoc.FilenameOrUrl + "\""))
                                            {
                                                string attachmentUrl = spiraImportExport.System_GetArtifactUrl(NAVIGATION_ID_ATTACHMENTS, importState.ProjectId, attachmentId, "");
                                                attachmentUrl = attachmentUrl.Replace("~", baseUrl);
                                                remoteTestCase.Description = remoteTestCase.Description.Replace("src=\"" + remoteDoc.FilenameOrUrl + "\"", "src=\"" + attachmentUrl + "\"");
                                            }
                                        }
                                    }
                                    if (remoteDoc.ArtifactTypeId == 7)
                                    {
                                        //Lookup the ID from its index
                                        if (testStepIdLookups.ContainsKey(remoteDoc.AttachedArtifacts[0].ArtifactId))
                                        {
                                            remoteDoc.AttachedArtifacts[0].ArtifactId = testStepIdLookups[remoteDoc.AttachedArtifacts[0].ArtifactId];
                                            remoteDoc = spiraImportExport.Document_AddFile(remoteDoc, kvp.Value);

                                            //Now we need to update the temporary URLs with the real attachment id
                                            if (remoteDoc.AttachmentId.HasValue && remoteTestCase.TestSteps != null)
                                            {
                                                int attachmentId = remoteDoc.AttachmentId.Value;
                                                string attachmentUrl = spiraImportExport.System_GetArtifactUrl(NAVIGATION_ID_ATTACHMENTS, importState.ProjectId, attachmentId, "");
                                                attachmentUrl = attachmentUrl.Replace("~", baseUrl);
                                                foreach (SpiraImportExport.RemoteTestStep remoteTestStep in remoteTestCase.TestSteps)
                                                {
                                                    if (remoteTestStep.TestStepId.Value == remoteDoc.AttachedArtifacts[0].ArtifactId)
                                                    {
                                                        //Description
                                                        if (remoteTestStep.Description.Contains("src=\"" + remoteDoc.FilenameOrUrl + "\""))
                                                        {
                                                            remoteTestStep.Description = remoteTestStep.Description.Replace("src=\"" + remoteDoc.FilenameOrUrl + "\"", "src=\"" + attachmentUrl + "\"");
                                                        }
                                                        //Expected Result
                                                        if (!String.IsNullOrEmpty(remoteTestStep.ExpectedResult) && remoteTestStep.ExpectedResult.Contains("src=\"" + remoteDoc.FilenameOrUrl + "\""))
                                                        {
                                                            remoteTestStep.ExpectedResult = remoteTestStep.ExpectedResult.Replace("src=\"" + remoteDoc.FilenameOrUrl + "\"", "src=\"" + attachmentUrl + "\"");
                                                        }
                                                        //Sample Data
                                                        if (!String.IsNullOrEmpty(remoteTestStep.SampleData) && remoteTestStep.SampleData.Contains("src=\"" + remoteDoc.FilenameOrUrl + "\""))
                                                        {
                                                            remoteTestStep.SampleData = remoteTestStep.SampleData.Replace("src=\"" + remoteDoc.FilenameOrUrl + "\"", "src=\"" + attachmentUrl + "\"");
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                spiraImportExport.TestCase_Update(remoteTestCase);
                            }

                            exportCount++;
                        }
                        //Reset the XML document and attachments
                        xhtmlDoc = new XmlDocument();
                        xhtmlRootNode = xhtmlDoc.CreateElement("html");
                        xhtmlDoc.AppendChild(xhtmlRootNode);
                        foundImages.Clear();
                        imageId = 1;

                        //Get the name of the next test case/folder
                        testCaseOrFolderName = paragraph.Range.Text;

                        //Determine if its a folder or test case
                        isFolder = IsTestFolder(style.NameLocal) ;
                    }
                    else
                    {
                        //Add to the body of the test case
                        
                        //Set the paragraph-level styles
                        string styleText = "";
                        if (paragraph.Range.Bold == COM_TRUE)
                        {
                            styleText += "font-weight:bold;";
                        }
                        if (paragraph.Range.Italic == COM_TRUE)
                        {
                            styleText += "font-style:italic;";
                        }
                        if (paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                        {
                            styleText += "text-decoration:bold;";
                        }
                        if (paragraph.Range.Font.Name != selection.Range.Font.Name)
                        {
                            styleText += "font-family:" + paragraph.Range.Font.Name + ";";
                        }
                        xmlPara.Attributes.Append(xhtmlDoc.CreateAttribute("style"));
                        xmlPara.Attributes["style"].Value = styleText;

                        //Loop through each word in the paragraph
                        string paraText = "";
                        foreach (Range word in paragraph.Range.Words)
                        {
                            //Handle the various word styles (not paragraph ones)
                            if (word.Bold == COM_TRUE && paragraph.Range.Bold == COM_FALSE)
                            {
                                paraText += "<b>";
                            }
                            if (word.Italic == COM_TRUE && paragraph.Range.Italic == COM_FALSE)
                            {
                                paraText += "<i>";
                            }
                            if (word.Underline != WdUnderline.wdUnderlineNone && paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                            {
                                paraText += "<u>";
                            }
                            if (word.Font.Name != paragraph.Range.Font.Name)
                            {
                                paraText += "<span style=\"font-family:" + word.Font.Name + "\">";
                            }

                            paraText += System.Security.SecurityElement.Escape(word.Text);

                            //Handle the various styles
                            if (word.Font.Name != paragraph.Range.Font.Name)
                            {
                                paraText += "</span>";
                            }
                            if (word.Underline != WdUnderline.wdUnderlineNone && paragraph.Range.Underline != WdUnderline.wdUnderlineNone)
                            {
                                paraText += "</u>";
                            }
                            if (word.Italic == COM_TRUE && paragraph.Range.Italic == COM_FALSE)
                            {
                                paraText += "</i>";
                            }
                            if (word.Bold == COM_TRUE && paragraph.Range.Bold == COM_FALSE)
                            {
                                paraText += "</b>";
                            }
                        }
                        xmlPara.InnerXml = MakeXmlSafe(paraText);

                        //Now loop through each image in the paragraph
                        foreach (InlineShape inlineShape in paragraph.Range.InlineShapes)
                        {
                            if (inlineShape != null)
                            {
                                string altText = inlineShape.AlternativeText;

                                //Need to copy into the clipboard
                                inlineShape.Select();
                                WordApplication.Selection.CopyAsPicture();
                                // get the object data from the clipboard
                                IDataObject ido = Clipboard.GetDataObject();
                                if (ido != null)
                                {
                                    // can convert to bitmap?
                                    if (ido.GetDataPresent(DataFormats.Bitmap))
                                    {
                                        // cast the data into a bitmap object
                                        Bitmap bmp = (Bitmap)ido.GetData(DataFormats.Bitmap);
                                        // validate that we got the data
                                        if (bmp != null)
                                        {
                                            //See which image format we have
                                            ImageFormat imageFormat = bmp.RawFormat;
                                            string fileExtension = "";
                                            if (imageFormat == ImageFormat.Bmp)
                                            {
                                                fileExtension = "bmp";
                                            }
                                            if (imageFormat == ImageFormat.Gif)
                                            {
                                                fileExtension = "gif";
                                            }
                                            if (imageFormat == ImageFormat.Jpeg)
                                            {
                                                fileExtension = "jpg";
                                            }
                                            if (imageFormat == ImageFormat.Png || imageFormat.Guid.ToString() == "b96b3caa-0728-11d3-9d7b-0000f81ef32e")
                                            {
                                                fileExtension = "png";
                                            }
                                            if (imageFormat == ImageFormat.Wmf)
                                            {
                                                fileExtension = "wmf";
                                            }
                                            if (imageFormat == ImageFormat.Emf)
                                            {
                                                fileExtension = "emf";
                                            }
                                            if (imageFormat == ImageFormat.Tiff)
                                            {
                                                fileExtension = "tiff";
                                            }
                                            //See if we have a known type and add as an attachment
                                            if (fileExtension != "")
                                            {
                                                byte[] rawData = (byte[])System.ComponentModel.TypeDescriptor.GetConverter(bmp).ConvertTo(bmp, typeof(byte[]));
                                                SpiraImportExport.RemoteDocument remoteDoc = new SpiraImportExport.RemoteDocument();
                                                remoteDoc.AttachedArtifacts = new SpiraImportExport.RemoteLinkedArtifact[1] { new SpiraImportExport.RemoteLinkedArtifact() };
                                                remoteDoc.AttachedArtifacts[0].ArtifactTypeId = 2;   //Test Case
                                                remoteDoc.AuthorId = null; // Default
                                                remoteDoc.FilenameOrUrl = "Inline" + imageId + "." + fileExtension;
                                                if (!String.IsNullOrEmpty(altText))
                                                {
                                                    remoteDoc.Description = altText;
                                                }
                                                foundImages.Add(remoteDoc, rawData);
                                                imageId++;

                                                //Also add an img tag
                                                //For now we use a temporary URL, will replace with attachment id once we have it
                                                XmlElement imgElement = xhtmlDoc.CreateElement("img");
                                                imgElement.Attributes.Append(xhtmlDoc.CreateAttribute("src"));
                                                imgElement.Attributes.Append(xhtmlDoc.CreateAttribute("alt"));
                                                imgElement.Attributes["src"].Value = remoteDoc.FilenameOrUrl;
                                                imgElement.Attributes["alt"].Value = altText;
                                                xmlPara.AppendChild(imgElement);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Move to the next row and update progress bar
                    progressCount++;
                    this.UpdateProgress(progressCount, null);

                    //Check for abort condition
                    if (this.IsAborted)
                    {
                        throw new ApplicationException("Import aborted by user.");
                    }
                }
                catch (Exception exception)
                {
                    //Record the error in the log and add to the error count, then continue
                    streamWriter.WriteLine("Error During Export from Word > Spira: " + exception.Message + " (" + exception.StackTrace + ")");
                    streamWriter.Flush();
                    errorCount++;

                    //Reset the XML document and attachments
                    xhtmlDoc = new XmlDocument();
                    xhtmlRootNode = xhtmlDoc.CreateElement("html");
                    xhtmlDoc.AppendChild(xhtmlRootNode);
                    foundImages.Clear();
                    imageId = 1;
                }
            }

            //Insert the final test folder/case
            string finalSafeName = MakeXmlSafe(testCaseOrFolderName).Trim();
            if (finalSafeName != "")
            {               
                markup = xhtmlRootNode.InnerXml;

                //Create either a folder or test case
                if (isFolder)
                {
                    SpiraImportExport.RemoteTestCaseFolder remoteTestCaseFolder = new SpiraImportExport.RemoteTestCaseFolder();
                    remoteTestCaseFolder.Name = finalSafeName;
                    remoteTestCaseFolder.Description = MakeXmlSafe(markup);
                    remoteTestCaseFolder = spiraImportExport.TestCase_CreateFolder(remoteTestCaseFolder);

                    if (lastFolderId == -1)
                    {
                        remoteTestCaseFolder.ParentTestCaseFolderId = null;
                    }
                    else
                    {
                        remoteTestCaseFolder.ParentTestCaseFolderId = lastFolderId;
                    }
                    lastFolderId = remoteTestCaseFolder.TestCaseFolderId.Value;
                }
                else
                {
                    SpiraImportExport.RemoteTestCase remoteTestCase = new SpiraImportExport.RemoteTestCase();
                    remoteTestCase.TestCaseStatusId = /*Draft*/1;
                    remoteTestCase.TestCaseTypeId = /*Functional Test*/3;
                    remoteTestCase.Name = finalSafeName;
                    remoteTestCase.Description = MakeXmlSafe(markup);

                    if (lastFolderId == -1)
                    {
                        remoteTestCase.TestCaseFolderId = null;
                    }
                    else
                    {
                        remoteTestCase.TestCaseFolderId = lastFolderId;
                    }
                    remoteTestCase = spiraImportExport.TestCase_Create(remoteTestCase);

                    //Add any test steps
                    if (testSteps.Count > 0)
                    {
                        foreach (SpiraImportExport.RemoteTestStep testStep in testSteps)
                        {
                            spiraImportExport.TestCase_AddStep(testStep, remoteTestCase.TestCaseId.Value);
                        }
                    }

                    //Now any images
                    foreach (KeyValuePair<SpiraImportExport.RemoteDocument, byte[]> kvp in foundImages)
                    {
                        SpiraImportExport.RemoteDocument remoteDoc = kvp.Key;
                        remoteDoc.AttachedArtifacts[0].ArtifactId = remoteTestCase.TestCaseId.Value;
                        remoteDoc = spiraImportExport.Document_AddFile(remoteDoc, kvp.Value);

                        //Now we need to update the temporary URLs with the real attachment id
                        if (remoteDoc.AttachmentId.HasValue)
                        {
                            int attachmentId = remoteDoc.AttachmentId.Value;
                            XmlNode xmlNode = xhtmlRootNode.SelectSingleNode(".//img[@src='" + remoteDoc.FilenameOrUrl + "']");
                            if (xmlNode != null)
                            {
                                string attachmentUrl = spiraImportExport.System_GetArtifactUrl(NAVIGATION_ID_ATTACHMENTS, importState.ProjectId, attachmentId, "");
                                xmlNode.Attributes["src"].Value = attachmentUrl.Replace("~", baseUrl);

                                //Now update the test case
                                remoteTestCase = spiraImportExport.TestCase_RetrieveById(remoteTestCase.TestCaseId.Value);
                                markup = xhtmlRootNode.InnerXml;
                                remoteTestCase.Description = MakeXmlSafe(markup);
                                spiraImportExport.TestCase_Update(remoteTestCase);
                            }
                        }
                    }
                }
                exportCount++;
            }

            //Finally we need to unset the table id of the various tables in case we want to import again
            foreach (Table table in range.Tables)
            {
                table.ID = "";
            }

            //Only throw one message if an error occurred
            if (errorCount > 0)
            {
                streamWriter.Close();
                throw new ApplicationException("Export failed with " + errorCount + " errors. Please check the export error log to view the details.");
            }
            streamWriter.Close();

            return exportCount;
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
