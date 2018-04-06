using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms.Integration;
using System.Windows.Interop;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace SpiraWordAddIn
{
    /// <summary>
    /// Interaction logic for ParametersDialog.xaml
    /// </summary>
    public partial class ParametersDialog : UserControl
    {
        #region Properties

        /// <summary>
        /// Gets the parent WinForm element host
        /// </summary>
        public ElementHost ParentElementHost
        {
            get
            {
                HwndSource wpfHandle = PresentationSource.FromVisual(this) as HwndSource;

                //the WPF control is hosted if the wpfHandle is not null
                if (wpfHandle == null)
                {
                    return null;
                }
                else
                {
                    ElementHost host = System.Windows.Forms.Control.FromChildHandle(wpfHandle.Handle) as ElementHost;
                    return host;
                }
            }
        }

        #endregion

        protected List<string> mappedCells = new List<string>() { "Column 1", "Column 2", "Column 3", "Column 4", "Column 5" };
        protected List<string> styles;
        protected Dictionary<SpiraRibbon.MappedStyleKeys, string> mappedStyles;

        public ParametersDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Populates the test step combo boxes
        /// </summary>
        /// <param name="styles">The list of styles in the current document</param>
        /// <param name="mappedStyles">the current style mappings</param>
        private void PopulateTestStepSettings(List<string> styles, Dictionary<SpiraRibbon.MappedStyleKeys, string> mappedStyles)
        {
            //See if we're mapping by style or by table cell
            if (this.radUseStyles.IsChecked.Value)
            {
                this.cboTestStepDescription.ItemsSource = styles;
                this.cboTestStepExectedResult.ItemsSource = styles;
                this.cboTestStepSampleData.ItemsSource = styles;
            }
            else
            {
                this.cboTestStepDescription.ItemsSource = mappedCells;
                this.cboTestStepExectedResult.ItemsSource = mappedCells;
                this.cboTestStepSampleData.ItemsSource = mappedCells;
            }

            //Set the current selections based on the mapped values
            if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.TestStep_Description))
            {
                string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_Description];
                if (this.cboTestStepDescription.Items.Contains(styleName))
                {
                    this.cboTestStepDescription.SelectedValue = styleName;
                }
            }
            if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.TestStep_ExpectedResult))
            {
                string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_ExpectedResult];
                if (this.cboTestStepExectedResult.Items.Contains(styleName))
                {
                    this.cboTestStepExectedResult.SelectedValue = styleName;
                }
            }
            if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.TestStep_SampleData))
            {
                string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_SampleData];
                if (this.cboTestStepSampleData.Items.Contains(styleName))
                {
                    this.cboTestStepSampleData.SelectedValue = styleName;
                }
            }
        }

        /// <summary>
        /// Sets up the page after loading
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucParametersDialog_Loaded(object sender, RoutedEventArgs e)
        {
            //Populate the various images. Can't do in designer as VS2008 has bug
            this.imgCompanyName.Source = new BitmapImage(new Uri("Images/CompanyLogoSmall.gif", UriKind.Relative));
            this.imgProductLogo.Source = new BitmapImage(new Uri("Images/WordImportIcon.gif", UriKind.Relative));

            //Need to get the list of available styles in the current document
            if (ParentElementHost != null)
            {
                ParametersForm parentForm = (ParametersForm)ParentElementHost.FindForm();
                if (parentForm != null)
                {
                    //Find the collection of mapped styles
                    if (parentForm.MappedStyles != null)
                    {
                        //Get a handle to the current document
                        Word.Document wordDocument = parentForm.WordDocument;
                        this.mappedStyles = parentForm.MappedStyles;

                        //Get the list of styles and convert to a simple .NET list (easier to manage)
                        this.styles = new List<string>();
                        foreach (Word.Style style in wordDocument.Styles)
                        {
                            styles.Add(style.NameLocal);
                        }

                        //Populate the combo-boxes
                        this.cboReqStyle1.ItemsSource = styles;
                        this.cboReqStyle2.ItemsSource = styles;
                        this.cboReqStyle3.ItemsSource = styles;
                        this.cboReqStyle4.ItemsSource = styles;
                        this.cboReqStyle5.ItemsSource = styles;
                        this.cboTestCaseFolder.ItemsSource = styles;
                        this.cboTestCaseName.ItemsSource = styles;

                        //Populate the test steps based on the radio button
                        PopulateTestStepSettings(styles, mappedStyles);

                        //Set the selections
                        //Req Indent 1
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.Requirement_Indent1))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent1];
                            if (this.cboReqStyle1.Items.Contains(styleName))
                            {
                                this.cboReqStyle1.SelectedValue = styleName;
                            }
                        }
                        //Req Indent 2
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.Requirement_Indent2))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent2];
                            if (this.cboReqStyle2.Items.Contains(styleName))
                            {
                                this.cboReqStyle2.SelectedValue = styleName;
                            }
                        }
                        //Req Indent 3
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.Requirement_Indent3))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent3];
                            if (this.cboReqStyle3.Items.Contains(styleName))
                            {
                                this.cboReqStyle3.SelectedValue = styleName;
                            }
                        }
                        //Req Indent 4
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.Requirement_Indent4))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent4];
                            if (this.cboReqStyle4.Items.Contains(styleName))
                            {
                                this.cboReqStyle4.SelectedValue = styleName;
                            }
                        }
                        //Req Indent 5
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.Requirement_Indent5))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent5];
                            if (this.cboReqStyle5.Items.Contains(styleName))
                            {
                                this.cboReqStyle5.SelectedValue = styleName;
                            }
                        }

                        //Test Case Folder
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.TestCase_Folder))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.TestCase_Folder];
                            if (this.cboTestCaseFolder.Items.Contains(styleName))
                            {
                                this.cboTestCaseFolder.SelectedValue = styleName;
                            }
                        }
                        //Test Case Name
                        if (mappedStyles.ContainsKey(SpiraRibbon.MappedStyleKeys.TestCase_TestCase))
                        {
                            string styleName = mappedStyles[SpiraRibbon.MappedStyleKeys.TestCase_TestCase];
                            if (this.cboTestCaseName.Items.Contains(styleName))
                            {
                                this.cboTestCaseName.SelectedValue = styleName;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Handles clicks on the Cancel button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //Just close the dialog box
            if (ParentElementHost != null)
            {
                System.Windows.Forms.Form parentForm = ParentElementHost.FindForm();
                if (parentForm != null)
                {
                    parentForm.Close();
                }
            }
        }

        /// <summary>
        /// Handles clicks on the Update button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (ParentElementHost != null)
            {
                ParametersForm parentForm = (ParametersForm)ParentElementHost.FindForm();
                if (parentForm != null)
                {
                    //Find the collection of mapped styles
                    if (parentForm.MappedStyles != null)
                    {

                        Dictionary<SpiraRibbon.MappedStyleKeys, string> mappedStyles = parentForm.MappedStyles;
                        //Update the mappings from the various controls
                        
                        //Req Indent 1
                        if (this.cboReqStyle1.SelectedValue != null && this.cboReqStyle1.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent1] = (string)this.cboReqStyle1.SelectedValue;
                        }
                        //Req Indent 2
                        if (this.cboReqStyle1.SelectedValue != null && this.cboReqStyle2.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent2] = (string)this.cboReqStyle2.SelectedValue;
                        }
                        //Req Indent 3
                        if (this.cboReqStyle1.SelectedValue != null && this.cboReqStyle3.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent3] = (string)this.cboReqStyle3.SelectedValue;
                        }
                        //Req Indent 4
                        if (this.cboReqStyle1.SelectedValue != null && this.cboReqStyle4.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent4] = (string)this.cboReqStyle4.SelectedValue;
                        }
                        //Req Indent 5
                        if (this.cboReqStyle1.SelectedValue != null && this.cboReqStyle5.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.Requirement_Indent5] = (string)this.cboReqStyle5.SelectedValue;
                        }
                        //Test Case Folder
                        if (this.cboTestCaseFolder.SelectedValue != null && this.cboTestCaseFolder.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.TestCase_Folder] = (string)this.cboTestCaseFolder.SelectedValue;
                        }
                        //Test Case Name
                        if (this.cboTestCaseName.SelectedValue != null && this.cboTestCaseName.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.TestCase_TestCase] = (string)this.cboTestCaseName.SelectedValue;
                        }
                        //Test Step Description
                        if (this.cboTestStepDescription.SelectedValue != null && this.cboTestStepDescription.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_Description] = (string)this.cboTestStepDescription.SelectedValue;
                        }
                        //Test Step Expected Result
                        if (this.cboTestStepExectedResult.SelectedValue != null && this.cboTestStepExectedResult.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_ExpectedResult] = (string)this.cboTestStepExectedResult.SelectedValue;
                        }
                        //Test Step Sample Data
                        if (this.cboTestStepSampleData.SelectedValue != null && this.cboTestStepSampleData.SelectedValue.ToString() != "")
                        {
                            mappedStyles[SpiraRibbon.MappedStyleKeys.TestStep_SampleData] = (string)this.cboTestStepSampleData.SelectedValue;
                        }
                    }

                    //Next close the dialog box
                    parentForm.Close();
                }
            }
        }

        /// <summary>
        /// Change the list of mapped test step styles/columns
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radUseTables_Checked(object sender, RoutedEventArgs e)
        {
            if (this.radUseStyles != null && this.radUseTables != null)
            {
                PopulateTestStepSettings(styles, this.mappedStyles);
            }
        }

        /// <summary>
        /// Change the list of mapped test step styles/columns
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radUseStyles_Checked(object sender, RoutedEventArgs e)
        {
            if (this.radUseStyles != null && this.radUseTables != null)
            {
                PopulateTestStepSettings(styles, this.mappedStyles);
            }
        }
    }
}
