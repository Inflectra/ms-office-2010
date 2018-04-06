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

namespace SpiraExcelAddIn
{
    /// <summary>
    /// Interaction logic for OptionsDialog.xaml
    /// </summary>
    public partial class OptionsDialog : UserControl
    {
        public OptionsDialog()
        {
            InitializeComponent();
        }

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

        /// <summary>
        /// Called when Update clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            //Update the settings and close
            if (this.chkRemoveFormatting.IsChecked.HasValue)
            {
                Configuration.Default.StripRichText = this.chkRemoveFormatting.IsChecked.Value;
            }
            Configuration.Default.TestRunDate = this.datTestRunExport.SelectedDate;
            Configuration.Default.Save();

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
        /// Called when Cancel clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //Close the dialog box
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
        /// Called when the user control first loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //Set the checkbox
            this.chkRemoveFormatting.IsChecked = Configuration.Default.StripRichText;
            this.datTestRunExport.SelectedDate = Configuration.Default.TestRunDate;
        }
    }
}
