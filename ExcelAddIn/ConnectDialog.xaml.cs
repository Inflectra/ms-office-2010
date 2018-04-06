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
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;

namespace SpiraExcelAddIn
{
    /// <summary>
    /// Interaction logic for ConnectDialog.xaml
    /// </summary>
    public partial class ConnectDialog : UserControl
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

        public ConnectDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Called when the connect dialog is first loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucConnectDialog_Loaded(object sender, RoutedEventArgs e)
        {
            //Populate the various images. Can't do in designer as VS2008 has bug
            this.imgCompanyName.Source = new BitmapImage(new Uri("Images/CompanyLogoSmall.gif", UriKind.Relative));
            this.imgProductLogo.Source = new BitmapImage(new Uri("Images/ExcelImportIcon.gif", UriKind.Relative));

            //Set the progress to not-indeterminate so that it doesn't move initially
            this.progressBar.IsIndeterminate = false;

            //Specify the values of the URL, UserName and Password
            this.txtUrl.Text = Configuration.Default.SpiraUrl;
            this.txtUsername.Text = Configuration.Default.SpiraUserName;
            if (String.IsNullOrEmpty(Configuration.Default.SpiraPassword))
            {
                this.txtPassword.Password = "";
                this.chkRemember.IsChecked = false;
            }
            else
            {
                this.txtPassword.Password = Configuration.Default.SpiraPassword;
                this.chkRemember.IsChecked = true;
            }
        }

        /// <summary>
        /// Handles clicks on the Connect button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            //First we need to validate the entry
            string spiraUrl = this.txtUrl.Text.Trim();
            string spiraUsername = this.txtUsername.Text.Trim();
            string spiraPassword = this.txtPassword.Password;
            if (!Uri.IsWellFormedUriString(spiraUrl, UriKind.Absolute))
            {
                MessageBox.Show("The Server URL entered is not a valid URL", "Connect to Server", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (spiraUrl.Contains(".aspx") || spiraUrl.Contains(".asmx") || spiraUrl.Contains(".svc"))
            {
                MessageBox.Show("The Server URL entered should only contain the server name, (port) and Virtual Directory (e.g. http://servername/SpiraTeam)", "Connect to Server", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (String.IsNullOrEmpty(spiraUsername))
            {
                MessageBox.Show("You need to enter a valid username", "Connect to Server", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (String.IsNullOrEmpty(spiraPassword))
            {
                MessageBox.Show("You need to enter a valid password", "Connect to Server", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (spiraUsername.Contains(" "))
            {
                MessageBox.Show("The Username field cannot contain any spaces", "Connect to Server", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            //Next we need to save the values entered by the user
            Configuration.Default.SpiraUrl = spiraUrl;
            Configuration.Default.SpiraUserName = spiraUsername;
            SpiraRibbon.SpiraPassword = spiraPassword;
            if (this.chkRemember.IsChecked.HasValue && this.chkRemember.IsChecked.Value)
            {
                Configuration.Default.SpiraPassword = spiraPassword;
            }
            else
            {
                Configuration.Default.SpiraPassword = "";
            }
            Configuration.Default.Save();

            Uri fullUri;
            if (!Importer.TryCreateFullUrl(spiraUrl, out fullUri))
            {
                MessageBox.Show("The Server URL entered is not a valid URL", "Connect to Server", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            //Display progress and start the authentication web service call
            this.progressBar.IsIndeterminate = true;
            this.btnConnect.IsEnabled = false;  //Prevent multiple connect calls

            SpiraImportExport.SoapServiceClient spiraImportExport = SpiraRibbon.CreateClient(fullUri);
            try
            {
                //Authenticate asynchronously
                spiraImportExport.Connection_AuthenticateCompleted += new EventHandler<SpiraExcelAddIn.SpiraImportExport.Connection_AuthenticateCompletedEventArgs>(spiraImportExport_Connection_AuthenticateCompleted);
                spiraImportExport.Connection_AuthenticateAsync(spiraUsername, spiraPassword);
            }
            catch (TimeoutException exception)
            {
                // Handle the timeout exception.
                this.progressBar.IsIndeterminate = false;
                spiraImportExport.Abort();
                MessageBox.Show("A timeout error occurred! (" + exception.Message + ")", "Timeout Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (CommunicationException exception)
            {
                // Handle the communication exception.
                this.progressBar.IsIndeterminate = false;
                spiraImportExport.Abort();
                MessageBox.Show("A communication error occurred! (" + exception.Message + ")", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Called when the Authentication call is completed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void spiraImportExport_Connection_AuthenticateCompleted(object sender, SpiraExcelAddIn.SpiraImportExport.Connection_AuthenticateCompletedEventArgs e)
        {
            //Stop the progress bar
            this.progressBar.IsIndeterminate = false;
            this.btnConnect.IsEnabled = true;  //Enable the connection button

            //See if we completed successfully or not
            if (e.Error != null)
            {
                //Error occurred
                MessageBox.Show("Unable to connect to Server. The error message was: '" + e.Error.Message + "'", "Connect Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (e.Result)
            {
                if (ParentElementHost != null)
                {
                    ConnectForm connectForm = (ConnectForm)ParentElementHost.FindForm();
                    if (connectForm != null)
                    {
                        //Authentication succeeded, so need to raise the Connect succeeded event
                        //so that the add-in knows to update the toolbar accordingly
                        connectForm.OnConnectSucceeded();
                
                        //Now close the dialog box
                        connectForm.Close();
                    }
                }
            }
            else
            {
                //Authentication failed
                MessageBox.Show("Unable to authenticate with Server. Please check the username and password and try again.", "Login Failed", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        /// <summary>
        /// Called when the Cancel button is clicked
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
    }
}
