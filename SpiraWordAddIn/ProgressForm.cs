using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpiraWordAddIn
{
    public partial class ProgressForm : Form
    {
        protected bool displayConfirmation = true;

        /// <summary>
        /// Get/sets the current progress value
        /// </summary>
        public int ProgressValue
        {
            get
            {
                return this.progressBar1.Value;
            }
            set
            {
                this.progressBar1.Value = value;
            }
        }

        /// <summary>
        /// Get/sets the maximum progress value
        /// </summary>
        public int ProgressMaximumValue
        {
            get
            {
                return this.progressBar1.Maximum;
            }
            set
            {
                this.progressBar1.Maximum = value;
            }
        }

        /// <summary>
        /// Gets/sets the caption for the progress form
        /// </summary>
        public string Caption
        {
            get
            {
                return this.lblTitle.Text;
            }
            set
            {
                this.lblTitle.Text = value;
            }
        }

        public ProgressForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Sets up the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ProgressForm_Load(object sender, EventArgs e)
        {
            //By default, display confirmation when close attempted
            displayConfirmation = true;
        }

        /// <summary>
        /// Closes the form without displaying a confirmation box
        /// </summary>
        /// <remarks>Used to force a close without a message</remarks>
        public new void Close()
        {
            displayConfirmation = false;
            base.Close();
        }

        /// <summary>
        /// Closes the progress form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnHide_Click(object sender, EventArgs e)
        {
            //Call the base method so that we get a confirmation box
            displayConfirmation = true;
            base.Close();
        }

        /// <summary>
        /// If the user tries to close warn them that it will stop the import
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ProgressForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (displayConfirmation)
            {
                DialogResult result = MessageBox.Show("This will stop the current import/export. Are you sure you want to proceed?", "Stop Import/Export", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.No)
                {
                    //Cancel the event
                    e.Cancel = true;
                }
            }
        }
    }
}
