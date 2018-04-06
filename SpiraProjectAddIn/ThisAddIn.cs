﻿using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.ServiceModel;
using System.ServiceModel.Description;

namespace SpiraProjectAddIn
{
    /// <summary>
    /// Adds a Spira toolbar to MS-Project versions 2010 or higher
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Called when the Add-In starts up
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Do nothing - the ribbon is loaded automatically by VSTO
        }

        /// <summary>
        /// Called when the Add-in shuts down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //Do nothing
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
