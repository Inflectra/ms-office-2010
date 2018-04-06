using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.MSProject;

namespace SpiraProjectAddIn
{
    /// <summary>
    /// Contains the state information passed to the background Import process
    /// </summary>
    public class ImportState
    {
        /// <summary>
        /// The id of the project being imported from
        /// </summary>
        public int ProjectId
        {
            get;
            set;
        }

        /// <summary>
        /// Handle to the MS-Project file being imported into
        /// </summary>
        public Project MSProject
        {
            get;
            set;
        }
    }
}
