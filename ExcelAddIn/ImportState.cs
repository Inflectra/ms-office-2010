using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;

namespace SpiraExcelAddIn
{
    /// <summary>
    /// Contains the state information passed to the background Import process
    /// </summary>
    public class ImportState
    {
        /// <summary>
        /// The type of data being imported
        /// </summary>
        public string ArtifactTypeName
        {
            get;
            set;
        }

        /// <summary>
        /// The id of the project being imported from
        /// </summary>
        public int ProjectId
        {
            get;
            set;
        }

        /// <summary>
        /// Handle to the excel worksheet being imported into
        /// </summary>
        public Worksheet ExcelWorksheet
        {
            get;
            set;
        }

        /// <summary>
        /// Handle to the excel worksheet that contains the lookups
        /// </summary>
        public Worksheet LookupWorksheet
        {
            get;
            set;
        }
    }
}
