using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Word;

namespace SpiraWordAddIn
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
        /// Handle to the MS-Word document being imported into
        /// </summary>
        public Selection WordSelection
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
    }
}
