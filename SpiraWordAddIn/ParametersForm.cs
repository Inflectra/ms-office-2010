using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace SpiraWordAddIn
{
    public partial class ParametersForm : Form
    {
        #region Properties

        /// <summary>
        /// Reference to the current word document
        /// </summary>
        public Word.Document WordDocument
        {
            get;
            set;
        }

        /// <summary>
        /// The mapped styles
        /// </summary>
        public Dictionary<SpiraRibbon.MappedStyleKeys, string> MappedStyles
        {
            get;
            set;
        }

        #endregion

        public ParametersForm()
        {
            InitializeComponent();
        }
    }
}
