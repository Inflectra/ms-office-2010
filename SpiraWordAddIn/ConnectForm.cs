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
    public partial class ConnectForm : Form
    {
        //This event is called when the connection succeeds
        public event System.EventHandler ConnectSucceeded;

        /// <summary>
        /// Constructor
        /// </summary>
        public ConnectForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Called to raise the OnConnectSucceeded event
        /// </summary>
        public void OnConnectSucceeded() 
        {
            if (ConnectSucceeded != null)
            {
                ConnectSucceeded(this, new EventArgs());
            }
        }
    }
}
