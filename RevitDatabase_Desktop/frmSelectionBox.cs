using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitDatabase
{
    public partial class frmSelectionBox : Form
    {
        public frmSelectionBox()
        {
            InitializeComponent();
        }

        private void frmSelectionBox_Load(object sender, EventArgs e)
        {
            btnOK.Enabled = false;
        }

        private void cbItems_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbItems.SelectedIndex == -1)
                btnOK.Enabled = false;
            else
                btnOK.Enabled = true;
        }
    }
}
