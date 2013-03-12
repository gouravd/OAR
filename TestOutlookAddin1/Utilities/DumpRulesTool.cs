using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookAttachmentReminder.Utilities
{
    public partial class DumpRulesTool : Form
    {
        public DumpRulesTool()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
        }

        private void DumpRulesTool_Load(object sender, EventArgs e)
        {

        }

        private void btnGenerateStream_Click(object sender, EventArgs e)
        {

        }
    }
}
