using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MLG_Fetch
{
    public partial class Welcome : Form
    {
        public Welcome()
        {
            InitializeComponent();
            checkBox1.Checked = Properties.Settings.Default.report_concent;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.first_run = false;
            Properties.Settings.Default.report_concent = checkBox1.Checked;
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                button1.Text = "Продолжить";
            } else
            {
                button1.Text = "Продолжить без соглашения";
            }
        }
    }
}
