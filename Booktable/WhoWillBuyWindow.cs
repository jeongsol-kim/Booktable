using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Booktable
{
    public partial class WhoWillBuyWindow : Form
    {
        public WhoWillBuyWindow()
        {
            InitializeComponent();
        }

        private void WhoWillBuyWindow_Load(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public string who { get; set; }
        public string optional { get; set; }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.who = this.textBox1.Text;
            this.optional = this.textBox2.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
