using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace WTools
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UserControl1 userControl1 = new UserControl1();
            panel2.Controls.Clear();
            panel2.Text = button1.Text;
            panel2.Controls.Add(userControl1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UserControl3 userControl3 = new UserControl3();
            panel2.Controls.Clear();
            panel2.Text = button1.Text;
            panel2.Controls.Add(userControl3);
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            UserControl2 userControl2 = new UserControl2();
            panel2.Controls.Clear();
            panel2.Text = button1.Text;
            panel2.Controls.Add(userControl2);
        }
    }
}
