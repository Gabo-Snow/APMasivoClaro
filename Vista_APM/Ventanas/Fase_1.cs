using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ventana_APM.Ventanas
{
    public partial class Fase_1 : Form
    {
        string path = string.Empty;
        public Fase_1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                path = fbd.SelectedPath.ToString();
            path_txtbox.Text = path;

        }
    }
}
