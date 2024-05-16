using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class Asignados_1 : Form
    {
        string path = string.Empty;
        public Asignados_1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Coleccion_Asignado_arr coleccion_Asignado_Arr = new Coleccion_Asignado_arr();
            //coleccion_Asignado_Arr.GenerarListado(@path);
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
