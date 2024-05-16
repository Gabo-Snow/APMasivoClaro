using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class Ventana_Compartidos : Form
    {
        string aux_path = string.Empty;
        Coleccion_Activacion coleccion_Activacion = new Coleccion_Activacion();
        Coleccion_Cartera_Empresarial _Cartera_Empresarial = new Coleccion_Cartera_Empresarial();
        Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
        List<Cartera_Empresarial> cartera_Empresarials_sice;
        List<Cartera> carteras_pablo;
        List<Activacion> activacions;
        List<string> tipos_ = new List<string>();
        List<Ajustes> ajustes;
        string carpeta_final = string.Empty;
        string carpeta_ruta = string.Empty;
        string carpeta_extraidas = string.Empty;
        string carpeta_respaldo = string.Empty;
        string carpeta_asignar = string.Empty;
        string carpeta_reportar = string.Empty;
        string quedaran_aqui_txt = string.Empty;
        public Ventana_Compartidos()
        {
            InitializeComponent();
            ejecutar_btn.Visible = true;
        }



        private void ejecutar_btn_Click_1(object sender, EventArgs e)
        {
            double algo =  100000000; //son los dummy
            for (int i = 0; i < algo; i++)
            {
                Console.WriteLine("Imprimiendo al aire");
            }
            Console.WriteLine("Termine!!");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            carpeta_txt.Text = aux_path;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            dummy_txt.Text = aux_path;
        }


    }
}
