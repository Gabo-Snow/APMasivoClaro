using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class Corporaciones : Form
    {
        Coleccion_Cargos_Corporaciones coleccion_Cargos_Corporaciones = new Coleccion_Cargos_Corporaciones();
        List<Cargos_Cuenta_Corporaciones>  cargos_Cuenta_Corporaciones = new List<Cargos_Cuenta_Corporaciones>();
        List<string> tipos_ = new List<string>();
        public Corporaciones()
        {
            InitializeComponent();
        }

        private void btn_Guardar_Click(object sender, EventArgs e)
        {
            ejecutar_btn.Visible = true;
        }

        private void analizar_Btn_Click(object sender, EventArgs e)
        {
            if (cargos_cuentas_txt.Text.Equals(""))
            {
                string message2 = "Ingrese las rutas de los archivos";
                MessageBox.Show(message2);
            }
            else
            {
                cargos_Cuenta_Corporaciones = coleccion_Cargos_Corporaciones.GenerarListado(cargos_cuentas_txt.Text);

                var coleccion_empresas2 = cargos_Cuenta_Corporaciones.Select(x => x.DESCRIPTION).Distinct().ToList();
                List<string> lista_tipos_descripcion = new List<string>();
                foreach (var item in coleccion_empresas2)
                {
                    lista_tipos_descripcion.Add(item.ToUpper());
                }                

                DataTable dtCargos_fijos = new DataTable();
                // add column to datatable  

                dtCargos_fijos.Columns.Add("Tipo", typeof(string));
                dtCargos_fijos.Columns.Add("check", typeof(bool));

                DataTable dtRemanentes = new DataTable();
                // add column to datatable  

                dtRemanentes.Columns.Add("Tipo", typeof(string));
                dtRemanentes.Columns.Add("check", typeof(bool));

                DataTable dtEquipos = new DataTable();
                // add column to datatable  

                dtEquipos.Columns.Add("Tipo", typeof(string));
                dtEquipos.Columns.Add("check", typeof(bool));

                foreach (var item in lista_tipos_descripcion.Distinct())
                {
                    if (lista_tipos_descripcion.Equals(""))
                    {

                    }
                    else
                    {
                        if (lista_tipos_descripcion.Equals("DESCRIPTION"))
                        {

                        }
                        else
                        {
                            dtCargos_fijos.Rows.Add(new Object[] { item, false });
                            dtRemanentes.Rows.Add(new Object[] { item, false });
                            dtEquipos.Rows.Add(new Object[] { item, false });
                        }
                        
                    }
                    
                }

                Cargos_Fijos_dgv.DataSource = dtCargos_fijos;
                Remanentes_dgv.DataSource = dtRemanentes;
                Equipos_dgv.DataSource = dtEquipos;
                //Cargos_Fijos_dgv_1.DataSource = dtEmp;

                Cargos_Fijos_dgv.Visible = true;
                Remanentes_dgv.Visible = true;
                Equipos_dgv.Visible = true;
                btn_Guardar.Visible = true;
            }
        }
        string aux_path = string.Empty;
        private void empresarial_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargos_cuentas_txt.Text = aux_path;
        }

        private void cartera_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            detenciones_txt.Text = aux_path;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            quedara_aqui_txt.Text = aux_path;
        }

        private void ejecutar_btn_Click(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(60000);
            File.Copy(@"c:\test.txt", @"c:\test\foo.txt");
        }
    }
}
