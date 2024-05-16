using APM;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.CONTROLADOR;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class Ajuste_CFM : Form
    {
        string aux_path = string.Empty;
        List<Fide_Cob_Pcs> cargosx_pcs1 = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> cargosx_pcs2 = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> cargosx_pcs3 = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> cargosx_pcs4 = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> cargosx_pcs5 = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> cargosx_pcs6 = new List<Fide_Cob_Pcs>();
        List<CFM_AJUSTES> ajuste_CFMs = new List<CFM_AJUSTES>();
        string carpeta_final = string.Empty;
        string carpeta_ruta = string.Empty;
        string carpeta_extraidas = string.Empty;
        string carpeta_respaldo = string.Empty;
        string carpeta_asignar = string.Empty;
        string carpeta_reportar = string.Empty;

        List<string> coleccion_tipo_cargo_filtrado = new List<string>();
        List<string> coleccion_descripcion_filtrado = new List<string>();
        SLDocument oSLDocument = new SLDocument();
        private BackgroundWorker bw_ejecutar;
        private BackgroundWorker bw_analizar;
        public Ajuste_CFM()
        {
            InitializeComponent();
            this.bw_ejecutar = new BackgroundWorker();
            this.bw_ejecutar.DoWork += new DoWorkEventHandler(bw_DoWork_ejecutar);//este sirve para los metodos
            this.bw_ejecutar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_ejecutar);//este envia el progreso
            this.bw_ejecutar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_ejecutar);//este cuando termina
            this.bw_ejecutar.WorkerReportsProgress = true;

            this.ejecutar_btn.Click += new EventHandler(Ejecutar_Click);

            this.bw_analizar = new BackgroundWorker();
            this.bw_analizar.DoWork += new DoWorkEventHandler(bw_DoWork_analizar);//este sirve para los metodos
            this.bw_analizar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_analizar);//este envia el progreso
            this.bw_analizar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_analizar);//este cuando termina
            this.bw_analizar.WorkerReportsProgress = true;

            this.analizar_Btn.Click += new EventHandler(analizar_fide_cob_btn_Click);
            ejecutar_btn.Visible = false;
            guardar.Visible = false;
            ejecutar_btn.Visible = false;
        }

        private void Abrir_Cob_pc_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            quedaran_aqui_txt.Text = aux_path;
        }

        private void analizar_fide_cob_btn_Click(object sender, EventArgs e)
        {
            if (!this.bw_analizar.IsBusy)
            {
                this.bw_analizar.RunWorkerAsync();
                this.analizar_Btn.Enabled = false;
            }
        }

        private void guardar_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tipo_dgv.RowCount - 1; i++)
            {
                if (tipo_dgv.Rows[i].Cells["check"].Value.Equals(true))
                {
                    coleccion_tipo_cargo_filtrado.Add(tipo_dgv.Rows[i].Cells["TIPO_CARGO"].Value.ToString());
                }
            }
            for (int i = 0; i < descripcion_dgv.RowCount - 1; i++)
            {
                if (descripcion_dgv.Rows[i].Cells["check"].Value.Equals(true))
                {
                    coleccion_descripcion_filtrado.Add(descripcion_dgv.Rows[i].Cells["DESCRIPTION"].Value.ToString());
                }
            }

            tipo_dgv.Visible = false;
            descripcion_dgv.Visible = false;
            guardar.Visible = false;
            string message = "Se agregaron los tipos";
            MessageBox.Show(message);
            ejecutar_btn.Visible = true;
        }

        private void Ejecutar_Click(object sender, EventArgs e) //almacenar los pcs y las cuentas
        {
            if (!this.bw_ejecutar.IsBusy)
            {
                this.bw_ejecutar.RunWorkerAsync();
                this.ejecutar_btn.Enabled = false;
                this.analizar_Btn.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            CargosxPcs_1_txt.Text = aux_path;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            //cartera_txt.Text = aux_path;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            //Cartera_Empresarial_txt.Text = aux_path;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            CargosxPcs_2_txt.Text = aux_path;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            ajustes_cfm_txt.Text = aux_path;
        }

        private void bw_ProgressChanged_analizar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;//izi pizzi funcionando el progress bar

        }
        private void bw_RunWorkerCompleted_analizar(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            //this.txt_cambiante.Text = "100" + "% complete";
            
            tipo_dgv.Visible = Visible;
            
            descripcion_dgv.Visible = Visible;
            guardar.Visible = true;
            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show("Proceso terminado!");
        }

        private void bw_DoWork_analizar(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            string message = "Analizando Codigos de Cargos";
            int i = 10;
            worker.ReportProgress(i);
            Coleccion_Fide_Cob_Pcs coleccion_Fide_Cob_Pcs = new Coleccion_Fide_Cob_Pcs();
            Coleccion_CFM coleccion_CFM = new Coleccion_CFM();
            cargosx_pcs1 = coleccion_Fide_Cob_Pcs.GenerarListado(CargosxPcs_1_txt.Text);
            if (CargosxPcs_2_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosx_pcs2 = coleccion_Fide_Cob_Pcs.GenerarListado(CargosxPcs_2_txt.Text);
            }
            if (CargosxPcs_3_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosx_pcs3 = coleccion_Fide_Cob_Pcs.GenerarListado(CargosxPcs_3_txt.Text);
            }
            if (CargosxPcs_4_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosx_pcs4 = coleccion_Fide_Cob_Pcs.GenerarListado(CargosxPcs_4_txt.Text);
            }
            if (CargosxPcs_5_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosx_pcs5 = coleccion_Fide_Cob_Pcs.GenerarListado(CargosxPcs_5_txt.Text);
            }
            if (CargosxPcs_6_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosx_pcs6 = coleccion_Fide_Cob_Pcs.GenerarListado(CargosxPcs_6_txt.Text);
            }
            i = 25;
            worker.ReportProgress(i);
            if (cargosx_pcs2.Count()>1)
            {
                foreach (var item in cargosx_pcs2)
                {
                    cargosx_pcs1.Add(item);
                }
            }
            if (cargosx_pcs3.Count() > 1)
            {
                foreach (var item in cargosx_pcs3)
                {
                    cargosx_pcs1.Add(item);
                }
            }
            if (cargosx_pcs4.Count() > 1)
            {
                foreach (var item in cargosx_pcs4)
                {
                    cargosx_pcs1.Add(item);
                }
            }
            if (cargosx_pcs5.Count() > 1)
            {
                foreach (var item in cargosx_pcs5)
                {
                    cargosx_pcs1.Add(item);
                }
            }
            if (cargosx_pcs6.Count() > 1)
            {
                foreach (var item in cargosx_pcs6)
                {
                    cargosx_pcs1.Add(item);
                }
            }
            

            cargosx_pcs1.Count();
            DataTable dtEmp1 = new DataTable();
            DataTable dtEmp2 = new DataTable();
            // add column to datatable 
            List<string> coleccion_tipo_cargo = new List<string>();
            List<string> coleccion_descripcion = new List<string>();
            var _tipo = cargosx_pcs1.Select(x => x.tipocargo).Distinct().ToList();
            var _descripcion = cargosx_pcs1.Select(x => x.description).Distinct().ToList();
            foreach (var item in _tipo)
            {
                coleccion_tipo_cargo.Add(item);
            }

            foreach (var item in cargosx_pcs1)
            {
                if (item.tipocargo.Equals("Cargo Fijo"))
                {
                    coleccion_descripcion.Add(item.description);
                }
            }
            i = 50;
            worker.ReportProgress(i);
            dtEmp1.Columns.Add("TIPO_CARGO", typeof(string));
            dtEmp1.Columns.Add("check", typeof(bool));

            dtEmp2.Columns.Add("DESCRIPTION", typeof(string));
            dtEmp2.Columns.Add("check", typeof(bool));

            foreach (var item in coleccion_tipo_cargo.Distinct())
            {
                dtEmp1.Rows.Add(new Object[] { item, false });
            }

            foreach (var item in coleccion_descripcion.Distinct())
            {
                dtEmp2.Rows.Add(new Object[] { item, false });
            }
            tipo_dgv.DataSource = dtEmp1;
            descripcion_dgv.DataSource = dtEmp2;
            i = 100;
            worker.ReportProgress(i);


            i = 0;
            worker.ReportProgress(i);

        }


        private void bw_RunWorkerCompleted_ejecutar(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            //this.txt_cambiante.Text = "100" + "% complete";
            this.ejecutar_btn.Enabled = true;
            //this.segmento_dgv.Visible = true;
            //this.btn_Guardar.Visible = true;
            // MessageBox.Show("Proceso terminado!");
        }

        private void bw_ProgressChanged_ejecutar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;//izi pizzi funcionando el progress bar

        }

        private void bw_DoWork_ejecutar(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B //aqui no estamos haciendo nada que te pasa
        {
            //this.ejecutar_btn.Enabled = false;
            BackgroundWorker worker = (BackgroundWorker)sender;
            int i = 40;
            worker.ReportProgress(i);
            CreacionCarpetas();
            Controlador_CFM controlador_CFM = new Controlador_CFM(cargosx_pcs1, ajustes_cfm_txt.Text, quedaran_aqui_txt.Text, coleccion_descripcion_filtrado, coleccion_tipo_cargo_filtrado[0]);
            controlador_CFM.proceso_cfm_filtrando();
            i = 60;
            worker.ReportProgress(i);
            controlador_CFM.proceso_cfm_duplicados();
            i = 80;
            worker.ReportProgress(i);
            controlador_CFM.proceso_cfm_sumados();
            i = 100;
            worker.ReportProgress(i);
            MessageBox.Show("Listo!");

        }
        public void CreacionCarpetas()
        {
            Console.WriteLine("Creando el directorio: {0}", @quedaran_aqui_txt.Text);
            carpeta_final = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Final";
            carpeta_ruta = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta";
            carpeta_extraidas = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas";
            carpeta_respaldo = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo";
            carpeta_asignar = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Asignar";
            carpeta_reportar = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Reportar";

            if (Directory.Exists(@quedaran_aqui_txt.Text + @"\Ciclo"))
            {

            }
            else
            {
                DirectoryInfo di1 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo");
            }

            if (Directory.Exists(@carpeta_final))
            {

            }
            else
            {
                DirectoryInfo di3 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Final");
            }

            if (Directory.Exists(carpeta_ruta))
            {

            }
            else
            {
                DirectoryInfo di4 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta");
            }
            if (Directory.Exists(carpeta_extraidas))
            {

            }
            else
            {
                DirectoryInfo di5 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas");
            }
            if (Directory.Exists(carpeta_respaldo))
            {

            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                ajustes_cfm_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                ajustes_cfm_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                ajustes_cfm_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                ajustes_cfm_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                ajustes_cfm_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void quedaran_aqui_txt_TextChanged(object sender, EventArgs e)
        {

        }

        private void CargosxPcs_1_txt_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
