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
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.Auxiliares;
using Ventana_APM.CONTROLADOR;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class Ajustes_Varios : Form
    {
        string aux_path = string.Empty;
        string carpeta_final = string.Empty;
        string carpeta_ruta = string.Empty;
        string carpeta_extraidas = string.Empty;
        string carpeta_respaldo = string.Empty;
        string carpeta_asignar = string.Empty;
        string carpeta_reportar = string.Empty;
        private BackgroundWorker bw_ejecutar;
        List<Cartera_Empresarial> cartera_Empresarials_sice;
        List<string> tipos_ = new List<string>();
        List<Cartera> carteras_pablo;
        
        public Ajustes_Varios(List<Cartera_Empresarial> cartera_Empresarials, List<Cartera> carteras, List<string> filtro)
        {
            InitializeComponent();
            cartera_Empresarials_sice = cartera_Empresarials;
            carteras_pablo = carteras;
            tipos_ = filtro;
            this.bw_ejecutar = new BackgroundWorker();
            this.bw_ejecutar.DoWork += new DoWorkEventHandler(bw_DoWork_ejecutar);//este sirve para los metodos
            this.bw_ejecutar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_ejecutar);//este envia el progreso
            this.bw_ejecutar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_ejecutar);//este cuando termina
            this.bw_ejecutar.WorkerReportsProgress = true;

            this.ejecutar_Btn.Click += new EventHandler(ejecutar_Btn_Click);

            panel4.Visible = false;
            panel6.Visible = false;
            panel10.Visible = false;
            panel16.Visible = false;
            panel15.Visible = false;
            panel14.Visible = false;

        }


        private void bw_RunWorkerCompleted_ejecutar(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            //this.txt_cambiante.Text = "100" + "% complete";
            this.ejecutar_Btn.Enabled = true;
            //this.segmento_dgv.Visible = true;
            //this.btn_Guardar.Visible = true;
            MessageBox.Show("Proceso terminado!");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_ProgressChanged_ejecutar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;
            StringBuilder sb = new StringBuilder();
            //if (e.ProgressPercentage == 10)
            //{
            //    sb.AppendLine("Creando Carpeta...");
            //    queestasucediendo_txt.Text = sb.ToString();
            //}
            //else if (e.ProgressPercentage == 20)
            //{
            //    sb.AppendLine("Cargando DownGrade...");
            //    queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            //}
            //else if (e.ProgressPercentage == 30)
            //{
            //    sb.AppendLine("Procesando DownGrade...");
            //    queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            //}
            if (e.ProgressPercentage == 10)
            {
                sb.AppendLine("Cargando Dummy...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 40)
            {
                sb.AppendLine("Procesando Dummy...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            //else if (e.ProgressPercentage == 55)
            //{
            //    sb.AppendLine("Procesando SimCard...");
            //    queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            //}
            else if (e.ProgressPercentage == 80)
            {
                sb.AppendLine("Creando archivos...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 100)
            {
                sb.AppendLine("Proceso Terminado...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }

        }

        private void bw_DoWork_ejecutar(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            Controlador_Hbo_Paramount controlador_Hbo_Paramount = new Controlador_Hbo_Paramount(@hbo_txt.Text, @paramount_txt.Text, @quedaran_aqui_txt.Text);
            int i = 10;
            worker.ReportProgress(i);
            i = 20;
            worker.ReportProgress(i);
            CreacionCarpetas();
            worker.ReportProgress(i);
            Controlador_Dummy controlador_Dummy = new Controlador_Dummy(@dummy_txt.Text, @quedaran_aqui_txt.Text, cartera_Empresarials_sice, carteras_pablo, tipos_);

            controlador_Dummy.filtrando_dummys();

            //Controlador_DownGrade controlador_DownGrade = new Controlador_DownGrade(@ajustes_downgrade_txt.Text,@ajustes_cfm_txt.Text, solidarios_Txt.Text, @quedaran_aqui_txt.Text);
            i = 40;
            worker.ReportProgress(i);
            controlador_Dummy.proceso_dummy();
            //controlador_DownGrade.proceso_downgrade();
            i = 50;
            worker.ReportProgress(i);

            //controlador_Dummy.aplicar_filtro(true);
            //i = 65;
            //controlador_Hbo_Paramount.hbo_31();
            //controlador_Hbo_Paramount.paramount_sum();

            //i = 75;
            //worker.ReportProgress(i);
            //Controlador_Simcard controlador_Simcard = new Controlador_Simcard(@simcard_txt.Text,@quedaran_aqui_txt.Text);
            i = 80;
            worker.ReportProgress(i);
            //controlador_Simcard.procesando_simcard();
            i = 100;
            worker.ReportProgress(i);
            //creo que ta listo

            //MessageBox.Show("Proceso Terminado");

        }

        private void seleccionar_carpeta_Btn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                aux_path = fbd.SelectedPath.ToString();
                quedaran_aqui_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void simcard_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                simcard_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void ajustes_especiales_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                aux_path = ofd.FileName;
                ajustes_downgrade_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }

        }

        private void ajustes_downgrade_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                ajustes_downgrade_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void ajustes_cfm_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
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

        private void seleccionar_carpeta_Btn_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                aux_path = fbd.SelectedPath.ToString();
                quedaran_aqui_txt.Text = aux_path;
            }

            else
            {
                aux_path = string.Empty;
            }
        }

        private void planilla_solidarios_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                dummy_txt.Text = aux_path;
            }
            else            
            {
                aux_path = string.Empty;
            }
                    
            


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
            


        }

        private void ejecutar_Btn_Click(object sender, EventArgs e)
        {
            if (!this.bw_ejecutar.IsBusy)
            {
                this.bw_ejecutar.RunWorkerAsync();
                this.ejecutar_Btn.Enabled = false;
                this.ejecutar_Btn.Enabled = false;
            }
            //PRIMERO SIMCARD


            //List<PlanSolidario> planSolidarios = new List<PlanSolidario>();
            //Coleccion_PlanSolidario coleccion_Plan = new Coleccion_PlanSolidario();
            //planSolidarios = coleccion_Plan.GenerarListado(planilla_solidarios_txt.Text);

            //List<DownGrade> downGrades = new List<DownGrade>();
            //Coleccion_Downgrade coleccion_Downgrade = new Coleccion_Downgrade();
            //downGrades = coleccion_Downgrade.GenerarListado(ajustes_downgrade_txt.Text);

            //List<CFM_AJUSTES> cFM_s = new List<CFM_AJUSTES>();
            //Coleccion_CFM coleccion_CFM = new Coleccion_CFM();
            //cFM_s = coleccion_CFM.GenerarListado(ajustes_cfm_txt.Text);

            //List<DownGrade> downGrades1 = new List<DownGrade>();
            //foreach (var item1 in downGrades)
            //{
            //    foreach (var item2 in planSolidarios)
            //    {
            //        if (item1.PCS.Equals(item2.PCS))
            //        {
            //            downGrades1.Add(item1);
            //        }
            //    }
            //}

            //foreach (var item1 in downGrades)
            //{
            //    foreach (var item2 in cFM_s)
            //    {
            //        if (item1.PCS.Equals(item2.PCS))
            //        {
            //            downGrades1.Add(item1);
            //        }
            //    }
            //}

            //var miTable26 = new DataTable("AJUSTES DOWNGRADE"); //D: cual es la base del monto
            //miTable26.Columns.Add("PCS");
            //miTable26.Columns.Add("CUENTA FINANCIERA");
            //miTable26.Columns.Add("CODIGO DE CARGO");
            //miTable26.Columns.Add("AMOUNT");
            //miTable26.Columns.Add("CICLO");


            //foreach (var item in downGrades1)
            //{
            //    if (item. > 0.0)
            //    {
            //        miTable26.Rows.Add(new Object[] { item.Pcs, item.CuentaFinanciera, item.CodigoCarga, "-" + item.Monto_Total, "" });
            //    }
            //}

            //SaveExcel.BuildExcel(miTable26, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Final\Ajustes especiales" + @"\AJUSTE_ESPECIALES CICLO_MES_ANIO.xlsx");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {            
                aux_path = ofd.FileName;
                solidarios_Txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void Paramount_Btn_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                paramount_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }

        }

        private void Hbo_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                aux_path = ofd.FileName;
                hbo_txt.Text = aux_path;
            }
            else
            {
                aux_path = string.Empty;
            }
        }

        private void paramount_txt_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void simcard_txt_TextChanged(object sender, EventArgs e)
        {

        }

        private void ajustes_downgrade_txt_TextChanged(object sender, EventArgs e)
        {

        }

        private void hbo_txt_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel14_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void quedaran_aqui_txt_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
