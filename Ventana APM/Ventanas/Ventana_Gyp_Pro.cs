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
using Ventana_APM.CONTROLADOR;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class Ventana_Gyp_Pro : Form
    {
        string carpeta_final = string.Empty;
        string carpeta_ruta = string.Empty;
        string carpeta_extraidas = string.Empty;
        string carpeta_respaldo = string.Empty;
        string carpeta_asignar = string.Empty;
        string carpeta_reportar = string.Empty;
        string quedaran_aqui_txt = string.Empty;
        string aux_path = string.Empty;
        List<string> tipos_ = new List<string>();
        Coleccion_CargosxPcs coleccion_Cargosx = new Coleccion_CargosxPcs();
        Coleccion_Cuentas_Asignadas _Asignadas = new Coleccion_Cuentas_Asignadas();
        List<CargosxPcs> cargosxPCs_Lista = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPCs_Lista2 = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPCs_Lista3 = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPCs_Lista4 = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPCs_Lista5 = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPCs_Lista6 = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPCs_Lista7 = new List<CargosxPcs>();
        List<Cuentas_Asignadas> cuentas = new List<Cuentas_Asignadas>();
        List<CargosxPcs> cargosxPcs = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPcs2 = new List<CargosxPcs>();
        private BackgroundWorker bw_ejecutar;
        private BackgroundWorker bw_analizar;
        SLDocument oSLDocument = new SLDocument();
        public Ventana_Gyp_Pro()
        {
            InitializeComponent();
            this.bw_ejecutar = new BackgroundWorker();
            this.bw_ejecutar.DoWork += new DoWorkEventHandler(bw_DoWork_ejecutar);//este sirve para los metodos
            this.bw_ejecutar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_ejecutar);//este envia el progreso
            this.bw_ejecutar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_ejecutar);//este cuando termina
            this.bw_ejecutar.WorkerReportsProgress = true;

            this.ejecutar_btn.Click += new EventHandler(ejecutar_btn_Click);

            this.bw_analizar = new BackgroundWorker();
            this.bw_analizar.DoWork += new DoWorkEventHandler(bw_DoWork_analizar);//este sirve para los metodos
            this.bw_analizar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_analizar);//este envia el progreso
            this.bw_analizar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_analizar);//este cuando termina
            this.bw_analizar.WorkerReportsProgress = true;

            this.analizar_Btn.Click += new EventHandler(analizar_Btn_Click);
            ejecutar_btn.Visible = false;
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
            StringBuilder sb = new StringBuilder();
            if (e.ProgressPercentage == 10)
            {
                sb.AppendLine("Creando Carpeta...");
                queestasucediendo_txt.Text = sb.ToString();
            }
            else if (e.ProgressPercentage == 20)
            {
                sb.AppendLine("Filtrando por codigos...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 40)
            {
                sb.AppendLine("Sumando montos de un pcs...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 90)
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

        private void bw_DoWork_ejecutar(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B //aqui no estamos haciendo nada que te pasa
        {
            //this.ejecutar_btn.Enabled = false;
            BackgroundWorker worker = (BackgroundWorker)sender;
            CreacionCarpetas();
            int i = 10;
            worker.ReportProgress(i);

            Controlador_Gyp controlador_Gyp = new Controlador_Gyp(gyp_path_txt.Text, cargosxPcs, tipos_, seleccionar_carpeta_txt.Text);
            controlador_Gyp.procesando_cargosx_pc_tipo();
            controlador_Gyp.procesando_cargosx_pc_x_gyp();
            controlador_Gyp.procesando_cargosx_pc_sin_0();
            controlador_Gyp.procesando_cargosx_pc_duplicados();

            // foreach (var item1 in cargosxPcs)
            // {
            //     foreach (var item2 in tipos_)
            //     {
            //         if (item1.codigodecargo.Equals(item2))
            //         {
            //             cargosxPcs2.Add(item1); //hasta aqui esta bien
            //         }
            //     }

            // }

            // var query = cargosxPcs2.GroupBy(d => d.PCS)
            //                 .Select(
            //                         g => new
            //                         {
            //                             Key = g.Key,
            //                             Monto_Total = g.Sum(s => s.montoTotal),
            //                             CuentaFinanciera = g.First().CuentaFinanciera,
            //                             Pcs = g.First().PCS,
            //                             CodigoCarga = g.First().codigodecargo
            //                         });

            // List<CargosxPcs> cargosxPcs_duplicados = new List<CargosxPcs>();
            // worker.ReportProgress(i);
            // i = 40;
            // worker.ReportProgress(i);
            // i = 50;

            // foreach (var item1 in query)
            // {
            //     int contador = 0;
            //     foreach (var item2 in cargosxPcs2)
            //     {
            //         if (item1.CuentaFinanciera.Equals(item2.CuentaFinanciera))
            //         {
            //             contador++;
            //             if (contador <= 2)
            //             {
            //                 if (item2.montoTotal == 0)
            //                 {

            //                 }
            //                 else
            //                 {
            //                     cargosxPcs_duplicados.Add(item2);
            //                 }

            //             }
            //         }
            //     }
            // }
            // worker.ReportProgress(i);
            // i = 70;
            // worker.ReportProgress(i);
            // // DirectoryInfo di5 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\Gyp Pro");
            //// SaveExcel.BuildExcel(miTable2, @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
            // var dt = new DataTable("CCUCTEL_ADIC");
            // dt.Columns.Add("CUENTA", typeof(long));
            // dt.Columns.Add("PCS", typeof(long));
            // dt.Columns.Add("CARGO", typeof(string));
            // dt.Columns.Add("AMOUNT", typeof(double));
            // dt.Columns.Add("CICLO", typeof(string));
            // foreach (var item in cargosxPcs_duplicados)
            // {
            //     item.montoTotal = item.montoTotal * -1;
            //         dt.Rows.Add(new Object[] {long.Parse(item.CuentaFinanciera),
            //                                    long.Parse(item.PCS),
            //                                    item.codigodecargo,
            //                                    item.montoTotal, ""});

            // }
            // i = 90;
            // worker.ReportProgress(i);
            // oSLDocument.ImportDataTable(1, 1, dt, true);
            // oSLDocument.RenameWorksheet("Sheet1", "CCUCTEL_ADIC");
            // oSLDocument.SaveAs(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
            //Habilitaciones();
            i = 100;
            worker.ReportProgress(i);
            MessageBox.Show("Listo!");

        }

        public void CreacionCarpetas()
        {
            Console.WriteLine("Creando el directorio: {0}", @seleccionar_carpeta_txt.Text);
            carpeta_final = @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Final";
            carpeta_ruta = @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta";
            carpeta_extraidas = @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Extraidas";
            carpeta_respaldo = @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Respaldo";
            carpeta_asignar = @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Asignar";
            carpeta_reportar = @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Reportar";

            if (Directory.Exists(@seleccionar_carpeta_txt.Text + @"\Ciclo"))
            {

            }
            else
            {
                DirectoryInfo di1 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo");
            }

            if (Directory.Exists(@carpeta_final))
            {

            }
            else
            {
                DirectoryInfo di3 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Final");
            }

            if (Directory.Exists(carpeta_ruta))
            {

            }
            else
            {
                DirectoryInfo di4 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta");
            }
            if (Directory.Exists(carpeta_extraidas))
            {

            }
            else
            {
                DirectoryInfo di5 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Extraidas");
            }
            if (Directory.Exists(carpeta_respaldo))
            {

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (var item1 in cargosxPcs)
            {
                foreach (var item2 in tipos_)
                {
                    if (item1.codigodecargo.Equals(item2))
                    {
                        cargosxPcs2.Add(item1); //hasta aqui esta bien
                    }
                }
                    
             }

            var query = cargosxPcs2.GroupBy(d => d.PCS)
                            .Select(
                                    g => new
                                    {
                                        Key = g.Key,
                                        Monto_Total = g.Sum(s => s.montoTotal),
                                        CuentaFinanciera = g.First().CuentaFinanciera,
                                        Pcs = g.First().PCS,
                                        CodigoCarga = g.First().codigodecargo
                                    });

            List<CargosxPcs> cargosxPcs_duplicados = new List<CargosxPcs>();

            foreach (var item1 in query)
            {
                int contador = 0;
                foreach (var item2 in cargosxPcs2)
                {
                    if (item1.CuentaFinanciera.Equals(item2.CuentaFinanciera))
                    {
                        contador++;
                        if (contador <= 2)
                        {
                            if (item2.montoTotal == 0)
                            {

                            }
                            else
                            {
                                cargosxPcs_duplicados.Add(item2);
                            }
                            
                        }
                    }
                }
            }

            var miTable2 = new DataTable("CCUCTEL_ADIC");
                 

            miTable2.Columns.Add("CUENTA");
            miTable2.Columns.Add("PCS");
            miTable2.Columns.Add("CARGO");
            miTable2.Columns.Add("AMOUNT");
            miTable2.Columns.Add("CICLO");

            foreach (var item in cargosxPcs_duplicados)
            {
                miTable2.Rows.Add(new Object[] {item.CuentaFinanciera, item.PCS,item.codigodecargo,"-"+item.montoTotal,""});
            }
           // DirectoryInfo di5 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\Gyp Pro");
            SaveExcel.BuildExcel(miTable2, @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
            MessageBox.Show("Termine!!");
            Console.WriteLine();
        }

            // este es el boton de ejecutar
        

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            seleccionar_carpeta_txt.Text = aux_path;
        }



        private void button6_Click(object sender, EventArgs e)
        {
            cargosxPCs_Lista = coleccion_Cargosx.GenerarListado(@cargosxpcs_txt.Text);
            cuentas = _Asignadas.GenerarListado(@gyp_path_txt.Text);
           
            List<string> coleccion_cargos_codigo = new List<string>();

            foreach (var item1 in cargosxPCs_Lista)
            {
                foreach (var item2 in cuentas)
                {
                    if (item1.CuentaFinanciera.Equals(item2.PCS))//no es pcs es cuenta lo debo cambiar
                    {
                        cargosxPcs.Add(item1);
                        coleccion_cargos_codigo.Add(item1.codigodecargo);

                    }
                }
            }
            DataTable dtEmp = new DataTable();
            // add column to datatable  

            dtEmp.Columns.Add("Tipo", typeof(string));
            dtEmp.Columns.Add("check", typeof(bool));

            foreach (var item in coleccion_cargos_codigo.Distinct())
            {
                dtEmp.Rows.Add(new Object[] { item, false });
            }
            segmento_dgv.DataSource = dtEmp;
            segmento_dgv.Visible = true;
            btn_Guardar.Visible = true;//este es el boton de analizar
        }

        private void btn_Guardar_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < segmento_dgv.RowCount - 1; i++)
            {
                if (segmento_dgv.Rows[i].Cells["check"].Value.Equals(true))
                {
                    tipos_.Add(segmento_dgv.Rows[i].Cells["Tipo"].Value.ToString());
                }
            }

            segmento_dgv.Visible = false;
            btn_Guardar.Visible = false;
            this.analizar_Btn.Visible = false;
            string message = "Se agregaron los tipos";
            MessageBox.Show(message);
        }

        private void analizar_Btn_Click(object sender, EventArgs e)
        {
            if (!this.bw_analizar.IsBusy)
            {
                this.bw_analizar.RunWorkerAsync();
                this.analizar_Btn.Enabled = false;
            }

        }

        private void btn_Guardar_Click_1(object sender, EventArgs e)
        {
            for (int i = 0; i < segmento_dgv.RowCount - 1; i++)
            {
                if (segmento_dgv.Rows[i].Cells["check"].Value.Equals(true))
                {
                    tipos_.Add(segmento_dgv.Rows[i].Cells["Tipo"].Value.ToString());
                }
            }

            segmento_dgv.Visible = false;
            btn_Guardar.Visible = false;
            ejecutar_btn.Visible = true;
            string message = "Se agregaron los tipos";
            MessageBox.Show(message);
        }

        private void ejecutar_btn_Click(object sender, EventArgs e)
        {
            if (!this.bw_ejecutar.IsBusy)
            {
                this.bw_ejecutar.RunWorkerAsync();
                this.ejecutar_btn.Enabled = false;
                this.analizar_Btn.Enabled = false;
            }
        }
        private void bw_ProgressChanged_analizar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;//izi pizzi funcionando el progress bar
            StringBuilder sb = new StringBuilder();
            if (e.ProgressPercentage == 10)
            {
                sb.AppendLine("Cargando codigos de cargos...");
                queestasucediendo_txt.Text = sb.ToString();
            }
            else if (e.ProgressPercentage == 100)
            {
                sb.AppendLine("Proceso Terminado");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }

        }
        private void bw_RunWorkerCompleted_analizar(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            //this.txt_cambiante.Text = "100" + "% complete";
            this.analizar_Btn.Enabled = false;
            this.analizar_Btn.Visible = false;
            this.segmento_dgv.Visible = true;
            this.btn_Guardar.Visible = true;
            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show("Proceso terminado!");
        }

        private void bw_DoWork_analizar(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            string message = "Analizando Codigos de Cargos";
            // MessageBox.Show(message);
            int i = 10;
            worker.ReportProgress(i);
          
            cargosxPCs_Lista = coleccion_Cargosx.GenerarListado(@cargosxpcs_txt.Text);
            if (cargosxpcs2_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosxPCs_Lista2 = coleccion_Cargosx.GenerarListado(@cargosxpcs2_txt.Text);
                foreach (var item in cargosxPCs_Lista2)
                {
                    cargosxPCs_Lista.Add(item);
                }
            }
            if (cargosxpcs3_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosxPCs_Lista3 = coleccion_Cargosx.GenerarListado(@cargosxpcs3_txt.Text);
                foreach (var item in cargosxPCs_Lista3)
                {
                    cargosxPCs_Lista.Add(item);
                }
            }
            if (cargosxpcs4_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosxPCs_Lista4 = coleccion_Cargosx.GenerarListado(@cargosxpcs4_txt.Text);
                foreach (var item in cargosxPCs_Lista4)
                {
                    cargosxPCs_Lista.Add(item);
                }
            }
            if (cargosxpcs5_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosxPCs_Lista5 = coleccion_Cargosx.GenerarListado(@cargosxpcs5_txt.Text);
                foreach (var item in cargosxPCs_Lista5)
                {
                    cargosxPCs_Lista.Add(item);
                }
            }
            if (cargosxpcs6_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosxPCs_Lista6 = coleccion_Cargosx.GenerarListado(@cargosxpcs6_txt.Text);
                foreach (var item in cargosxPCs_Lista6)
                {
                    cargosxPCs_Lista.Add(item);
                }
            }
            if (cargosxpcs7_txt.Text.Equals(""))
            {

            }
            else
            {
                cargosxPCs_Lista7 = coleccion_Cargosx.GenerarListado(@cargosxpcs7_txt.Text);
                foreach (var item in cargosxPCs_Lista7)
                {
                    cargosxPCs_Lista.Add(item);
                }
            }


            cuentas = _Asignadas.GenerarListado(gyp_path_txt.Text);

            List<string> coleccion_cargos_codigo = new List<string>();
            //"Cargos por Tráfico Adicional"
            i = 50;
            worker.ReportProgress(i);
            foreach (var item1 in cargosxPCs_Lista)
            {
                foreach (var item2 in cuentas)
                {
                    if (item1.CuentaFinanciera.Equals(item2.PCS))//no es pcs es cuenta lo debo cambiar
                    {
                        cargosxPcs.Add(item1);
                        coleccion_cargos_codigo.Add(item1.codigodecargo);

                    }
                }
            }

            DataTable dtEmp = new DataTable();
            // add column to datatable  
            i = 70;
            worker.ReportProgress(i);
            dtEmp.Columns.Add("Tipo", typeof(string));
            dtEmp.Columns.Add("check", typeof(bool));

            foreach (var item in coleccion_cargos_codigo.Distinct())
            {
                dtEmp.Rows.Add(new Object[] { item, false });
            }
            segmento_dgv.DataSource = dtEmp;
            i = 100;
            worker.ReportProgress(i);


            i = 0;
            worker.ReportProgress(i);
            if (coleccion_cargos_codigo.Count() == 0)
            {
                Crear_Gyp_Vacio();
            }

        }

        private void Crear_Gyp_Vacio()
        {
            try
            {

                var dt = new DataTable("Gyp");
                dt.Columns.Add("CUENTA", typeof(string));
                dt.Columns.Add("PCS", typeof(string));
                dt.Columns.Add("CARGO", typeof(string));
                dt.Columns.Add("AMOUNT", typeof(string));
                dt.Columns.Add("CICLO", typeof(string));
                dt.Rows.Add(new Object[] {"",
                                               "",
                                               "",
                                               "", ""});

                oSLDocument.ImportDataTable(1, 1, dt, true);
                oSLDocument.RenameWorksheet("Sheet1", "GYP");
                oSLDocument.SaveAs(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
                string message = "El archivo GYP PRO fue creado";
                MessageBox.Show(message);
            }
            catch (Exception ex)
            {

                CreacionCarpetas();
                var dt = new DataTable("Gyp");
                dt.Columns.Add("CUENTA", typeof(string));
                dt.Columns.Add("PCS", typeof(string));
                dt.Columns.Add("CARGO", typeof(string));
                dt.Columns.Add("AMOUNT", typeof(string));
                dt.Columns.Add("CICLO", typeof(string));
                dt.Rows.Add(new Object[] {"",
                                               "",
                                               "",
                                               "", ""});

                oSLDocument.ImportDataTable(1, 1, dt, true);
                oSLDocument.RenameWorksheet("Sheet1", "GYP");
                oSLDocument.SaveAs(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
                string message = "El archivo GYP PRO fue creado";
                MessageBox.Show(message);

            }

        }

        private void btn_gyp_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            //ofd.Filter;
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            gyp_path_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs2_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs3_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs4_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs5_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs6_txt.Text = aux_path;
        }

        private void btn_cargosxpcs_7_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargosxpcs7_txt.Text = aux_path;
        }
    }
}
