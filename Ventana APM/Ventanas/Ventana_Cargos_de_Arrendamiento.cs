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
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;
using Ventana_APM.CONTROLADOR;

namespace Ventana_APM.Ventanas
{
    public partial class Ventana_Cargos_de_Arrendamiento : Form
    {
        string aux_path = string.Empty;
        Coleccion_Cartera_Empresarial _Cartera_Empresarial = new Coleccion_Cartera_Empresarial();
        Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
        List<Cartera_Empresarial> cartera_Empresarials_sice = new List<Cartera_Empresarial>();
        List<Cartera> carteras_pablo = new List<Cartera>();
        List<string> tipos_ = new List<string>();
        string carpeta_final = string.Empty;
        string carpeta_ruta = string.Empty;
        string carpeta_extraidas = string.Empty;
        string carpeta_respaldo = string.Empty;
        string carpeta_asignar = string.Empty;
        string carpeta_reportar = string.Empty;
        private BackgroundWorker bw_ejecutar;

        public Ventana_Cargos_de_Arrendamiento(List<Cartera_Empresarial> cartera_Empresarials, List<Cartera> carteras, List<string> filtro)
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
            this.ejecutar_Btn.Visible = true;
        }

        private void bw_RunWorkerCompleted_ejecutar(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            //this.txt_cambiante.Text = "100" + "% complete";
            this.ejecutar_Btn.Enabled = true;
            //this.segmento_dgv.Visible = true;
            //this.btn_Guardar.Visible = true;
            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show("Proceso terminado!");
        }

        private void bw_ProgressChanged_ejecutar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;
            StringBuilder sb = new StringBuilder();
            if (e.ProgressPercentage == 10)
            {
                sb.AppendLine("Creando Carpeta...");
                queestasucediendo_txt.Text = sb.ToString();
            }
            else if (e.ProgressPercentage == 20)
            {
                sb.AppendLine("Comparando Cargos de Arrendamiento con Carteras Empresariales...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 30)
            {
                sb.AppendLine("Cargos de Arrendamiento Neteados...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 40)
            {
                sb.AppendLine("Ajustes Cargos de Arrendamiento...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 60)
            {
                sb.AppendLine("Buscando Duplicados...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 70)
            {
                sb.AppendLine("Guardando datos de Hoja Original...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 90)
            {
                sb.AppendLine("Creando Archivos de Arrendamiento...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 100)
            {
                sb.AppendLine("Proceso Terminado...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }

        }

        //private void bw_DoWork_ejecutar(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B
        //{
        //    BackgroundWorker worker = (BackgroundWorker)sender;
        //    int i = 10;
        //    worker.ReportProgress(i);

        //    CreacionCarpetas();
        //    //cargos = coleccion_C_Ar.GenerarListado(cargos_arrendamiento_txt.Text);
        //    //ajustes = coleccion_Ajustes.GenerarListado(ajustes_especiales_Txt.Text);
        //    //archivo original
        //    Console.WriteLine("Cargos Actuales: {0}", cargos.Count());



        //    Console.WriteLine("------------------------------");
        //    //DirectoryInfo di3 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Final\Cargos de Arrendamiento");
        //    //DirectoryInfo di4 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta\Cargos de Arrendamiento");
        //    //DirectoryInfo di5 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas\Cargos de Arrendamiento");
        //    //DirectoryInfo di6 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo\Cargos de Arrendamiento");
        //    //DirectoryInfo di7 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Reportar\Cargos de Arrendamiento");
        //    //DirectoryInfo di8 = Directory.CreateDirectory(@quedaran_aqui_txt.Text + @"\Ciclo" + @"\Asignar\Cargos de Arrendamiento");

        //    //Empresa_Arrendamiendo(); //aqui estoy 6:53
        //    i = 20;
        //    worker.ReportProgress(i);
        //    //Neteados_Arrendamiento();
        //    i = 40;
        //    worker.ReportProgress(i);
        //    //Ajustes_Arrendamiento(); //HASTA ACA JOYA
        //    i = 60;
        //    worker.ReportProgress(i);

        //    ///// ------------------------------------------asignados
        //    //List<Cargos_Arrendamiento> cargos_duplicados = new List<Cargos_Arrendamiento>();
        //    //int repetidos = 0;
        //    //foreach (var item1 in cargos)
        //    //{
        //    //    repetidos = 0;
        //    //    foreach (var item2 in cargos)
        //    //    {
        //    //        if (item1.PCS.Equals(item2.PCS))
        //    //        {
        //    //            repetidos++;
        //    //        }
        //    //    }
        //    //    if (repetidos >= 2)
        //    //    {
        //    //        cargos_duplicados.Add(item1);
        //    //    }
        //    //}



        //    //foreach (var item in cargos_duplicados)
        //    //{
        //    //    cargos.Remove(item);
        //    //}

        //    //Console.WriteLine("Cargos Actuales: {0}", cargos.Count()); //esta de mas
        //    //Console.WriteLine("------------------------------");

        //    //var miTable4 = new DataTable("DUPLICADOS");
        //    //miTable4.Columns.Add("CuentaFinanciera");
        //    //miTable4.Columns.Add("recveivercustomer");
        //    //miTable4.Columns.Add("tipocargo");
        //    //miTable4.Columns.Add("codigodecargo");
        //    //miTable4.Columns.Add("offwer");
        //    //miTable4.Columns.Add("nombreOFFER");
        //    //miTable4.Columns.Add("promo");
        //    //miTable4.Columns.Add("description");
        //    //miTable4.Columns.Add("montoCargos");
        //    //miTable4.Columns.Add("montoDescuentos");
        //    //miTable4.Columns.Add("montoTotal");
        //    //miTable4.Columns.Add("tipodeCobro");
        //    //miTable4.Columns.Add("PCS");
        //    //miTable4.Columns.Add("secuenciadeCiclo");
        //    //miTable4.Columns.Add("SEGMENTO");
        //    //miTable4.Columns.Add("prorrateo");
        //    //miTable4.Columns.Add("RUT");

        //    ////  SaveExcel.BuildExcel(miTable4, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo\Cargos de Arrendamiento" + @"\CARGOS DE ARRENDAMIENTO DUPLICADOS.xlsx");

        //    ////MAYORES A 14K

        //    //List<Cargos_Arrendamiento> cargos_mayores14 = new List<Cargos_Arrendamiento>();
        //    //List<Cargos_Arrendamiento> cargos_extraidos = new List<Cargos_Arrendamiento>();
        //    //foreach (var item_1 in cargos)
        //    //{
        //    //    if (item_1.montoTotal >= 14000.0)
        //    //    {
        //    //        cargos_mayores14.Add(item_1);
        //    //    }
        //    //}


        //    //foreach (var item in cargos_mayores14)
        //    //{
        //    //    cargos.Remove(item);
        //    //    cargos_extraidos.Add(item);
        //    //}

        //    //foreach (var item in cargos_duplicados)
        //    //{
        //    //    cargos_extraidos.Add(item);
        //    //}

        //    //Console.WriteLine("Cargos Actuales: {0}", cargos.Count());
        //    //Console.WriteLine("------------------------------");

        //    //i = 70;
        //    //worker.ReportProgress(i);


        //    //foreach (var item in cargos_extraidos)
        //    //{
        //    //    miTable4.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //    //        item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
        //    //        , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut});
        //    //}

        //    //oSLDocument.AddWorksheet("CARGOS DE ARRENDAMIENTO");//necesitamos a los duplciados aqui tambien
        //    //oSLDocument.ImportDataTable(1, 1, miTable4, true); //NO ESTOY SEGURO SI FUNCIONA
        //    ////SaveExcel.BuildExcel(miTable500, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo\Cargos de Arrendamiento" + @"\CARGOS MAYORESA14K.xlsx");//ESTO VA A CONTINUACION DEL DE ARRIBA

        //    //var miTable5001 = new DataTable("CARGOS DE ARRENDAMIENTO");
        //    //miTable5001.Columns.Add("CuentaFinanciera");
        //    //miTable5001.Columns.Add("recveivercustomer");
        //    //miTable5001.Columns.Add("tipocargo");
        //    //miTable5001.Columns.Add("codigodecargo");
        //    //miTable5001.Columns.Add("offwer");
        //    //miTable5001.Columns.Add("nombreOFFER");
        //    //miTable5001.Columns.Add("promo");
        //    //miTable5001.Columns.Add("description");
        //    //miTable5001.Columns.Add("montoCargos");
        //    //miTable5001.Columns.Add("montoDescuentos");
        //    //miTable5001.Columns.Add("montoTotal");
        //    //miTable5001.Columns.Add("tipodeCobro");
        //    //miTable5001.Columns.Add("PCS");
        //    //miTable5001.Columns.Add("secuenciadeCiclo");
        //    //miTable5001.Columns.Add("SEGMENTO");
        //    //miTable5001.Columns.Add("prorrateo");
        //    //miTable5001.Columns.Add("RUT");


        //    //foreach (var item in cargos_extraidos)//aqui es cargos
        //    //{
        //    //    miTable5001.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //    //        item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
        //    //        , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut});
        //    //}

        //    //SaveExcel.BuildExcel(miTable5001, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas" + @"\CARGOS DE ARRENDAMIENTO CICLO_XX_MES_ANIO.xlsx");

        //    i = 80;
        //    worker.ReportProgress(i);

        //    //List<Cargos_Arrendamiento> cargos_arrendamiento_con_0 = new List<Cargos_Arrendamiento>();

        //    //foreach (var item in cargos_duplicados)
        //    //{
        //    //    if (item.montoTotal == 0 || item.montoTotal == 0.0)
        //    //    {
        //    //        cargos_arrendamiento_con_0.Add(item);
        //    //    }
        //    //}

        //    //foreach (var item in cargos_arrendamiento_con_0)
        //    //{
        //    //    cargos_duplicados.Remove(item);
        //    //}

        //    //var miTable6 = new DataTable("CARGOS DE ARRENDAMIENTO");
        //    //miTable6.Columns.Add("ASIGNADO");
        //    //miTable6.Columns.Add("CuentaFinanciera");
        //    //miTable6.Columns.Add("recveivercustomer");
        //    //miTable6.Columns.Add("tipocargo");
        //    //miTable6.Columns.Add("codigodecargo");
        //    //miTable6.Columns.Add("offwer");
        //    //miTable6.Columns.Add("nombreOFFER");
        //    //miTable6.Columns.Add("promo");
        //    //miTable6.Columns.Add("description");
        //    //miTable6.Columns.Add("montoCargos");
        //    //miTable6.Columns.Add("montoDescuentos");
        //    //miTable6.Columns.Add("montoTotal");
        //    //miTable6.Columns.Add("tipodeCobro");
        //    //miTable6.Columns.Add("PCS");
        //    //miTable6.Columns.Add("secuenciadeCiclo");
        //    //miTable6.Columns.Add("SEGMENTO");
        //    //miTable6.Columns.Add("prorrateo");
        //    //miTable6.Columns.Add("RUT");
        //    //miTable6.Columns.Add("COBRAR");
        //    //miTable6.Columns.Add("MONTO");
        //    //miTable6.Columns.Add("CUOTA");
        //    //miTable6.Columns.Add("OBSERVACION");

        //    //foreach (var item in cargos_duplicados)
        //    //{
        //    //    miTable6.Rows.Add(new Object[] { "",item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //    //        item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
        //    //        , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut,"","","",""});
        //    //}

        //    //SaveExcel.BuildExcel(miTable6, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo\Cargos de Arrendamiento" + @"\CARGOS DE ARRENDAMIENTO.xlsx");


        //    //especiales arrendamiento suma

        //    //var query2 = ajustes_especiales.GroupBy(d => d.PCS)
        //    // .Select(
        //    //         g => new
        //    //         {
        //    //             Key = g.Key,
        //    //             Monto_Total = g.Sum(s => s.montoTotal),
        //    //             Pcs = g.First().PCS,
        //    //             cc = g.First().CuentaFinanciera,
        //    //             codigo_cargo = g.First().codigodecargo
        //    //         });


        //    //Console.WriteLine("Cargos Actuales: {0}", cargos.Count());
        //    //Console.WriteLine("------------------------------");

        //    //var miTable600 = new DataTable("CARGOS ARRENDAMIENTO AJUSTES ESPECIALES");
        //    //miTable600.Columns.Add("PCS");
        //    //miTable600.Columns.Add("CUENTA");
        //    //miTable600.Columns.Add("CODIGO DE CARGO");
        //    //miTable600.Columns.Add("AMOUNT");
        //    //miTable600.Columns.Add("CICLO");
        //    //i = 90;
        //    //worker.ReportProgress(i);

        //    //foreach (var item in query2)
        //    //{
        //    //    miTable600.Rows.Add(new Object[] { item.Pcs, item.cc, item.codigo_cargo, "-" + item.Monto_Total, "" });
        //    //}

        //    //SaveExcel.BuildExcel(miTable600, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\C_ARRENDAMIENTO AJUSTES ESPECIALES.xlsx");

        //    //var miTable7 = new DataTable("ORIGINAL");
        //    //miTable7.Columns.Add("CuentaFinanciera");
        //    //miTable7.Columns.Add("recveivercustomer");
        //    //miTable7.Columns.Add("tipocargo");
        //    //miTable7.Columns.Add("codigodecargo");
        //    //miTable7.Columns.Add("offwer");
        //    //miTable7.Columns.Add("nombreOFFER");
        //    //miTable7.Columns.Add("promo");
        //    //miTable7.Columns.Add("description");
        //    //miTable7.Columns.Add("montoCargos");
        //    //miTable7.Columns.Add("montoDescuentos");
        //    //miTable7.Columns.Add("montoTotal");
        //    //miTable7.Columns.Add("tipodeCobro");
        //    //miTable7.Columns.Add("PCS");
        //    //miTable7.Columns.Add("secuenciadeCiclo");
        //    //miTable7.Columns.Add("SEGMENTO");
        //    //miTable7.Columns.Add("prorrateo");
        //    //miTable7.Columns.Add("RUT");


        //    //foreach (var item in cargos)
        //    //{
        //    //    miTable7.Rows.Add(new Object[] {item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //    //        item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
        //    //        , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut});
        //    //}

        //    //i = 95;
        //    //worker.ReportProgress(i);
        //    //oSLDocument.SelectWorksheet("Sheet1");
        //    //oSLDocument.ImportDataTable(1, 1, miTable7, true);
        //    //oSLDocument.RenameWorksheet("Sheet1", "ORIGINAL");
        //    //oSLDocument.SaveAs(@quedaran_aqui_txt.Text + @"\Ciclo\Extraidas" + @"\EXTRAIDOS CARGOS DE ARRENDAMIENTO.xlsx");
        //    ////falto crear excel en ruta
        //    //SaveExcel.BuildExcel(miTable7, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta" + @"\C_ARRENDAMIENTO.xlsx"); //este no es el que tengo que imprimir
        //    i = 100;
        //    worker.ReportProgress(i);
        //    //creo que ta listo

        //    MessageBox.Show("Proceso Terminado");

        //}
        private void bw_DoWork_ejecutar(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            Controlador_Arrendamiento controlador_Arrendamiento = new Controlador_Arrendamiento(cargos_arrendamiento_txt.Text, ajustes_especiales_Txt.Text,@e_comerce_txt.Text, @quedaran_aqui_txt.Text, cartera_Empresarials_sice, carteras_pablo, tipos_);
            int i = 5;
            worker.ReportProgress(i);
            //
            //
            //if (controlador_Arrendamiento.corroboracion_de_colecciones_0()) //aqui en el otro lado hay que hacer una corroboracion de uno en uno mas personalizada o si no queda la caga, teni que ponerte las pilas marcelo o te van a echar y en este minuto eres la esperanza de tu familia para poder obtener la vivienda, no por tu dinero y no por nada especial tuyo, solo por la antiguedad, no te pueden echar, no te pueden echar...
            //{
            //    DialogResult dialogResult = MessageBox.Show("¿Estás seguro de continuar?", "Error de Archivos", MessageBoxButtons.YesNo);
            //    if (dialogResult == DialogResult.Yes)
            //    {
                    
            //    }
            //    else if (dialogResult == DialogResult.No)
            //    {
            //        i = 0;
            //        worker.ReportProgress(i);
            //        return;
            //    }

            //}
            i = 10;
            worker.ReportProgress(i);
            CreacionCarpetas();
            i = 20;
            worker.ReportProgress(i);
            controlador_Arrendamiento.proceso_arrendamiento_empresas_1();
            i = 30;
            worker.ReportProgress(i);
            controlador_Arrendamiento.proceso_arrendamiento_neteados_2();
            i = 40;
            worker.ReportProgress(i);
            controlador_Arrendamiento.proceso_arrendamiento_ecomerce_4();
            controlador_Arrendamiento.proceso_arrendamiento_ajustes_3();
            i = 60;
            worker.ReportProgress(i);
            
            controlador_Arrendamiento.proceso_arrendamiento_duplicados_5();
            i = 70;
            worker.ReportProgress(i);
            i = 90;
            controlador_Arrendamiento.proceso_arrendamiento_4_5();
            worker.ReportProgress(i);
            controlador_Arrendamiento.proceso_arrendamiento_guardando_6();
            //controlador_Arrendamiento.copiar_pegar();
            i = 100;
            worker.ReportProgress(i);
            //creo que ta listo

            //MessageBox.Show("Proceso Terminado");

        }

        public void CreacionCarpetas()
        {
            Console.WriteLine("Creando el directorio: {0}", @quedaran_aqui_txt.Text);
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

        
        private void ajustes_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
           ajustes_especiales_Txt.Text = aux_path;
        }
        //private void btn_Guardar_Click(object sender, EventArgs e)
        //{
        //    for (int i = 0; i < segmento_dgv.RowCount - 1; i++)
        //    {
        //        if (segmento_dgv.Rows[i].Cells["check"].Value.Equals(true))
        //        {
        //            tipos_.Add(segmento_dgv.Rows[i].Cells["Tipo"].Value.ToString());
        //        }
        //    }

        //    segmento_dgv.Visible = false;
        //    btn_Guardar.Visible = false;
        //    string message = "Se agregaron los tipos";
        //    MessageBox.Show(message);
        //    ejecutar_Btn.Visible = true;
        //}

        private void seleccionar_carpeta_Btn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            quedaran_aqui_txt.Text = aux_path;
        }
        private void ejecutar_Btn_Click(object sender, EventArgs e)
        {

            if (!this.bw_ejecutar.IsBusy)
            {
                this.bw_ejecutar.RunWorkerAsync();
                this.ejecutar_Btn.Enabled = false;
            }
        }

        private void seleccionar_carpeta_Btn_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            quedaran_aqui_txt.Text = aux_path;
        }

        private void activacion_Btn_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivos Txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cargos_arrendamiento_txt.Text = aux_path;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            e_comerce_txt.Text = aux_path;
        }

    }
}
