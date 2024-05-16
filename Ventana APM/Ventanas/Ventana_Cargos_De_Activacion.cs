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
    public partial class Ventana_Cargos_De_Activacion : Form
    {
        string aux_path = string.Empty;
        Coleccion_Activacion coleccion_Activacion = new Coleccion_Activacion();
        Coleccion_Cartera_Empresarial _Cartera_Empresarial = new Coleccion_Cartera_Empresarial();
        Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
        List<Activacion> activacion_empresas = new List<Activacion>();
        List<Portabilidad> portabilidads;
        Coleccion_Portabilidad coleccion_Portabilidad = new Coleccion_Portabilidad();
        List<FVM> fVMs;
        Coleccion_FVM coleccion_FVM = new Coleccion_FVM();
        List<Activacion> activacion_neteados = new List<Activacion>();
        List<HABILITACIONES> hABILITACIONEs;
        Coleccion_HABILITACIONES coleccion_HABILITACIONES = new Coleccion_HABILITACIONES();
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
        SLDocument oSLDocument = new SLDocument();
        private BackgroundWorker bw_ejecutar;
        public Ventana_Cargos_De_Activacion(List<Cartera_Empresarial> cartera_Empresarials, List<Cartera> carteras, List<string> filtro)
        {
            InitializeComponent();
            cartera_Empresarials_sice = cartera_Empresarials;
            carteras_pablo = carteras;
            tipos_ = filtro;
            ejecutar_Btn.Visible = true;
            this.bw_ejecutar = new BackgroundWorker();
            this.bw_ejecutar.DoWork += new DoWorkEventHandler(bw_DoWork_ejecutar);//este sirve para los metodos
            this.bw_ejecutar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_ejecutar);//este envia el progreso
            this.bw_ejecutar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_ejecutar);//este cuando termina
            this.bw_ejecutar.WorkerReportsProgress = true;

            this.ejecutar_Btn.Click += new EventHandler(ejecutar_Btn_Click);

            panel12.Visible = false;
        }

        private void bw_ProgressChanged_analizar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;//izi pizzi funcionando el progress bar

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
                sb.AppendLine("Comparando Cargos de Activacion con Carteras Empresariales...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 30)
            {
                sb.AppendLine("Cargos de Activacion Neteados y Portados...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 40)
            {
                sb.AppendLine("Comparando solidarios...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 50)
            {
                sb.AppendLine("Comparando FVMS...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 70)
            {
                sb.AppendLine("Comparando Habilitaciones...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 80)
            {
                sb.AppendLine("Comparando Duplicados...");
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

        private void bw_DoWork_ejecutar(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            int i = 10;
            worker.ReportProgress(i);
            CreacionCarpetas();
            i = 20;
            worker.ReportProgress(i);
            activacions = coleccion_Activacion.GenerarListado(@activacion_txt.Text);
            Controlador_Activacion controlador_Activacion = new Controlador_Activacion(@activacion_txt.Text,@portabilidad_txt.Text,@fvm_txt.Text,@habilitaciones_txt.Text,@solidarios_Txt.Text,@quedaran_aqui_txt.Text,@consolidado_txt.Text, cartera_Empresarials_sice, carteras_pablo, tipos_,double.Parse(monto_filtrado.Text));
            controlador_Activacion.proceso_activacion_empresas_1();
            i = 30;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_activacion_portados_2();
            //
            //AQUI VOY
            i = 40;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_activacion_solidarios_3();
            i = 50;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_activacion_fvm_4();//
            i = 70;
            worker.ReportProgress(i);//
            controlador_Activacion.proceso_activacion_habilitaciones_5();
            //Habilitaciones();
            i = 80;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_activacion_duplicados_6();
            ///controlador_Activacion.proceso_auxiliar();
            i = 90;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_consolidacion_6_5();
            i = 94;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_consolidacion_6_6_portados();
            i = 98;
            worker.ReportProgress(i);
            controlador_Activacion.proceso_activacion_guardando_7();

            //try
            //{
            //    int hola = int.Parse("hola");
            //}
            //catch (Exception ex)
            //{
                
            //    MessageBox.Show("El error es " + ex);
            //}
            i = 100;
            worker.ReportProgress(i);
            // MessageBox.Show("Listo!");

        }



        private void activacion_Btn_Click(object sender, EventArgs e)
        {
            aux_path = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivo txt (*.txt)|*.txt|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            activacion_txt.Text = aux_path;
        }


        public void CreacionCarpetas()
        {
            Console.WriteLine("Creando el directorio: {0}", @quedaran_aqui_txt.Text);
            carpeta_final = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Final";
            carpeta_ruta = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta";
            carpeta_extraidas = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas";
            carpeta_respaldo = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo";
            carpeta_asignar  = @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Asignar";
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
        private void seleccionar_carpeta_Btn_Click(object sender, EventArgs e)
        {
            aux_path = string.Empty;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                aux_path = fbd.SelectedPath.ToString();
            quedaran_aqui_txt.Text = aux_path;
        }

        private void portabilidad_Btn_Click(object sender, EventArgs e)
        {
            aux_path = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Archivo txt (*.txt)|*.txt|All files (*.*)|*.*";
            
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            portabilidad_txt.Text = aux_path;
        }


        private void fvm_Btn_Click(object sender, EventArgs e)
        {
            aux_path = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            fvm_txt.Text = aux_path;
        }

        /// <summary>
        private void habilitaciones_Btn_Click(object sender, EventArgs e)
        {
            aux_path = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            habilitaciones_txt.Text = aux_path;
        }

        private void solidarios_Btn_Click(object sender, EventArgs e)
        {
            aux_path = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "txt";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            solidarios_Txt.Text = aux_path;
        }

        private void ejecutar_Btn_Click(object sender, EventArgs e)
        {
            if (!this.bw_ejecutar.IsBusy)
            {
                this.bw_ejecutar.RunWorkerAsync();
                this.ejecutar_Btn.Enabled = false;
            }


        }

        //private void Empresas_Activacion()
        //{
        //    foreach (var item1 in activacions)
        //    {

        //        foreach (var item2 in cartera_Empresarials_sice) //cartera empresarial sice s ecruza por rut y se debe traer el segmento
        //        {
        //            if (item1.rut.Equals(item2.RUT))
        //            {
        //                foreach (var item3 in tipos_)
        //                {
        //                    if (item2.SEGMENTO.Equals(item3))
        //                    {
        //                        item1.segmento = item2.SEGMENTO;
        //                        activacion_empresas.Add(item1);
        //                        break;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    foreach (var item in activacion_empresas)
        //    {
        //        activacions.Remove(item);
        //    }


        //    foreach (var item1 in activacions)
        //    {

        //        foreach (var item2 in carteras_pablo)
        //        {
        //            if (item1.CuentaFinanciera.Equals(item2.BA_NO))
        //            {
        //                foreach (var item3 in tipos_)
        //                {
        //                    if (item2.SEGMENTO.Equals(item3))
        //                    {
        //                        item1.segmento = item2.SEGMENTO;
        //                        activacion_empresas.Add(item1);
        //                        break;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    foreach (var item in activacion_empresas)
        //    {
        //        activacions.Remove(item);
        //    }

        //    var miTable = new DataTable("EMPRESAS");
        //    miTable.Columns.Add("CuentaFinanciera");
        //    miTable.Columns.Add("recveivercustomer");
        //    miTable.Columns.Add("tipocargo");
        //    miTable.Columns.Add("codigodecargo");
        //    miTable.Columns.Add("offwer");
        //    miTable.Columns.Add("nombreOFFER");
        //    miTable.Columns.Add("promo");
        //    miTable.Columns.Add("description");
        //    miTable.Columns.Add("montoCargos");
        //    miTable.Columns.Add("montoDescuentos");
        //    miTable.Columns.Add("montoTotal");
        //    miTable.Columns.Add("tipodeCobro");
        //    miTable.Columns.Add("PCS");
        //    miTable.Columns.Add("secuenciadeCiclo");
        //    miTable.Columns.Add("SEGMENTO");
        //    miTable.Columns.Add("prorrateo");
        //    miTable.Columns.Add("RUT");


        //    foreach (var item in activacion_empresas)
        //    {
        //        miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //            item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal.ToString() , item.tipodeCobro
        //            , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut});
        //    }



        //    oSLDocument.AddWorksheet("EMPRESAS");
        //    oSLDocument.ImportDataTable(1, 1, miTable, true);


        //    //ta wenooo el codigosss
        //    //
        //    // Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
        //    //SaveExcel.BuildExcel(miTable, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo\Cargos de Activacion" + @"\ACTIVACION EMPRESAS.xlsx");
        //    //SaveExcel.BuildExcel(miTable, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas\Cargos de Activacion\CARGOS DE ACTIVACION CICLO MES_ANIO" + @"\ACTIVACION EMPRESAS.xlsx");
        //}

        //private void Neteados_Activacion()
        //{

        //    portabilidads = coleccion_Portabilidad.GenerarListado(portabilidad_txt.Text);

        //    foreach (var item1 in activacions)
        //    {
        //        string TIPO = string.Empty;
        //        foreach (var item2 in portabilidads)
        //        {
        //            if (item1.PCS.Equals(item2.NRO_DE_PCS))
        //            {
        //                TIPO = item2.MODALIDAD;
        //            }
        //        }
        //        item1.TIPO = TIPO;
        //        if (TIPO.Equals(""))
        //        {
        //            item1.TIPO = "N/A";
        //        }
        //    }
        //    int contador_1 = 0;
        //    int contador_2 = 0;
        //    int contador_3 = 0;

        //    foreach (var item in activacions)
        //    {
        //        if (item.TIPO.Equals("PRE A POST"))
        //        {
        //            contador_1++;
        //        }else if (item.TIPO.Equals("PORTADO"))
        //        {
        //            contador_2++;
        //        }else if (item.TIPO.Equals("N/A"))
        //        {
        //            contador_3++;
        //        }
        //    }


        //    var query = activacions.GroupBy(d => d.PCS)
        //                                .Select(
        //                                        g => new
        //                                        {
        //                                            Key = g.Key,
        //                                            Monto_Total = g.Sum(s => s.montoTotal),
        //                                            CuentaFinanciera = g.First().CuentaFinanciera,
        //                                            Pcs = g.First().PCS,
        //                                            CodigoCarga = g.First().codigodecargo
        //                                        });

        //    int contador_4 = 0;
        //    int contador_de_zeros = 0;
        //    int cuantos_del_sum = query.Count();

        //    foreach (var item1 in query)
        //    {
        //        if (item1.Monto_Total == 0 || item1.Monto_Total == 0.0)
        //        {
        //            contador_de_zeros++;
        //        }
        //    }
        //    foreach (var item1 in query)
        //    {
        //        if (item1.Monto_Total == 0 || item1.Monto_Total == 0.0)
        //        {
        //            foreach (var item2 in activacions)
        //            {
        //                if (item1.Pcs.Equals(item2.PCS))
        //                {
        //                    item2.TIPO = "NETEADO";
        //                    contador_4++;
        //                }
        //            }
        //        }

        //    }




            


        //    foreach (var item1 in activacions)
        //    {
        //        if (item1.TIPO.Equals("NETEADO"))
        //        {
        //            activacion_neteados.Add(item1);
        //        }
        //    } //AQUI CREAR C_ACTIVACION



        //    var miTable9 = new DataTable("NETEADOS");
        //    miTable9.Columns.Add("CuentaFinanciera");
        //    miTable9.Columns.Add("recveivercustomer");
        //    miTable9.Columns.Add("tipocargo");
        //    miTable9.Columns.Add("codigodecargo");
        //    miTable9.Columns.Add("offwer");
        //    miTable9.Columns.Add("nombreOFFER");
        //    miTable9.Columns.Add("promo");
        //    miTable9.Columns.Add("description");
        //    miTable9.Columns.Add("montoCargos");
        //    miTable9.Columns.Add("montoDescuentos");
        //    miTable9.Columns.Add("montoTotal");
        //    miTable9.Columns.Add("tipodeCobro");
        //    miTable9.Columns.Add("PCS");
        //    miTable9.Columns.Add("secuenciadeCiclo");
        //    miTable9.Columns.Add("SEGMENTO");
        //    miTable9.Columns.Add("prorrateo");
        //    miTable9.Columns.Add("RUT");
        //    miTable9.Columns.Add("TIPO");

        //    foreach (var item in activacions)
        //    {
        //        miTable9.Rows.Add(new Object[] {item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //            item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
        //            , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut,item.TIPO});
        //    }

        //    SaveExcel.BuildExcel(miTable9, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta\"+ @"\C_ACTIVACION.xlsx"); //todo correcto hasta aqui
        //    //VA EN RUTA

        //    foreach (var item in activacion_neteados)
        //    {
        //        activacions.Remove(item);
        //    }

        //    var miTable3 = new DataTable("NETEADOS");
        //    miTable3.Columns.Add("CuentaFinanciera");
        //    miTable3.Columns.Add("recveivercustomer");
        //    miTable3.Columns.Add("tipocargo");
        //    miTable3.Columns.Add("codigodecargo");
        //    miTable3.Columns.Add("offwer");
        //    miTable3.Columns.Add("nombreOFFER");
        //    miTable3.Columns.Add("promo");
        //    miTable3.Columns.Add("description");
        //    miTable3.Columns.Add("montoCargos");
        //    miTable3.Columns.Add("montoDescuentos");
        //    miTable3.Columns.Add("montoTotal");
        //    miTable3.Columns.Add("tipodeCobro");
        //    miTable3.Columns.Add("PCS");
        //    miTable3.Columns.Add("secuenciadeCiclo");
        //    miTable3.Columns.Add("SEGMENTO");
        //    miTable3.Columns.Add("prorrateo");
        //    miTable3.Columns.Add("RUT");
        //    miTable3.Columns.Add("TIPO");
        //    miTable3.Columns.Add("COBRAR");
        //    miTable3.Columns.Add("MONTO");
        //    miTable3.Columns.Add("CUOTA");
        //    miTable3.Columns.Add("OBSERVACION");

        //    foreach (var item in activacion_neteados)
        //    {
        //        miTable3.Rows.Add(new Object[] {item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
        //            item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
        //            , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo, item.rut,item.TIPO,"","","",""});
        //    }


            


        //    oSLDocument.AddWorksheet("NETEADOS");
        //    oSLDocument.ImportDataTable(1, 1, miTable3, true);
        //    contador_1 = 0;
        //}
        private void Fvm_Activacion()
        {

            fVMs = coleccion_FVM.GenerarListado(@fvm_txt.Text);


            List<Activacion> activacion_fvm = new List<Activacion>();

            List<FVM> cruce_con_activacion = new List<FVM>();

            foreach (var item_1 in activacions)
            {
                foreach (var item_2 in fVMs)
                {
                    if (item_1.PCS.Equals(item_2.PCS))
                    {
                        item_1.COBRAR = item_2.MONTO_SIN_IMPUESTO;
                        activacion_fvm.Add(item_1);
                        cruce_con_activacion.Add(item_2);
                        break;
                    }
                }
            }

            foreach (var item in activacion_fvm)
            {
                activacions.Remove(item);
            }

            var miTable5 = new DataTable("FVMS");
            miTable5.Columns.Add("CuentaFinanciera");
            miTable5.Columns.Add("RECEIVER_CUSTOMER");
            miTable5.Columns.Add("TIPO_CARGO");
            miTable5.Columns.Add("Codigo de Cargo");
            miTable5.Columns.Add("OFFER");
            miTable5.Columns.Add("Nombre OFFER");
            miTable5.Columns.Add("PROMO");
            miTable5.Columns.Add("DESCRIPTION");
            miTable5.Columns.Add("MONTO_CARGOS");
            miTable5.Columns.Add("MONTO_DESCUENTOS");
            miTable5.Columns.Add("MONTO_TOTAL");
            miTable5.Columns.Add("Tipo de Cobro");
            miTable5.Columns.Add("PCS");
            miTable5.Columns.Add(" Secuencia de Ciclo");
            miTable5.Columns.Add("PRORRATEO");
            miTable5.Columns.Add("RUT");
            miTable5.Columns.Add("TIPO");
            miTable5.Columns.Add("COBRAR");
            miTable5.Columns.Add("MONTO");
            miTable5.Columns.Add("CUOTA");
            miTable5.Columns.Add("OBSERVACION");


            foreach (var item in activacion_fvm) //falta el cruce por las otras dos hojas
            {
                miTable5.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo, item.rut,item.MODALIDAD,item.COBRAR,"","",""});
            }

            oSLDocument.AddWorksheet("FVM");
            oSLDocument.ImportDataTable(1, 1, miTable5, true); //porque mierda no agrega esta hoja

            //imprimir fvm en la carpeta ciclo

            foreach (var item in cruce_con_activacion)
            {
                fVMs.Remove(item);
            }

            var miTable9 = new DataTable("FVM");       

            miTable9.Columns.Add("CCOCSUBEQPCOM");
            miTable9.Columns.Add("PCS");
            miTable9.Columns.Add("MONTO_SIN_IMPUESTO");
            miTable9.Columns.Add("INGRESO_VTA");
            miTable9.Columns.Add("FA");
            miTable9.Columns.Add("CBP");
            miTable9.Columns.Add("FECHA_ACTIVACION");

            foreach (var item in fVMs)
            {
                miTable9.Rows.Add(new Object[] {item.CCOCSUBEQPCOM,item.PCS,item.MONTO_SIN_IMPUESTO,item.INGRESO_VTA,item.FA,item.CBP, item.FECHA_ACTIVACION});
            }

            SaveExcel.BuildExcel(miTable9, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\CICLO FVM.xlsx");
            //SaveExcel.BuildExcel(miTable5, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Respaldo\Cargos de Activacion" + @"\ACTIVACION FVMS.xlsx"); //...

            //SaveExcel.BuildExcel(miTable5, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas\Cargos de Activacion\CARGOS DE ACTIVACION CICLO MES_ANIO" + @"\ACTIVACION FVMS.xlsx");

            List<FVM> fvm_activacion = new List<FVM>();

            //cruce inverso
            foreach (var item_1 in fVMs)
            {
                foreach (var item_2 in activacion_fvm)
                {
                    if (item_1.PCS.Equals(item_2.PCS))
                    {
                        fvm_activacion.Add(item_1);
                        break;
                    }
                }
            }

            foreach (var item in fvm_activacion)
            {
                fVMs.Remove(item);
            }

            List<FVM> fvm_activacion_neteados = new List<FVM>();

            foreach (var item_1 in fVMs)
            {
                foreach (var item_2 in activacion_neteados)
                {
                    if (item_1.PCS.Equals(item_2.PCS))
                    {
                        fvm_activacion_neteados.Add(item_1);
                        break;
                    }
                }
            }

            foreach (var item in fvm_activacion_neteados)
            {
                fVMs.Remove(item);
            }


            var miTable10 = new DataTable("CICLO X NETEADOS");

            miTable10.Columns.Add("CCOCSUBEQPCOM");
            miTable10.Columns.Add("MONTO_SIN_IMPUESTO");
            miTable10.Columns.Add("PCS");
            miTable10.Columns.Add("INGRESO_VTA");
            miTable10.Columns.Add("FA");
            miTable10.Columns.Add("CBP");
            miTable10.Columns.Add("FECHA_ACTIVACION");



            foreach (var item in fVMs)
            {
                miTable10.Rows.Add(new Object[] { item.CCOCSUBEQPCOM, item.MONTO_SIN_IMPUESTO, item.PCS, item.INGRESO_VTA, item.FA, item.CBP, item.FECHA_ACTIVACION });
            }

        } //terminar completo falta cruce inverso para enviar por correo

        private void Habilitaciones()
        {

            hABILITACIONEs = coleccion_HABILITACIONES.GenerarListado(@habilitaciones_txt.Text);
            List<Activacion> activacion_habilitaciones = new List<Activacion>();

            foreach (var item_1 in activacions)
            {
                foreach (var item_2 in hABILITACIONEs)
                {
                    if (item_1.PCS.Equals(item_2.PCS))
                    {
                        item_1.COBRAR = item_2.MONTO_SIN_IMPUESTO;
                        activacion_habilitaciones.Add(item_1);
                        break;
                    }
                }
            }

            foreach (var item in activacion_habilitaciones)
            {
                activacions.Remove(item);
            }

            List<HABILITACIONES> hABILITACIONEs1 = new List<HABILITACIONES>();
            int CONTADOR = 0;
            foreach (var item in hABILITACIONEs)
            {
                CONTADOR++;
                if (CONTADOR > 1)
                {
                    if (item.MONTO_SIN_IMPUESTO >= 0.0) //ACA HAY ALGO QUE ESTA MAL REVISAR MARCELO DE LAS 7:00
                    {
                        Activacion anticuota_ = new Activacion();
                        anticuota_.tipocargo = activacions[3].tipocargo;
                        anticuota_.codigodecargo = activacions[3].codigodecargo;
                        anticuota_.offwer = activacions[3].offwer;
                        anticuota_.nombreOFFER = activacions[3].nombreOFFER;
                        anticuota_.promo = activacions[3].promo;
                        anticuota_.description = activacions[3].description;
                        anticuota_.montoCargos = 0.0;
                        anticuota_.montoDescuentos = 0.0;
                        anticuota_.montoTotal = 0.0;
                        anticuota_.tipodeCobro = "CLP";
                        anticuota_.PCS = item.PCS;
                        anticuota_.secuenciadeCiclo = activacions[3].secuenciadeCiclo;
                        anticuota_.TIPO = item.INGRESO_VTA;
                        anticuota_.COBRAR = item.MONTO_SIN_IMPUESTO;
                        anticuota_.COBRAR_SI = "SI";
                        anticuota_.CUOTA = "1 de 1";



                        anticuota_.CuentaFinanciera = item.FA;
                        anticuota_.recveivercustomer = item.CBP;
                        activacion_habilitaciones.Add(anticuota_);
                        hABILITACIONEs1.Add(item);
                    }
                }

            }

            foreach (var item in hABILITACIONEs1)
            {
                hABILITACIONEs.Remove(item);
            }

           
            var miTable12 = new DataTable("HABILITACIONES");
            miTable12.Columns.Add("CuentaFinanciera");
            miTable12.Columns.Add("RECEIVER_CUSTOMER");
            miTable12.Columns.Add("TIPO_CARGO");
            miTable12.Columns.Add("Codigo de Cargo");
            miTable12.Columns.Add("OFFER");
            miTable12.Columns.Add("Nombre OFFER");
            miTable12.Columns.Add("PROMO");
            miTable12.Columns.Add("DESCRIPTION");
            miTable12.Columns.Add("MONTO_CARGOS");
            miTable12.Columns.Add("MONTO_DESCUENTOS");
            miTable12.Columns.Add("MONTO_TOTAL");
            miTable12.Columns.Add("Tipo de Cobro");
            miTable12.Columns.Add("PCS");
            miTable12.Columns.Add(" Secuencia de Ciclo");
            miTable12.Columns.Add("PRORRATEO");
            miTable12.Columns.Add("RUT");
            miTable12.Columns.Add("TIPO");
            miTable12.Columns.Add("COBRAR");
            miTable12.Columns.Add("MONTO");
            miTable12.Columns.Add("CUOTA");
            miTable12.Columns.Add("OBSERVACION");


            foreach (var item in activacion_habilitaciones)
            {
                if (activacion_habilitaciones.Count==0)
                {
                    miTable12.Rows.Add(new Object[] { "", "", "","", "",
                    "", "", "", "", "", "" , ""
                    , "", "", "", "","","","","",""});
                }
                else
                {
                    miTable12.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo, item.rut,item.MODALIDAD,item.COBRAR_SI,item.COBRAR,item.CUOTA,""});
                }
            }

            oSLDocument.AddWorksheet("HABILITACIONES");
            oSLDocument.ImportDataTable(1, 1, miTable12, true);
        } //terminar completo falta cruce inverso para enviar por correo

        private void Duplicados()
        {
            List<Activacion> acticuota_duplicados = new List<Activacion>();
            int repetidos2 = 0;
            //var duplicados = activacions.GroupBy(x => x.PCS)
            //                                .Where(g => g.Count() > 1)
            //                                .Select(g => g.Key);
            //int duplicadosnumero = duplicados.Count();

            int repetido1 = 0;
            int repetido2 = 0;
            foreach (var item1 in activacions)
            {
                foreach (var item2 in activacions)
                {
                    if (item1.PCS.Equals(item2.PCS))
                    {
                        repetido1++;
                    }
                }
                if (repetido1 > 1)
                {
                    acticuota_duplicados.Add(item1);
                }
                repetido1 = 0;
            }

            //foreach (var item1 in activacions)
            //{
            //    repetidos2 = 0;
            //    foreach (var item2 in duplicados)
            //    {
            //        if (item1.PCS.Equals(item2))
            //        {
            //            repetidos2++;
            //            acticuota_duplicados.Add(item1);
            //        }
            //    }

            //}

            foreach (var item in acticuota_duplicados)//HASTA AQUI TODO CORRECTO Y YO QUE ME ALEGRO
            {
                activacions.Remove(item);
            }


            List<Activacion> acticuota_duplicados_verdadero_falso = new List<Activacion>();
            int contador_verdadero = 0;
            int contador_falso = 0;
            foreach (var item in acticuota_duplicados)
            {
                if (item.montoTotal > 0)
                {
                    item.verdaredo_falso = "VERDADERO";
                    acticuota_duplicados_verdadero_falso.Add(item);
                    contador_verdadero++;
                }
                else
                {
                    item.verdaredo_falso = "FALSO";
                    acticuota_duplicados_verdadero_falso.Add(item);
                    contador_falso++;
                }
            }

            List<Activacion> acticuota_duplicados_verdadero_repetidos = new List<Activacion>();

            List<Activacion> acticuota_duplicados_verdadero = new List<Activacion>();
            List<Activacion> acticuota_duplicados_falso = new List<Activacion>();
            List<Activacion> acticuota_duplicados_verdadero_falso_repetidos = new List<Activacion>();
            

            foreach (var item in acticuota_duplicados_verdadero_falso)
            {
                if (item.verdaredo_falso.Equals("VERDADERO"))
                {
                    acticuota_duplicados_verdadero.Add(item);
                }
                else
                {
                    acticuota_duplicados_falso.Add(item);
                }
            }

            foreach (var item1 in acticuota_duplicados_verdadero)
            {
                foreach (var item2 in acticuota_duplicados_verdadero)
                {
                    if (item1.PCS.Equals(item2.PCS))
                    {
                        repetido1++;
                    }
                }
                if (repetido1 > 1)
                {
                    acticuota_duplicados_verdadero_repetidos.Add(item1);
                }
                repetido1 = 0;
            }

            foreach (var item1 in acticuota_duplicados_verdadero_repetidos)//esto no esta bien hay que arregñlarlo
            {
                foreach (var item2 in acticuota_duplicados_falso)
                {
                    if (item1.PCS.Equals(item2.PCS))
                    {
                        acticuota_duplicados_verdadero_falso_repetidos.Add(item2);
                    }
                    
                }
            }


            

            //foreach (var item1 in acticuota_duplicados_verdadero_falso) //esto esta mal
            //{
            //    repetidos2 = 0;
            //    foreach (var item2 in acticuota_duplicados_verdadero_falso)
            //    {
            //        if (item1.PCS.Equals(item2.PCS))
            //        {

            //            if (item1.verdaredo_falso.Equals(item2.verdaredo_falso))
            //            {
            //                if (item1.verdaredo_falso.Equals("VERDADERO"))
            //                {
            //                    repetidos2++;
            //                }
            //            }
            //        }

            //    }
            //    if (repetidos2 >= 2)
            //    {
            //        acticuota_duplicados_verdadero_falso_repetidos.Add(item1);
            //    }
            //}
            List<Activacion> acticuota_duplicados_verdadero_falso_repetidos_2 = new List<Activacion>(); //nose que hago aqui
            foreach (var item1 in acticuota_duplicados_verdadero_falso)
            {
                repetidos2 = 0;
                foreach (var item2 in acticuota_duplicados_verdadero_falso_repetidos)
                {
                    if (item1.PCS.Equals(item2.PCS))
                    {
                        repetidos2++;
                    }

                }
                if (repetidos2 >= 2)
                {
                    acticuota_duplicados_verdadero_falso_repetidos_2.Add(item1);
                }
            }

            foreach (var item in activacions)
            {
                if (item.montoTotal < 26000.0)
                {
                    item.verdaredo_falso = "";
                    acticuota_duplicados_verdadero_falso_repetidos_2.Add(item);
                }
            }

            foreach (var item in acticuota_duplicados_verdadero_falso_repetidos_2)
            {
                acticuota_duplicados.Remove(item);
            }

            foreach (var item in acticuota_duplicados)
            {
                activacions.Remove(item);
            }

            var miTable17 = new DataTable("DUPLICADOS");
            miTable17.Columns.Add("CuentaFinanciera");
            miTable17.Columns.Add("RECEIVER_CUSTOMER");
            miTable17.Columns.Add("TIPO_CARGO");
            miTable17.Columns.Add("Codigo de Cargo");
            miTable17.Columns.Add("OFFER");
            miTable17.Columns.Add("Nombre OFFER");
            miTable17.Columns.Add("PROMO");
            miTable17.Columns.Add("DESCRIPTION");
            miTable17.Columns.Add("MONTO_CARGOS");
            miTable17.Columns.Add("MONTO_DESCUENTOS");
            miTable17.Columns.Add("MONTO_TOTAL");
            miTable17.Columns.Add("VERDADERO O FALSO");
            miTable17.Columns.Add("Tipo de Cobro");
            miTable17.Columns.Add("PCS");
            miTable17.Columns.Add(" Secuencia de Ciclo");
            miTable17.Columns.Add("PRORRATEO");
            miTable17.Columns.Add("RUT");
            miTable17.Columns.Add("TIPO");




            foreach (var item in acticuota_duplicados_verdadero_falso_repetidos_2)
            {
                miTable17.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal,item.verdaredo_falso , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo, item.rut,item.TIPO});
            }
            //esto va al nuevo excel



            SaveExcel.BuildExcel(miTable17, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Cargos de Activacion.xlsx");// este es el excel nuevo





            List<Activacion> activacion_portados = new List<Activacion>();

            foreach (var item in acticuota_duplicados)
            {
                if (item.TIPO.Equals("PORTADO"))
                {
                    activacion_portados.Add(item);
                }
            }

            foreach (var item in activacion_portados)
            {
                acticuota_duplicados.Remove(item);
            }




            var miTable18 = new DataTable("CARGOS DE ACTIVACION");
            miTable18.Columns.Add("CuentaFinanciera");
            miTable18.Columns.Add("RECEIVER_CUSTOMER");
            miTable18.Columns.Add("TIPO_CARGO");
            miTable18.Columns.Add("Codigo de Cargo");
            miTable18.Columns.Add("OFFER");
            miTable18.Columns.Add("Nombre OFFER");
            miTable18.Columns.Add("PROMO");
            miTable18.Columns.Add("DESCRIPTION");
            miTable18.Columns.Add("MONTO_CARGOS");
            miTable18.Columns.Add("MONTO_DESCUENTOS");
            miTable18.Columns.Add("MONTO_TOTAL");
            miTable18.Columns.Add("VERDADERO O FALSO");
            miTable18.Columns.Add("Tipo de Cobro");
            miTable18.Columns.Add("PCS");
            miTable18.Columns.Add(" Secuencia de Ciclo");
            miTable18.Columns.Add("PRORRATEO");
            miTable18.Columns.Add("RUT");
            miTable18.Columns.Add("TIPO");


            foreach (var item in activacion_portados)
            {
                miTable18.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal,item.verdaredo_falso , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo, item.rut,item.TIPO});
            }


            SaveExcel.BuildExcel(miTable18, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Extraidas" + @"\PORTADOS CICLO XX_MES_ANIO.xlsx");


            var miTable16 = new DataTable("DUPLICADOS");
            miTable16.Columns.Add("ASIGNACION");
            miTable16.Columns.Add("CuentaFinanciera");
            miTable16.Columns.Add("RECEIVER_CUSTOMER");
            miTable16.Columns.Add("TIPO_CARGO");
            miTable16.Columns.Add("Codigo de Cargo");
            miTable16.Columns.Add("OFFER");
            miTable16.Columns.Add("Nombre OFFER");
            miTable16.Columns.Add("PROMO");
            miTable16.Columns.Add("DESCRIPTION");
            miTable16.Columns.Add("MONTO_CARGOS");
            miTable16.Columns.Add("MONTO_DESCUENTOS");
            miTable16.Columns.Add("MONTO_TOTAL");
            miTable16.Columns.Add("VERDADERO O FALSO");
            miTable16.Columns.Add("Tipo de Cobro");
            miTable16.Columns.Add("PCS");
            miTable16.Columns.Add(" Secuencia de Ciclo");
            miTable16.Columns.Add("PRORRATEO");
            miTable16.Columns.Add("RUT");
            miTable16.Columns.Add("TIPO");


            foreach (var item in acticuota_duplicados)
            {
                miTable16.Rows.Add(new Object[] {"" ,item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal,item.verdaredo_falso , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo, item.rut,item.TIPO});
            }

                // SaveExcel.BuildExcel(miTable16, @quedaran_aqui_txt.Text + @"\Activacion CARGOS DE ACTIVACION.xlsx");

            var miTable100 = new DataTable("DUPLICADOS");
            miTable100.Columns.Add("ASIGNACION");
            miTable100.Columns.Add("CuentaFinanciera");//que a mi lado estas uuuuuuuuuuuuh  
            miTable100.Columns.Add("RECEIVER_CUSTOMER");
            miTable100.Columns.Add("TIPO_CARGO");
            miTable100.Columns.Add("Codigo de Cargo");
            miTable100.Columns.Add("OFFER");
            miTable100.Columns.Add("Nombre OFFER");
            miTable100.Columns.Add("PROMO");
            miTable100.Columns.Add("DESCRIPTION");
            miTable100.Columns.Add("MONTO_CARGOS");
            miTable100.Columns.Add("MONTO_DESCUENTOS");
            miTable100.Columns.Add("MONTO_TOTAL");
            miTable100.Columns.Add("VERDADERO O FALSO");
            miTable100.Columns.Add("Tipo de Cobro");//quiero saber si es verdad
            miTable100.Columns.Add("PCS");
            miTable100.Columns.Add(" Secuencia de Ciclo");
            miTable100.Columns.Add("PRORRATEO");
            miTable100.Columns.Add("RUT");
            miTable100.Columns.Add("TIPO");


            foreach (var item in activacions)
            {
                miTable16.Rows.Add(new Object[] {"" ,item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal,item.verdaredo_falso , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo, item.rut,item.TIPO});
            }

            oSLDocument.AddWorksheet("DUPLICADOS"); //AQUI VAMOS
            oSLDocument.ImportDataTable(1, 1, miTable100, true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();//paso 1
            ofd.DefaultExt = "txt";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            consolidado_txt.Text = aux_path;
        }

        private void panel12_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
