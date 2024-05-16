using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.CONTROLADOR
{
    public class Controlador_DownGrade
    {
        //Down Grade yo solo se decir de nada, por el mar nunca mas jamas 
        //-----------------------------------COLECCIONES---------------------------------------------
        Coleccion_Penalidad coleccion_Penalidad = new Coleccion_Penalidad();
        Coleccion_CFM coleccion_CFM = new Coleccion_CFM();
        Coleccion_PlanSolidario Coleccion_PlanSolidario = new Coleccion_PlanSolidario();
        //------------------------------------------------------------------------------------------

        //---------------------------------LISTAS---------------------------------------------------
        List<Penalidad> penalidads = new List<Penalidad>();
        List<CFM_AJUSTES> cFM_s = new List<CFM_AJUSTES>();
        List<CFM_AJUSTES> cFM_si = new List<CFM_AJUSTES>();
        List<PlanSolidario> planSolidarios = new List<PlanSolidario>();
        List<PlanSolidario> planSolidarios_si = new List<PlanSolidario>();
        //-----------------------------------------------------------------------------------------
        string aux_path = string.Empty;

        ////  //-----------------------CREADOR DEL EXCEL---------------------------------------
        SLDocument oSLDocument = new SLDocument();
        SLDocument oSLDocument2 = new SLDocument();
        public Controlador_DownGrade(string path_penalidad,string path_cfm,string path_solidarios , string path_salida)
        {
            penalidads = coleccion_Penalidad.GenerarListado(@path_penalidad);//GenerarListado(@"C:\Users\Marcelo\Desktop\CICLO 4\DownGrade 04.04.txt");
            cFM_s = coleccion_CFM.GenerarListado(@path_cfm);//GenerarListado(@"C:\Users\Marcelo\Desktop\CICLO 4\CFM CREADO POR MI.xlsx");
            planSolidarios = Coleccion_PlanSolidario.GenerarListado(@path_solidarios);
            aux_path = path_salida; 
        }

        public bool proceso_downgrade()
        {
            try
            {
             

                //foreach (var item in planSolidarios)
                //{
                //    if (item.APLICA_CDP.Equals("SI"))
                //    {
                //        planSolidarios_si.Add(item);
                //    } 


                var penalidad_linq_solidario = from penali in penalidads //porque no dio //aqui mismo entramos el catch por estar en null
                                         join cfm in planSolidarios
                                                on penali.PCS equals cfm.PCS
                                         select new Penalidad
                                         {
                                             Cuenta_Financiera = penali.Cuenta_Financiera,
                                             RECEIVER_CUSTOMER = penali.RECEIVER_CUSTOMER,
                                             TIPO_CARGO = penali.TIPO_CARGO,
                                             Codigo_de_Cargo = penali.Codigo_de_Cargo,
                                             OFFER = penali.OFFER,
                                             Nombre_OFFER = penali.Nombre_OFFER,
                                             PROMO = penali.PROMO,
                                             DESCRIPTION = penali.DESCRIPTION,
                                             MONTO_CARGOS = penali.MONTO_CARGOS,
                                             MONTO_DESCUENTOS = penali.MONTO_DESCUENTOS,
                                             MONTO_TOTAL = penali.MONTO_TOTAL,
                                             Tipo_de_Cobro = penali.Tipo_de_Cobro,
                                             PCS = penali.PCS,
                                             Secuencia_de_Ciclo = penali.Secuencia_de_Ciclo,
                                             PRORRATEO = penali.PRORRATEO,
                                             RUT = penali.RUT
                                         };
                ///------------------------------------------------------------------------------------
                foreach (var item in cFM_s)
                {
                    if (item.APLICA_CDP.Equals("SI"))
                    {
                        cFM_si.Add(item);
                    }
                }
                var penalidad_linq_cfm = from penali in penalidads
                                         join cfm in cFM_si
                                                on penali.PCS equals cfm.PCS
                                         select new Penalidad
                                         {
                                             Cuenta_Financiera = penali.Cuenta_Financiera,
                                             RECEIVER_CUSTOMER = penali.RECEIVER_CUSTOMER,
                                             TIPO_CARGO = penali.TIPO_CARGO,
                                             Codigo_de_Cargo = penali.Codigo_de_Cargo,
                                             OFFER = penali.OFFER,
                                             Nombre_OFFER = penali.Nombre_OFFER,
                                             PROMO = penali.PROMO,
                                             DESCRIPTION = penali.DESCRIPTION,
                                             MONTO_CARGOS = penali.MONTO_CARGOS,
                                             MONTO_DESCUENTOS = penali.MONTO_DESCUENTOS,
                                             MONTO_TOTAL = penali.MONTO_TOTAL,
                                             Tipo_de_Cobro = penali.Tipo_de_Cobro,
                                             PCS = penali.PCS,
                                             Secuencia_de_Ciclo = penali.Secuencia_de_Ciclo,
                                             PRORRATEO = penali.PRORRATEO,
                                             RUT = penali.RUT
                                         };

                int count = penalidad_linq_cfm.Count();


                var dt1 = new DataTable("CARGOS DE ARRENDAMIENTO");//NOMBRE POR AHORA
                dt1.Columns.Add("PCS", typeof(long));
                dt1.Columns.Add("CuentaFinanciera", typeof(long));
                dt1.Columns.Add("Codigo de Cargo", typeof(string));
                dt1.Columns.Add("AMOUNT", typeof(double));
                dt1.Columns.Add("CICLO", typeof(string));
                // O_w_O
                // --|--
                //   |
                //   |
                //   -
                //  / \
                // si duplicados se suman y verificar si me da 0
                //tambien hay que verificar si hay duplicados po mijo obvio
                foreach (var item in penalidad_linq_cfm)
                {
                    if (item.MONTO_TOTAL == 0)
                    {

                    }
                    else
                    {
                        item.MONTO_TOTAL = item.MONTO_TOTAL * -1;

                        dt1.Rows.Add(new Object[] {long.Parse(item.PCS),
                                               long.Parse(item.Cuenta_Financiera),
                                               item.Codigo_de_Cargo,
                                               Math.Round(item.MONTO_TOTAL),
                                               "DIA_MES"});
                    }

                }

                oSLDocument.ImportDataTable(1, 1, dt1, true);
                oSLDocument.RenameWorksheet("Sheet1", "AJUSTES PENALIDAD");
                
                //ring ring ring banana phone ------------------------------------------------------------------------------------
                //
                //----------------------------------------------------------------------------------------------------------------
                
                foreach (var item in penalidad_linq_solidario)
                {
                        item.MONTO_TOTAL = item.MONTO_TOTAL * -1;
                        dt1.Rows.Add(new Object[] {long.Parse(item.PCS),
                                               long.Parse(item.Cuenta_Financiera),
                                               item.Codigo_de_Cargo,
                                               Math.Round(item.MONTO_TOTAL),
                                               "DIA_MES"});
                }



                oSLDocument.ImportDataTable(1, 1, dt1, true);
                oSLDocument.RenameWorksheet("Sheet1", "AJUSTES PENALIDAD");

                oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\Ruta\"+ @"\AJUSTES DONWGRADE _CICLO_X_X.xlsx");//@"C:\Users\Marcelo\Desktop\" + @"\AJUSTES DONWGRADE _CICLO_X_X.xlsx"
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex);
                return false;
            }
        }


    }
}
