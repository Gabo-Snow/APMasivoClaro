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
    public class Controlador_Hbo_Paramount
    {
        Coleccion_Hbo coleccion_hvo             =       new Coleccion_Hbo();
        Coleccion_Paramount colecion_Paramount  =       new Coleccion_Paramount();
        List<Hbo> hbos                          =       new List<Hbo>();
        List<Paramount> paramounts              =       new List<Paramount>();
        List<Hbo> hbos_resultado                =       new List<Hbo>();
        double valor                            =       5798; //Math.Round(6900 / 1.19);//esto es un calor estatico
        SLDocument oSLDocument                  =       new SLDocument();
        SLDocument oSLDocument2                 =       new SLDocument();
        string salida                           =       string.Empty;

        public Controlador_Hbo_Paramount(string path_hbo,string path_para, string path_salida)
        {
            hbos = coleccion_hvo.GenerarListado(@path_hbo);
            paramounts = colecion_Paramount.GenerarListado(@path_para);
            salida = path_salida;
        }

        public bool hbo_31()
        {
            int cuanto      =       0;
            try
            {
                //int currentMonth    =       DateTime.Now.Month;
                //int currentYear     =       DateTime.Now.Year;
                //int days            =       DateTime.DaysInMonth(currentYear, currentMonth);

                int max_day = hbos.Max(t => t.PRORRATEO);
                //(\_ /)
                //( •,•)
                //(")_(")
                List<string> auxiliar = new List<string>();
                foreach (var item in hbos)
                {
                    if (item.PRORRATEO.Equals(max_day))
                    {
                        item.es_31      =        true;
                        auxiliar.Add(item.PCS);
                    }
                }

                foreach (var item1 in auxiliar)
                {
                    foreach (var item2 in hbos)
                    {
                        if (item1.Equals(item2.PCS))
                        {
                            item2.es_31     =       true;
                        }
                    }
                }

                var query = hbos.GroupBy(d => d.PCS)
                .Select(
                        g => new                                               
                        {
                            Key                 =     g.Key,
                            Monto_Total         =     g.Sum(s => s.MONTO_TOTAL),
                            CuentaFinanciera    =     g.First().CuentaFinanciera,
                            Pcs                 =     g.First().PCS,
                            prorateo            =     g.First().PRORRATEO,
                            ultimo_dia          =     g.First().es_31
                        });

                
                foreach (var item in query)
                {
                    Hbo hboss = new Hbo();
                    if (item.ultimo_dia)
                    {
                        hboss.PCS                   =    item.Pcs;
                        hboss.MONTO_TOTAL           =    Math.Round(valor - item.Monto_Total);
                        hboss.CuentaFinanciera      =    item.CuentaFinanciera;
                        //resultado = valor - item.Monto_Total;
                        hbos_resultado.Add(hboss);
                        cuanto++;
                    }
                    else
                    {
                        hboss.PCS               =        item.Pcs;
                        hboss.MONTO_TOTAL       =        Math.Round(item.Monto_Total * -1);
                        hboss.CuentaFinanciera  =        item.CuentaFinanciera;

                        hbos_resultado.Add(hboss);
                        //resultado = ;
                    }
                }


                var dt = new DataTable("HBO");

                dt.Columns.Add("Cuenta Financiera"  ,         typeof(long)      );
                dt.Columns.Add("MOTIVO"             ,         typeof(string)    );//Control HBO
                dt.Columns.Add("OBSERVACIÓN"        ,         typeof(long)      );//PCS
                dt.Columns.Add("TIPO_CARGO"         ,         typeof(string)    );//Cargo Fijo
                dt.Columns.Add("Codigo de Cargo"    ,         typeof(string)    );//CCRCSM_HBO
                dt.Columns.Add("DESCRIPTION"        ,         typeof(string)    );//HBO Mensual
                dt.Columns.Add("AJUSTE"             ,         typeof(double)    );//MONTO TOTAL

                foreach (var item in hbos_resultado)
                {
                    if (item.MONTO_TOTAL == 0)
                    {

                    }
                    else
                    {
                        dt.Rows.Add(new Object[] 
                        {
                            long.Parse(item.CuentaFinanciera)   ,
                            "Control HBO"                       ,
                            long.Parse(item.PCS)                ,
                            "Cargo Fijo"                        ,
                            "CCRCSM_HBO"                        ,
                            "HBO Mensual"                       ,
                            item.MONTO_TOTAL 
                        });
                    }

                }
                //oSLDocument.ImportDataTable(1, 1, dt, true);
                //oSLDocument.RenameWorksheet("Sheet1", "AJUSTES HBO");
                //oSLDocument.SaveAs(@salida + @"\Ciclo" + @"\Ruta\" + @"\HBO.xlsx"); //
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

            // A____A
            //|・ㅅ・|
            //|っ ｃ |
            //|      |
            //|      |
            //|      |
            //|      |
            //|      |
            // U￣U￣
        } //esta listo
        public bool paramount_sum() //igual a simcard, no olvidar los downgrade y mandarlo, dsp ponerme a revisar empresas ya ya ya
        {
            try
            {
                var query = paramounts
                    .GroupBy(d => d.PCS)
                      .Select(
                              g => new
                              {
                                  Key                 =       g.Key,
                                  Monto_Total         =       g.Sum(s => s.MONTO_TOTAL),
                                  CuentaFinanciera    =       g.First().CuentaFinanciera,
                                  Pcs                 =       g.First().PCS,
                                  CodigoCarga         =       g.First().Codigo_de_Cargo
                              });

                var dt = new DataTable("PARAMOUNT");

                dt.Columns.Add("Cuenta Financiera" ,        typeof(long));
                dt.Columns.Add("MOTIVO"            ,        typeof(string));//Control Paramount
                dt.Columns.Add("OBSERVACIÓN"       ,        typeof(long));//PCS
                dt.Columns.Add("TIPO_CARGO"        ,        typeof(string));//Cargo Fijo
                dt.Columns.Add("Codigo de Cargo"   ,        typeof(string));//CCRCSM_PARAMOUN
                dt.Columns.Add("DESCRIPTION"       ,        typeof(string));//PARAMOUNT Mensual
                dt.Columns.Add("AJUSTE"            ,        typeof(double));//MONTO TOTAL


                foreach (var item in query)
                {
                    if (item.Monto_Total > 0.0)
                    {
                        double monto    =   item.Monto_Total * -1;
                        monto           =   Math.Round(monto);
                        if (monto == 0.0)
                        {
                            continue;
                        }
                        else
                        {
//-------------------------------------------------------------------------------------------------------
                            dt.Rows.Add(new Object[] 
                            {
                                long.Parse(item.CuentaFinanciera)        ,
                                "Control Paramount"                      ,
                                long.Parse(item.Pcs)                     ,
                                "Cargo Fijo"                             ,
                                "CCRCSM_PARAMOUN"                        ,
                                "PARAMOUNT Mensual"                      ,
                                Math.Round(monto)
                            });
//-------------------------------------------------------------------------------------------------------
                        }

                    }

                }

                //oSLDocument2.ImportDataTable(1, 1, dt, true);

                //oSLDocument2.RenameWorksheet("Sheet1", "AJUSTES PARAMOUNT");

                //oSLDocument2.SaveAs(@salida +  @"\Ciclo" + @"\Ruta\" + @"\PARAMOUNT.xlsx");//@"C:\Users\Marcelo\Desktop\" + @"\AJUSTES DONWGRADE _CICLO_X_X.xlsx"
                
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }
    }
}
