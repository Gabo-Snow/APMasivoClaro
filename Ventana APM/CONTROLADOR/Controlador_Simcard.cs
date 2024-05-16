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
   public class Controlador_Simcard
    {
        List<SimCard> simcards = new List<SimCard>();
        Coleccion_SimCard coleccion_Sim = new Coleccion_SimCard();
        SLDocument oSLDocument = new SLDocument();
        string aux_path = string.Empty;

        public Controlador_Simcard(string path_simcard,string path_salida)
        {
            simcards = coleccion_Sim.GenerarListado(@path_simcard);
            aux_path = path_salida;
        }

        public bool procesando_simcard()
        {
            try
            {
                var query = simcards.GroupBy(d => d.PCS)
                                .Select(
                                        g => new
                                        {
                                            Key = g.Key,
                                            Monto_Total = g.Sum(s => s.montoTotal),
                                            CuentaFinanciera = g.First().CuentaFinanciera,
                                            Pcs = g.First().PCS,
                                            CodigoCarga = g.First().codigodecargo
                                        });
                var dt = new DataTable("SIMCARD");
                dt.Columns.Add("PCS", typeof(long));
                dt.Columns.Add("CuentaFinanciera", typeof(long)); 
                dt.Columns.Add("Codigo de Cargo", typeof(string));
                dt.Columns.Add("AMOUNT", typeof(double));//quitar decimal
                dt.Columns.Add("CICLO", typeof(string));
                foreach (var item in query)
                {
                    if (item.Monto_Total > 0.0)
                    {
                        double monto = item.Monto_Total * -1;
                        monto = Math.Round(monto);
                        if (monto == 0.0)
                        {
                            continue;
                        }
                        else
                        {
                            
                            dt.Rows.Add(new Object[] {long.Parse(item.Pcs),
                                               long.Parse(item.CuentaFinanciera),
                                               item.CodigoCarga,
                                               Math.Round(monto),
                                               "DIA_MES"});
                        }

                    }

                }

                //oSLDocument.ImportDataTable(1, 1, dt, true);
                //oSLDocument.RenameWorksheet("Sheet1", "AJUSTES SIMCARD");
                //oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\Ruta\" + @"\SIMCARD CICLO_MES_ANIO.xlsx");//@"C:\Users\Marcelo\Desktop\" + @"\AJUSTES DONWGRADE _CICLO_X_X.xlsx"
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }


        ////i = 80;
        ////    worker.ReportProgress(i);
        ////    var miTable24 = new DataTable("SIMCARD");
        ////miTable24.Columns.Add("PCS");
        ////    miTable24.Columns.Add("CUENTA FINANCIERA");
        ////    miTable24.Columns.Add("CODIGO DE CARGO");
        ////    miTable24.Columns.Add("AMOUNT");
        ////    miTable24.Columns.Add("CICLO");


        ////    foreach (var item in query4)
        ////    {
        ////        if (item.Monto_Total > 0.0)
        ////        {
        ////            miTable24.Rows.Add(new Object[] { item.Pcs, item.CuentaFinanciera, item.CodigoCarga, "-" + item.Monto_Total, "" });
        ////        }
        ////    }

        ////    SaveExcel.BuildExcel(miTable24, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Ruta" + @"\SIMCARD CICLO_MES_ANIO.xlsx");
    }
}
