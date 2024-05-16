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
    public class Controlador_Gyp
    {
        List<CargosxPcs> cargosxPcs = new List<CargosxPcs>();
        List<Gyp_Pro> gyp_Pros = new List<Gyp_Pro>();
        Coleccion_Gyp_Pro coleccion_Gyp = new Coleccion_Gyp_Pro();
        List<CargosxPcs> cargosxPcs_cargos = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPcs_cargos_gyp = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPcs_sin_cero = new List<CargosxPcs>();
        List<CargosxPcs> cargosxPcs_duplicados_sumados = new List<CargosxPcs>();
        SLDocument oSLDocument = new SLDocument();
        List<string> tipos = new List<string>();
        string aux_path = string.Empty;

        public Controlador_Gyp(string path_gyp, List<CargosxPcs> cargosxes,List<string> _tipos,string path_salida)
        {
            cargosxPcs = cargosxes;
            gyp_Pros = coleccion_Gyp.GenerarListado(@path_gyp);
            aux_path = path_salida;
            tipos = _tipos;
        }

        public bool procesando_cargosx_pc_tipo()
        {
            try 
            {
                foreach (var item1 in cargosxPcs)
                {
                    foreach (var item2 in tipos)
                    {
                        if (item1.codigodecargo.Equals(item2))
                        {
                            cargosxPcs_cargos.Add(item1); //hasta aqui esta bien
                        }
                    }

                }

                //var query = cargosxPcs2.GroupBy(d => d.PCS)
                //                .Select(
                //                        g => new
                //                        {
                //                            Key = g.Key,
                //                            Monto_Total = g.Sum(s => s.montoTotal),
                //                            CuentaFinanciera = g.First().CuentaFinanciera,
                //                            Pcs = g.First().PCS,
                //                            CodigoCarga = g.First().codigodecargo
                //                        });

                //List<CargosxPcs> cargosxPcs_duplicados = new List<CargosxPcs>();
                //worker.ReportProgress(i);
                //i = 40;
                //worker.ReportProgress(i);
                //i = 50;

                // DirectoryInfo di5 = Directory.CreateDirectory(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\Gyp Pro");
                // SaveExcel.BuildExcel(miTable2, @seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
                //var dt = new DataTable("CCUCTEL_ADIC");
                //dt.Columns.Add("CUENTA", typeof(long));
                //dt.Columns.Add("PCS", typeof(long));
                //dt.Columns.Add("CARGO", typeof(string));
                //dt.Columns.Add("AMOUNT", typeof(double));
                //dt.Columns.Add("CICLO", typeof(string));
                //foreach (var item in cargosxPcs_duplicados)
                //{
                //    item.montoTotal = item.montoTotal * -1;
                //    dt.Rows.Add(new Object[] {long.Parse(item.CuentaFinanciera),
                //                               long.Parse(item.PCS),
                //                               item.codigodecargo,
                //                               item.montoTotal, ""});

                //}
                //oSLDocument.ImportDataTable(1, 1, dt, true);
                //oSLDocument.RenameWorksheet("Sheet1", "CCUCTEL_ADIC");
                //oSLDocument.SaveAs(@seleccionar_carpeta_txt.Text + @"\Ciclo" + @"\Ruta\" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");
                return true;
            } catch(Exception ex) 
            {
                return false;
            }
        }
        public bool procesando_cargosx_pc_x_gyp()
        {
            try 
            {


                var pcs_linq_gyp = from arrendamiento in cargosxPcs_cargos
                                   join sice in gyp_Pros
                                                on arrendamiento.CuentaFinanciera equals sice.cuenta
                                           select new CargosxPcs
                                           {
                                               CuentaFinanciera = arrendamiento.CuentaFinanciera,
                                               recveivercustomer = arrendamiento.recveivercustomer,
                                               tipocargo = arrendamiento.tipocargo,
                                               codigodecargo = arrendamiento.codigodecargo,
                                               offwer = arrendamiento.offwer,
                                               nombreOFFER = arrendamiento.nombreOFFER,
                                               promo = arrendamiento.promo,
                                               description = arrendamiento.description,
                                               montoCargos = arrendamiento.montoCargos,
                                               montoDescuentos = arrendamiento.montoDescuentos,
                                               montoTotal = arrendamiento.montoTotal,
                                               tipodeCobro = arrendamiento.tipodeCobro,
                                               PCS = arrendamiento.PCS,
                                               secuenciadeCiclo = arrendamiento.secuenciadeCiclo,
                                               prorrateo = arrendamiento.prorrateo,
                                               rut = arrendamiento.rut
                                           };

                cargosxPcs_cargos_gyp = pcs_linq_gyp.ToList();
                
                return true;
            
            }
            catch (Exception ex) 
            {
                return false;
            }
        }
        public bool procesando_cargosx_pc_sin_0() 
        {
            try 
            {
                foreach (var item in cargosxPcs_cargos_gyp)
                {
                    if (item.montoTotal == 0)
                    {
                        continue;
                    }
                    else
                    {
                        cargosxPcs_sin_cero.Add(item);
                    }
                }
                
                return true;
            }catch(Exception ex) 
            {
                return false;
            }
        }
        public bool procesando_cargosx_pc_duplicados()
        {
            try 
            {
                List<CargosxPcs> cargosxpcs_duplicados = new List<CargosxPcs>();

                var final_results = cargosxPcs_sin_cero.GroupBy(n => new { n.PCS, n.codigodecargo})
                       .Select(g => g.FirstOrDefault()).ToList();

                foreach (var item1 in final_results)
                {
                    CargosxPcs cargosx = new CargosxPcs();
                    foreach (var item2 in cargosxPcs_sin_cero)
                    {
                        if (item1.PCS.Equals(item2.PCS)&&item1.codigodecargo.Equals(item2.codigodecargo))
                        {
                            cargosx.PCS = item2.PCS;
                            cargosx.codigodecargo = item2.codigodecargo;
                            cargosx.montoTotal = cargosx.montoTotal + item2.montoTotal;
                            cargosx.CuentaFinanciera = item2.CuentaFinanciera;
                        }
                       
                    }
                    cargosxpcs_duplicados.Add(cargosx);
                }

                var dt = new DataTable("Gyp");
                dt.Columns.Add("CUENTA", typeof(long));
                dt.Columns.Add("PCS", typeof(long));
                dt.Columns.Add("CARGO", typeof(string));
                dt.Columns.Add("AMOUNT", typeof(double));
                dt.Columns.Add("CICLO", typeof(string));
                foreach (var item in cargosxpcs_duplicados)
                {
                    item.montoTotal = item.montoTotal * -1;
                    if (Math.Round(item.montoTotal)==0)
                    {
                        continue;
                    }
                    else
                    {
                        dt.Rows.Add(new Object[] {long.Parse(item.CuentaFinanciera),
                                               long.Parse(item.PCS),
                                               item.codigodecargo,
                                               Math.Round(item.montoTotal), "XX"});
                    }


                }
                oSLDocument.ImportDataTable(1, 1, dt, true);
                oSLDocument.RenameWorksheet("Sheet1", "GYP");
                oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\Ruta\" + @"\CICLO_GYP_PRO_XX_MES_ANIO.xlsx");


                string hola = string.Empty;
                return true;
            }catch(Exception ex) 
            {
                return false;
            }
        }
    }
}
