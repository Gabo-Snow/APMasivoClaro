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
    public class Controlador_CFM
    {
        //string filename = @"C:\Users\Marcelo\Desktop\CFM\CARGOS X PCS JUNTOS.txt";
        List<string> filtro_descripcion = new List<string>() { "Telefonia",
                                                               "Impresion Detalle de Llamadas",
                                                               "BAM Controlada",
                                                               "Proteccion Movil CAT 2",
                                                               "Plan base promocion 2x1",
                                                               "Numero Privado",
                                                               "Plan base promocion lineas adicionales",
                                                               "Proteccion Movil CAT 4",
                                                               "Promocion CE",
                                                               "Proteccion Movil CAT 3",
                                                               "Telefonia Controlada",
                                                               "HBO Mensual",
                                                               "Proteccion Movil CAT 1",
                                                               "Proteccion Movil CAT 5",
                                                               "Suscripcion Mensual Claro Musica",
                                                               "Claro Video Mensual",
                                                               "PARAMOUNT Mensual",
                                                               "TV en vivo nacional",
                                                               "A3PLAYER Mensual",
                                                               "Internet",
                                                               "CDF HD",
                                                               "Regularizacion de Cargo Fijo"};
        List<string> filtro_descripcion_2 = new List<string>();
        ///--------------------------------------Colecciones-------------------------------
        Coleccion_CFM coleccion_CFM = new Coleccion_CFM();
        Coleccion_Fide_Cob_Pcs coleccion_Fide_Cob_Pcs = new Coleccion_Fide_Cob_Pcs();
        //--------------------------------------------------------------------------------
        //--------------------------------------Listas-----------------------------------
        List<Fide_Cob_Pcs> Cargosx_Pcs1 = new List<Fide_Cob_Pcs>();
        List<CFM_AJUSTES> cFMs = new List<CFM_AJUSTES>();
        List<Fide_Cob_Pcs> pcs_cargo_fijo = new List<Fide_Cob_Pcs>();
        List<Aux_Pcs_CC> aux_Pcs_CCs = new List<Aux_Pcs_CC>();
        List<Fide_Cob_Pcs> pcs_cargo_pcs_aux_descripcion = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> fide_Cob_Pcs_list = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> fide_Spynner_Pcs_list = new List<Fide_Cob_Pcs>();
        List<Fide_Cob_Pcs> fide_Cob_Pcs_filtrado = new List<Fide_Cob_Pcs>();

        List<CFM_AJUSTES> cFMs_CRUZA = new List<CFM_AJUSTES>();

        string aux_path = string.Empty;
        string cargo_fij = string.Empty;
          //-----------------------CREADOR DEL EXCEL---------------------------------------
        SLDocument oSLDocument = new SLDocument();
        SLDocument oSLDocument2 = new SLDocument();
        SLDocument oSLDocument3 = new SLDocument();
        public Controlador_CFM(List<Fide_Cob_Pcs> cargosxpcs, string path_cfm, string path_salida, List<string> filtro, string es_cargo_fijo)
        {
            aux_path = path_salida;
            Cargosx_Pcs1 = cargosxpcs;

            //agregar las carteras
            cFMs = coleccion_CFM.GenerarListado(@path_cfm);//GenerarListado(@"C:\Users\Marcelo\Desktop\CFM\AJUSTES CFM C10.03.2021.xlsx");
            filtro_descripcion_2 = filtro;
            cargo_fij = es_cargo_fijo;
        }

        public bool proceso_cfm_filtrando()
        {
            //telefonia
            //bam 
            //internet
            //telefonia controlada
            try
            {
                foreach (var item1 in Cargosx_Pcs1)
                {
                    if (item1.tipocargo.Equals(cargo_fij))
                    {
                        pcs_cargo_fijo.Add(item1);
                    }
                }

                foreach (var item1 in pcs_cargo_fijo)
                {
                    foreach (var item2 in filtro_descripcion_2)
                    {
                        if (item1.description.Equals(item2))
                        {
                            fide_Cob_Pcs_filtrado.Add(item1);
                        }
                    }
                }
                return true;
            }catch(Exception ex) 
            {
                return false;
            }
        }

        public bool proceso_cfm_duplicados()
        {
            try
            {
                List<CargosxPcs> cargosxpcs_duplicados = new List<CargosxPcs>();
                var QS = (from std in fide_Cob_Pcs_filtrado
                          select std)
                            .Select(std => new { std.PCS, std.CuentaFinanciera })//ESTO TA BIEN AHORA HAY QUE CONTAR
                            .Distinct().ToList();

                //var final_results = fide_Cob_Pcs_filtrado.GroupBy(n => new { n.PCS, n.CuentaFinanciera })
                //       .Select(g => g.FirstOrDefault()).ToList();
                var result = QS.
                                GroupBy(x => x.PCS).
                                Select(x => new
                                {
                                    cuentas = x.Select(v => v.CuentaFinanciera).Distinct().Count(),
                                    pcs = x.Key
                                });


                var dt2 = new DataTable("CON DOS CUENTAS");
                dt2.Columns.Add("PCS", typeof(long));
                dt2.Columns.Add("CUENTA FINANCIERA", typeof(int));



                //List<CFM_AJUSTES> cFMs_aux = new List<CFM_AJUSTES>();
                foreach (var item in result)
                {
                    if (item.pcs.Equals(""))
                    {
                        continue;
                    }
                    else
                    {
                        if (item.cuentas >= 2)
                        {
                            dt2.Rows.Add(new Object[] { long.Parse(item.pcs), item.cuentas });
                        }
                    }


                }
                oSLDocument.AddWorksheet("CARGO CON DOS CUENTAS");
                oSLDocument.ImportDataTable(1, 1, dt2, true);
                //oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\RESPALDO CFM.xlsx");
                return true;
            }catch(Exception ex)
            {
                return false;
            }
        }

        public bool proceso_cfm_sumados()
        {
            try
            {
                var query = fide_Cob_Pcs_filtrado.GroupBy(d => d.PCS)
                                        .Select(
                                                g => new
                                                {
                                                    Key = g.Key,
                                                    Monto_Total = g.Sum(s => s.montoTotal)
                                                });

                //var dt3 = new DataTable("SUM");
                //dt3.Columns.Add("PCS", typeof(long));
                //dt3.Columns.Add("MONTO TOTAL", typeof(double));
                //foreach (var item in query)
                //{
                //    if (item.Key.Equals(""))
                //    {
                //        continue;
                //    }
                //    else
                //    {
                //        dt3.Rows.Add(new Object[] { long.Parse(item.Key), item.Monto_Total });
                //    }
                //}
                //oSLDocument.AddWorksheet("SUM");
                //oSLDocument.ImportDataTable(1, 1, dt3, true);

                //var dt4 = new DataTable("CRUZA CON CARGOS");
                //dt4.Columns.Add("PCS", typeof(long));
                //dt4.Columns.Add("CUENTA", typeof(double));
                //dt4.Columns.Add("MONTO", typeof(double));
                //dt4.Columns.Add("CF MAS IVA", typeof(double));
                //dt4.Columns.Add("CARGOS", typeof(double));
                //dt4.Columns.Add("AJUSTE", typeof(double));
                //dt4.Columns.Add("SOLICITANTE", typeof(string));

                foreach (var item in cFMs)
                {
                    item.CF_MAS_IVA = item.MONTO / 1.19;
                }

                var quer_cf = from cf in cFMs
                                           join quer in query
                                                on cf.PCS equals quer.Key
                                           select new CFM_AJUSTES
                                           {
                                               PCS = cf.PCS,
                                               CUENTA = cf.CUENTA,
                                               CICLO = cf.CICLO,
                                               MONTO = cf.MONTO,
                                               CF_MAS_IVA = cf.CF_MAS_IVA,
                                               CARGOS = quer.Monto_Total,
                                               AJUSTE = cf.AJUSTE,
                                               APLICA_CDP = cf.APLICA_CDP,
                                               CAMPANA = cf.CAMPANA
                                            };

                List<CFM_AJUSTES> auxiliar_cfm_1 = quer_cf.ToList();
                List<CFM_AJUSTES> auxiliar_aplicados = new List<CFM_AJUSTES>();

                HashSet<string> pcs_cruce_cfm = new HashSet<string>(auxiliar_cfm_1.Select(x => x.PCS));

                cFMs.RemoveAll(x => pcs_cruce_cfm.Contains(x.PCS));
                foreach (var item in auxiliar_cfm_1)
                {
                    item.AJUSTE = item.CF_MAS_IVA - item.CARGOS;
                    if (item.AJUSTE >= -99)
                    {
                        auxiliar_aplicados.Add(item);
                    }
                }

                HashSet<string> pcs_aplicados_cfm = new HashSet<string>(auxiliar_aplicados.Select(x => x.PCS));

                auxiliar_cfm_1.RemoveAll(x => pcs_aplicados_cfm.Contains(x.PCS));
                var dt5 = new DataTable("APLICADOS");
                dt5.Columns.Add("PCS", typeof(long));
                dt5.Columns.Add("CUENTA", typeof(double));
                dt5.Columns.Add("MONTO", typeof(double));
                dt5.Columns.Add("CF MAS IVA", typeof(double));
                dt5.Columns.Add("CARGOS", typeof(double));
                dt5.Columns.Add("AJUSTE", typeof(double));
                dt5.Columns.Add("SOLICITANTE", typeof(string));
                foreach (var item in auxiliar_aplicados)
                {
                    dt5.Rows.Add(new Object[] { long.Parse(item.PCS),long.Parse(item.CUENTA), item.MONTO,item.CF_MAS_IVA,item.CARGOS,item.AJUSTE,item.CAMPANA }); //correcto
                }
                oSLDocument.AddWorksheet("APLICADOS");
                oSLDocument.ImportDataTable(1, 1, dt5, true);

                var dt6 = new DataTable("SIN CARGOS");
                dt6.Columns.Add("PCS", typeof(long));
                dt6.Columns.Add("CUENTA", typeof(double));
                dt6.Columns.Add("MONTO", typeof(double));
                dt6.Columns.Add("CF MAS IVA", typeof(double));
                dt6.Columns.Add("CARGOS", typeof(double));
                dt6.Columns.Add("AJUSTE", typeof(double));
                dt6.Columns.Add("SOLICITANTE", typeof(string));
                foreach (var item in cFMs)
                {
                    dt6.Rows.Add(new Object[] { long.Parse(item.PCS), long.Parse(item.CUENTA), item.MONTO, item.CF_MAS_IVA, item.CARGOS, item.AJUSTE, item.CAMPANA }); //correcto
                }
                oSLDocument.AddWorksheet("SIN CARGOS");
                oSLDocument.ImportDataTable(1, 1, dt6, true);

                var dt7 = new DataTable("PLANILLA");
                dt7.Columns.Add("PCS", typeof(long));
                dt7.Columns.Add("CUENTA", typeof(double));
                dt7.Columns.Add("MONTO", typeof(double));
                dt7.Columns.Add("CF MAS IVA", typeof(double));
                dt7.Columns.Add("CARGOS", typeof(double));
                dt7.Columns.Add("AJUSTE", typeof(double));
                dt7.Columns.Add("SOLICITANTE", typeof(string));
                foreach (var item in auxiliar_cfm_1)
                {
                    dt7.Rows.Add(new Object[] { long.Parse(item.PCS), long.Parse(item.CUENTA), item.MONTO, item.CF_MAS_IVA, item.CARGOS, item.AJUSTE, item.CAMPANA }); //correcto
                }

                //oSLDocument.ImportDataTable(1, 1, dt7, true);
                //oSLDocument.RenameWorksheet("Sheet1", "PLANILLA CFM");
                oSLDocument.AddWorksheet("PLANILLA CFM");
                oSLDocument.ImportDataTable(1, 1, dt7, true);
                oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\RESPALDO CFM.xlsx");

                            

                var dt8 = new DataTable("PLANILLA");
                dt8.Columns.Add("CUENTA", typeof(long));
                dt8.Columns.Add("PCS", typeof(double));
                dt8.Columns.Add("MONTO FACTURADO", typeof(double));//CARGOS
                dt8.Columns.Add("MONTO SOLICITADO", typeof(double));//CF MAS IVA
                dt8.Columns.Add("AJUSTES", typeof(double));// AJUSTE
                dt8.Columns.Add("SOLICITANTE", typeof(string));//CAMPAÑA
                foreach (var item in auxiliar_cfm_1)
                {
                    dt8.Rows.Add(new Object[] { long.Parse(item.CUENTA), long.Parse(item.PCS), Math.Round(item.CARGOS),Math.Round(item.CF_MAS_IVA), Math.Round(item.AJUSTE), item.CAMPANA });
                }

                oSLDocument2.ImportDataTable(1, 1, dt8, true);
                oSLDocument2.RenameWorksheet("Sheet1", "AJUSTES CFM");
                oSLDocument2.SaveAs(@aux_path + @"\Ciclo" + @"\Final\" + @"\AJUSTES CFM CICLO XX LYON.xlsx");
                return true;

            }catch(Exception ex)
            {
                return false;
            }

        }

        public bool proceso_cfm()
        {
            try
            {
                //-------------------------ALGORITMOS------------------------------------


                //columnas




                foreach (var item1 in Cargosx_Pcs1)
                {
                    if (item1.tipocargo.Equals(cargo_fij))
                    {

                        //dt1.Rows.Add(new Object[] { long.Parse(item1.CuentaFinanciera)
                        //,item1.recveivercustomer
                        //,item1.tipocargo
                        //,item1.codigodecargo
                        //,item1.offwer
                        //,item1.nombreOFFER
                        //,item1.promo
                        //,item1.description
                        //,item1.montoCargos
                        //,item1.montoDescuentos
                        //,item1.montoTotal
                        //,item1.tipodeCobro
                        //,item1.PCS
                        //,item1.secuenciadeCiclo
                        //,item1.prorrateo});

                        pcs_cargo_fijo.Add(item1);
                        //dt.Rows.Add(new Object[] { long.Parse(item1.CuentaFinanciera)
                        //,item1.recveivercustomer
                        //,item1.tipocargo
                        //,item1.codigodecargo
                        //,item1.offwer
                        //,item1.nombreOFFER
                        //,item1.promo
                        //,item1.description
                        //,item1.montoCargos
                        //,item1.montoDescuentos
                        //,item1.montoTotal
                        //,item1.tipodeCobro
                        //,long.Parse(item1.PCS)
                        //,item1.secuenciadeCiclo
                        //,item1.prorrateo});

                    }
                }



                HashSet<string> diffids = new HashSet<string>(cFMs.Select(s => s.PCS));
                IEnumerable<Fide_Cob_Pcs> bs =
                                    from a in diffids
                                    join b in pcs_cargo_fijo on a equals b.PCS
                                    select b;
                List<Fide_Cob_Pcs> list_pcs_cfm = bs.ToList();

                
                foreach (var item in list_pcs_cfm)
                {
                    Aux_Pcs_CC aux_Pcs_CC = new Aux_Pcs_CC();
                    aux_Pcs_CC.CuentaFinanciera = item.CuentaFinanciera;
                    aux_Pcs_CC.PCS = item.PCS;
                    aux_Pcs_CCs.Add(aux_Pcs_CC);
                }

                var QS = (from std in aux_Pcs_CCs
                          select std)
              .Select(std => new { std.PCS, std.CuentaFinanciera })//ESTO TA BIEN AHORA HAY QUE CONTAR
              .Distinct().ToList();

                QS.Count();

                var result = QS.
                    GroupBy(x => x.PCS).
                    Select(x => new
                    {
                        cuentas = x.Select(v => v.CuentaFinanciera).Distinct().Count(),
                        pcs = x.Key
                    });
                System.Data.DataTable dt = new System.Data.DataTable("CFM CARGO FIJO FILTRADO");
                dt.TableName = "Cargo Fijo";
                dt.Columns.Add("CuentaFinanciera", typeof(long));
                dt.Columns.Add("recveivercustomer", typeof(string));
                dt.Columns.Add("tipocargo", typeof(string));
                dt.Columns.Add("codigodecargo", typeof(string));
                dt.Columns.Add("offwer", typeof(string));
                dt.Columns.Add("nombreOFFER", typeof(string));
                dt.Columns.Add("promo", typeof(string));
                dt.Columns.Add("description", typeof(string));
                dt.Columns.Add("montoCargos", typeof(double));
                dt.Columns.Add("montoDescuentos", typeof(double));
                dt.Columns.Add("montoTotal", typeof(double));
                dt.Columns.Add("tipodeCobro", typeof(string));
                dt.Columns.Add("PCS", typeof(long));
                dt.Columns.Add("secuenciadeCiclo", typeof(string));
                dt.Columns.Add("prorrateo", typeof(string));

                System.Data.DataTable dt1 = new System.Data.DataTable("CFM CARGO FIJO SIN PCS");
                dt1.TableName = "Cargo Fijo";
                dt1.Columns.Add("CuentaFinanciera", typeof(long));
                dt1.Columns.Add("recveivercustomer", typeof(string));
                dt1.Columns.Add("tipocargo", typeof(string));
                dt1.Columns.Add("codigodecargo", typeof(string));
                dt1.Columns.Add("offwer", typeof(string));
                dt1.Columns.Add("nombreOFFER", typeof(string));
                dt1.Columns.Add("promo", typeof(string));
                dt1.Columns.Add("description", typeof(string));
                dt1.Columns.Add("montoCargos", typeof(double));
                dt1.Columns.Add("montoDescuentos", typeof(double));
                dt1.Columns.Add("montoTotal", typeof(double));
                dt1.Columns.Add("tipodeCobro", typeof(string));
                dt1.Columns.Add("PCS", typeof(string));
                dt1.Columns.Add("secuenciadeCiclo", typeof(string));
                dt1.Columns.Add("prorrateo", typeof(string));
                foreach (var item1 in pcs_cargo_fijo)
                {
                    foreach (var item2 in filtro_descripcion_2)
                    {
                        if (item1.description.Equals(item2))
                        {
                            if (item1.PCS.Equals(""))
                            {
                                dt1.Rows.Add(new Object[] { long.Parse(item1.CuentaFinanciera)
                                  ,item1.recveivercustomer
                                  ,item1.tipocargo
                                  ,item1.codigodecargo
                                  ,item1.offwer
                                  ,item1.nombreOFFER
                                  ,item1.promo
                                  ,item1.description
                                  ,item1.montoCargos
                                  ,item1.montoDescuentos
                                  ,item1.montoTotal
                                  ,item1.tipodeCobro
                                  ,item1.PCS
                                  ,item1.secuenciadeCiclo
                                  ,item1.prorrateo});
                            }
                            else
                            {
                                pcs_cargo_pcs_aux_descripcion.Add(item1);
                                //dt.Rows.Add(new Object[] { long.Parse(item1.CuentaFinanciera)
                                //                            ,item1.recveivercustomer
                                //                            ,item1.tipocargo
                                //                            ,item1.codigodecargo
                                //                            ,item1.offwer
                                //                            ,item1.nombreOFFER
                                //                            ,item1.promo
                                //                            ,item1.description
                                //                            ,item1.montoCargos
                                //                            ,item1.montoDescuentos
                                //                            ,item1.montoTotal
                                //                            ,item1.tipodeCobro
                                //                            ,long.Parse(item1.PCS)
                                //                            ,item1.secuenciadeCiclo
                                //                            ,item1.prorrateo});
                            }

                            break;
                        }
                    }
                }

                if (pcs_cargo_pcs_aux_descripcion.Count > 400000)
                {

                    System.Data.DataTable dtaux1 = new System.Data.DataTable("CFM CARGO FIJO FILTRADO 2");
                    dtaux1.TableName = "Cargo Fijo";
                    dtaux1.Columns.Add("CuentaFinanciera", typeof(long));
                    dtaux1.Columns.Add("recveivercustomer", typeof(long));
                    dtaux1.Columns.Add("tipocargo", typeof(string));
                    dtaux1.Columns.Add("codigodecargo", typeof(string));
                    dtaux1.Columns.Add("offwer", typeof(long));
                    dtaux1.Columns.Add("nombreOFFER", typeof(string));
                    dtaux1.Columns.Add("promo", typeof(string));
                    dtaux1.Columns.Add("description", typeof(string));
                    dtaux1.Columns.Add("montoCargos", typeof(double));
                    dtaux1.Columns.Add("montoDescuentos", typeof(double));
                    dtaux1.Columns.Add("montoTotal", typeof(double));
                    dtaux1.Columns.Add("tipodeCobro", typeof(string));
                    dtaux1.Columns.Add("PCS", typeof(long));
                    dtaux1.Columns.Add("secuenciadeCiclo", typeof(long));
                    dtaux1.Columns.Add("prorrateo", typeof(string));
                    int contador_hojas = 0;
                    foreach (var item in pcs_cargo_pcs_aux_descripcion)
                    {
                        contador_hojas++;
                        if (contador_hojas > 440000)
                        {
                            dtaux1.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera)
                                                          ,long.Parse(item.recveivercustomer)
                                                          ,item.tipocargo
                                                          ,item.codigodecargo
                                                          ,long.Parse(item.offwer)
                                                          ,item.nombreOFFER
                                                          ,item.promo
                                                          ,item.description
                                                          ,item.montoCargos
                                                          ,item.montoDescuentos
                                                          ,item.montoTotal
                                                          ,item.tipodeCobro
                                                          ,long.Parse(item.PCS)
                                                          ,long.Parse(item.secuenciadeCiclo)
                                                          ,item.prorrateo});
                        }
                        else
                        {
                            dt.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera)
                                                          ,long.Parse(item.recveivercustomer)
                                                          ,item.tipocargo
                                                          ,item.codigodecargo
                                                          ,long.Parse(item.offwer)
                                                          ,item.nombreOFFER
                                                          ,item.promo
                                                          ,item.description
                                                          ,item.montoCargos
                                                          ,item.montoDescuentos
                                                          ,item.montoTotal
                                                          ,item.tipodeCobro
                                                          ,long.Parse(item.PCS)
                                                          ,long.Parse(item.secuenciadeCiclo)
                                                          ,item.prorrateo});
                        }

                    }
                    SLDocument oSLDocument2 = new SLDocument();
                    oSLDocument2.ImportDataTable(1, 1, dt, true);
                    oSLDocument2.RenameWorksheet("Sheet1", "CARGO FIJO FILTRADO 2");
                    oSLDocument2.SaveAs(@aux_path + @"\Ciclo"+ @"\RESPALDO CFM_2.xlsx");

                    oSLDocument.ImportDataTable(1, 1, dtaux1, true);
                    oSLDocument.RenameWorksheet("Sheet1", "CARGO FIJO FILTRADO 1");

                }
                else
                {
                    oSLDocument.ImportDataTable(1, 1, dt, true);
                    oSLDocument.RenameWorksheet("Sheet1", "CARGO FIJO FILTRADO");
                }


                oSLDocument.AddWorksheet("CARGO SIN PCS");
                oSLDocument.ImportDataTable(1, 1, dt1, true);



                var dt2 = new DataTable("CON DOS CUENTAS");
                dt2.Columns.Add("PCS", typeof(long));
                dt2.Columns.Add("CUENTA FINANCIERA", typeof(int));

                foreach (var item in result)
                {
                    if (item.cuentas >= 2)
                    {
                        dt2.Rows.Add(new Object[] { long.Parse(item.pcs), item.cuentas });
                    }

                }
                oSLDocument.AddWorksheet("CARGO CON DOS CUENTAS");
                oSLDocument.ImportDataTable(1, 1, dt2, true);
                //corregir POR LAS TILDES INSERTAR UTF


                var suma_pcs_cargos = pcs_cargo_pcs_aux_descripcion.GroupBy(d => d.PCS)
                                        .Select(
                                                g => new
                                                {
                                                    Key = g.Key,
                                                    Monto_Total = g.Sum(s => s.montoTotal),
                                                    Pcs = g.First().PCS
                                                });


                var SUMADOS_CFM = from suma in suma_pcs_cargos
                                  join ajust in cFMs
                                       on suma.Pcs equals ajust.PCS
                                  select new
                                  {
                                      suma.Pcs,
                                      suma.Monto_Total,
                                      ajust.MONTO,
                                      ajust.CUENTA,
                                      ajust.CAMPANA
                                  };

                // SLDocument oSLDocument4 = new SLDocument();
                var dt3 = new DataTable("SUMA");
                dt3.Columns.Add("PCS", typeof(long));
                dt3.Columns.Add("MONTO FACTURADO", typeof(double));
                dt3.Columns.Add("MONTO AJUSTADO", typeof(double));
                dt3.Columns.Add("CUENTA FINANCIERA", typeof(long));


                foreach (var item in SUMADOS_CFM)
                {
                    dt3.Rows.Add(new Object[] { long.Parse(item.Pcs), item.Monto_Total, item.MONTO, long.Parse(item.CUENTA) });
                }
                oSLDocument.AddWorksheet("SUMA");
                oSLDocument.ImportDataTable(1, 1, dt3, true);




                var dt4 = new DataTable("-99");
                dt4.Columns.Add("PCS", typeof(long));
                dt4.Columns.Add("MONTO", typeof(double));
                dt4.Columns.Add("SOLICITADO", typeof(double));
                dt4.Columns.Add("MONTO SIN IVA", typeof(double));
                dt4.Columns.Add("MONTO AJUSTADO IVA", typeof(double));
                dt4.Columns.Add("CUENTA FINANCIERA", typeof(long));


                var dt5 = new DataTable("+99");
                dt5.Columns.Add("PCS", typeof(long));
                dt5.Columns.Add("MONTO", typeof(double));
                dt5.Columns.Add("SOLICITADO", typeof(double));
                dt5.Columns.Add("MONTO SIN IVA", typeof(double));
                dt5.Columns.Add("MONTO AJUSTADO IVA", typeof(double));
                dt5.Columns.Add("CUENTA FINANCIERA", typeof(long));

                var dt6 = new DataTable("-99");
                dt6.Columns.Add("CUENTA", typeof(long));
                dt6.Columns.Add("PCS", typeof(double));
                dt6.Columns.Add("MONTO FACTURADO", typeof(double));
                dt6.Columns.Add("MONTO SOLICITADO", typeof(double));
                dt6.Columns.Add("AJUSTE", typeof(double));
                dt6.Columns.Add("SOLICITANTE", typeof(string));

                foreach (var item in SUMADOS_CFM)
                {
                    double monto_iva = item.MONTO / 1.19;
                    double monto_ajustado = monto_iva - item.Monto_Total;
                    if (monto_ajustado < -99)
                    {
                        dt4.Rows.Add(new Object[] { long.Parse(item.Pcs), item.Monto_Total, item.MONTO, monto_iva, monto_ajustado, item.CUENTA });
                        dt6.Rows.Add(new Object[] { long.Parse(item.CUENTA), long.Parse(item.Pcs), item.Monto_Total, monto_iva, monto_ajustado, item.CAMPANA });
                    }
                    else
                    {
                        dt5.Rows.Add(new Object[] { long.Parse(item.Pcs), item.Monto_Total, item.MONTO, monto_iva, monto_ajustado, item.CUENTA });
                    }

                }

                oSLDocument.AddWorksheet("-99");
                oSLDocument.ImportDataTable(1, 1, dt4, true);

                oSLDocument.AddWorksheet("+99");
                oSLDocument.ImportDataTable(1, 1, dt5, true);

                oSLDocument3.ImportDataTable(1, 1, dt6, true);
                oSLDocument3.RenameWorksheet("Sheet1", "AJUSTE CFM");
                oSLDocument3.SaveAs(@aux_path + @"\Ciclo" + @"\AJUSTE CFM CICLO _XX.xlsx");

                oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\RESPALDO CFM.xlsx"); //esta listo el cfm///ESTO ESTA BUEEEEEEEEEEEEENARDOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO********************************************************************************************* CFM
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : "+ ex);
                return false;
            }
        }



    }
}
