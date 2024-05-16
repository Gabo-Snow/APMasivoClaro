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
   public class Controlador_Dummy
    {
       public List<Dummy> dummies = new List<Dummy>();
       List<Dummy> dummies_bc = new List<Dummy>();
       public List<Dummy> dummies_original = new List<Dummy>();
       public List<Cartera_Empresarial> cartera_Empresarials_sice;
       public List<Cartera> carteras_pablo;
       public Coleccion_Dummy coleccion_Dummy = new Coleccion_Dummy();
       public SLDocument oSLDocument = new SLDocument();
       public List<string> filtro_bc = new List<string>() { "bc" };
       public List<string> tipos_ = new List<string>();
       public string aux_path = string.Empty;
       SLDocument oSLDocument2 = new SLDocument();
       SLDocument oSLDocument3 = new SLDocument();
        public Controlador_Dummy(string path_dummy,string path_salida, List<Cartera_Empresarial> cartera_Empresarials, List<Cartera> carteras, List<string> filtro )
        {
            cartera_Empresarials_sice = cartera_Empresarials;
            carteras_pablo = carteras;
            dummies = coleccion_Dummy.GenerarListado(@path_dummy);//GenerarListado(@"C:\Users\Marcelo\Desktop\CICLO 4 trabajar este 26-04-2021\DUMMY 04.04.txt");
            tipos_ = filtro;
            aux_path = path_salida;
            Console.Beep();
        }

        
        Coleccion_Cartera_Empresarial _Cartera_Empresarial = new Coleccion_Cartera_Empresarial();
        Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();

        public bool aplicar_filtro(bool si_no)
        {
            if (si_no)
            {
                filtro_bc = new List<string>() { "bc" }; //por defecto necesito un txt que controle los cambios
                return true;
            }
            else
            {
                return false;
            }
        }

        public string filtrando_dummys() //tengo que hacer algo con esto del bc
        {
            string filtro = "bc";
            foreach (var item in dummies)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(item.nombreOFFER.ToLower(), filtro, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    Console.WriteLine($"  (match for '{filtro}' found)");
                    dummies_bc.Add(item);
                }
                else
                {
                    dummies_original.Add(item);
                }
            }
            return "termine";
        }


        public bool proceso_dummy()
        {
            try
            {
                //cambiar nombre de original a sin asignacion en todos
                //dummmy solo se cruza con sise
                //si no tienen pcs los cruzo con sise
                //si son dos cuentas se deben sumar para que solo quede una
                //tipo tiene que cambiar a cartera si es el tipo de cartera
                var cargos_linq_empresas = from arrendamiento in dummies_bc
                                           join sice in cartera_Empresarials_sice
                                            on arrendamiento.rut equals sice.RUT
                                           select new Dummy
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
                                               segmento = sice.SEGMENTO,
                                               rut = arrendamiento.rut
                                           };



                var cargos_linq_empresas_aux_1 = from arrendamiento in cargos_linq_empresas
                                                 join tipo in tipos_
                                                on arrendamiento.segmento equals tipo
                                                 select new Dummy
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
                                                     segmento = tipo,
                                                     rut = arrendamiento.rut
                                                 };
                List<Dummy> dummies_empresas = new List<Dummy>();
                dummies_empresas = cargos_linq_empresas_aux_1.ToList();
                HashSet<string> rut_bc = new HashSet<string>(cargos_linq_empresas_aux_1.Select(x => x.rut));
                dummies_bc.RemoveAll(x => rut_bc.Contains(x.rut));

                var cargos_linq_pablo = from arrendamiento in dummies_bc
                                        join sice in carteras_pablo
                                         on arrendamiento.CuentaFinanciera equals sice.BA_NO
                                        select new Dummy
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
                                            segmento = sice.SEGMENTO,
                                            rut = arrendamiento.rut
                                        };

                var cargos_linq_pablo_aux_1 = from arrendamiento in cargos_linq_pablo
                                              join tipo in tipos_
                                                on arrendamiento.segmento equals tipo
                                              select new Dummy
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
                                                  segmento = tipo,
                                                  rut = arrendamiento.rut
                                              };
                List<Dummy> dummies_pablo = new List<Dummy>();
                dummies_pablo = cargos_linq_pablo_aux_1.ToList();
                HashSet<string> cuenta_financiera_bc = new HashSet<string>(cargos_linq_pablo_aux_1.Select(x => x.CuentaFinanciera));
                dummies_bc.RemoveAll(x => rut_bc.Contains(x.rut));
                foreach (var item in dummies_empresas)
                {
                    dummies_pablo.Add(item); //solo dejar los sin pcs
                }

                var dt11 = new DataTable("EMPRESAS");
                dt11.Columns.Add("CuentaFinanciera", typeof(long));
                dt11.Columns.Add("recveivercustomer", typeof(long));
                dt11.Columns.Add("tipocargo", typeof(string));
                dt11.Columns.Add("codigodecargo", typeof(string));
                dt11.Columns.Add("offwer", typeof(long));
                dt11.Columns.Add("nombreOFFER", typeof(string));
                dt11.Columns.Add("promo", typeof(string));
                dt11.Columns.Add("description", typeof(string));
                dt11.Columns.Add("montoCargos", typeof(double));
                dt11.Columns.Add("montoDescuentos", typeof(double));
                dt11.Columns.Add("montoTotal", typeof(double));
                dt11.Columns.Add("tipodeCobro", typeof(string));
                dt11.Columns.Add("PCS", typeof(long));
                dt11.Columns.Add("secuenciadeCiclo", typeof(long));
                dt11.Columns.Add("RUT", typeof(string));
                dt11.Columns.Add("CARTERA", typeof(string));

                var dt12 = new DataTable("EMPRESAS");
                dt12.Columns.Add("CuentaFinanciera", typeof(long));
                dt12.Columns.Add("recveivercustomer", typeof(long));
                dt12.Columns.Add("tipocargo", typeof(string));
                dt12.Columns.Add("codigodecargo", typeof(string));
                dt12.Columns.Add("offwer", typeof(long));
                dt12.Columns.Add("nombreOFFER", typeof(string));
                dt12.Columns.Add("promo", typeof(string));
                dt12.Columns.Add("description", typeof(string));
                dt12.Columns.Add("montoCargos", typeof(double));
                dt12.Columns.Add("montoDescuentos", typeof(double));
                dt12.Columns.Add("montoTotal", typeof(double));
                dt12.Columns.Add("tipodeCobro", typeof(string));
                dt12.Columns.Add("PCS");
                dt12.Columns.Add("secuenciadeCiclo", typeof(long));
                dt12.Columns.Add("RUT", typeof(string));
                dt12.Columns.Add("CARTERA", typeof(string));

                foreach (var item in dummies_pablo)//sigue dando 0..... arreglarlo
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    if (item.PCS.Equals(""))
                    {
                        dt12.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
                                               long.Parse(item.recveivercustomer),
                                               item.tipocargo,
                                               item.codigodecargo,
                                               long.Parse(item.offwer),
                                               item.nombreOFFER,
                                               item.promo,
                                               item.description,
                                               item.montoCargos,
                                               item.montoDescuentos,
                                               item.montoTotal ,
                                               item.tipodeCobro,
                                               item.PCS,
                                               long.Parse(item.secuenciadeCiclo),
                                               rut_con_guion,
                                               item.segmento});
                    }
                    //else
                    //{
                    //    dt11.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
                    //                           long.Parse(item.recveivercustomer),
                    //                           item.tipocargo,
                    //                           item.codigodecargo,
                    //                           long.Parse(item.offwer),
                    //                           item.nombreOFFER,
                    //                           item.promo,
                    //                           item.description,
                    //                           item.montoCargos,
                    //                           item.montoDescuentos,
                    //                           item.montoTotal ,
                    //                           item.tipodeCobro,
                    //                           long.Parse(item.PCS),
                    //                           long.Parse(item.secuenciadeCiclo),
                    //                           rut_con_guion,
                    //                           item.segmento});
                    //}

                }
                oSLDocument3.ImportDataTable(1, 1, dt12, true);
                oSLDocument3.RenameWorksheet("Sheet1", "EMPRESAS");

                //oSLDocument3.AddWorksheet("SIN PCS"); //algo de este esta mal y debemos usar culture info para los montos y las letras
                //oSLDocument3.ImportDataTable(1, 1, dt12, true);

                var dt1 = new DataTable("SIN PCS");//NOMBRE POR AHORA
                dt1.Columns.Add("CuentaFinanciera", typeof(long));
                dt1.Columns.Add("recveivercustomer", typeof(long));
                dt1.Columns.Add("tipocargo", typeof(string));
                dt1.Columns.Add("codigodecargo", typeof(string));
                dt1.Columns.Add("offwer", typeof(long));
                dt1.Columns.Add("nombreOFFER", typeof(string));
                dt1.Columns.Add("promo", typeof(string));
                dt1.Columns.Add("description", typeof(string));
                dt1.Columns.Add("montoCargos", typeof(double));
                dt1.Columns.Add("montoDescuentos", typeof(double));
                dt1.Columns.Add("montoTotal", typeof(double));
                dt1.Columns.Add("tipodeCobro", typeof(string));
                dt1.Columns.Add("PCS", typeof(string));
                dt1.Columns.Add("secuenciadeCiclo", typeof(long));
                dt1.Columns.Add("prorrateo", typeof(string));
                dt1.Columns.Add("RUT", typeof(string));

                var dt2 = new DataTable("CON PCS");//NOMBRE POR AHORA
                dt2.Columns.Add("CuentaFinanciera", typeof(long));
                dt2.Columns.Add("recveivercustomer", typeof(long));
                dt2.Columns.Add("tipocargo", typeof(string));
                dt2.Columns.Add("codigodecargo", typeof(string));
                dt2.Columns.Add("offwer", typeof(long));
                dt2.Columns.Add("nombreOFFER", typeof(string));
                dt2.Columns.Add("promo", typeof(string));
                dt2.Columns.Add("description", typeof(string));
                dt2.Columns.Add("montoCargos", typeof(double));
                dt2.Columns.Add("montoDescuentos", typeof(double));
                dt2.Columns.Add("montoTotal", typeof(double));
                dt2.Columns.Add("tipodeCobro", typeof(string));
                dt2.Columns.Add("PCS", typeof(long));
                dt2.Columns.Add("secuenciadeCiclo", typeof(long));
                dt2.Columns.Add("prorrateo", typeof(string));
                dt2.Columns.Add("RUT", typeof(string));

                var dt3 = new DataTable("SIN PCS");//NOMBRE POR AHORA
                dt3.Columns.Add("CUENTA", typeof(long));
                dt3.Columns.Add("AMOUNT", typeof(double));
                dt3.Columns.Add("CODIGO DE CARGO", typeof(string));
                dt3.Columns.Add("NOMBRE OFFER", typeof(string));
                dt3.Columns.Add("CICLO");
                List<Dummy> compartidos_duplicados = new List<Dummy>();
                foreach (var item in dummies_bc)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    Dummy auxli_dummy = new Dummy();
                    if (item.PCS.Equals(""))
                    {
                        dt1.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
                                                           long.Parse(item.recveivercustomer),
                                                           item.tipocargo,
                                                           item.codigodecargo,
                                                           long.Parse(item.offwer),
                                                           item.nombreOFFER,
                                                           item.promo,
                                                           item.description,
                                                           item.montoCargos,
                                                           item.montoDescuentos,
                                                           item.montoTotal ,
                                                           item.tipodeCobro,
                                                           item.PCS,
                                                           long.Parse(item.secuenciadeCiclo),
                                                           item.prorrateo,
                                                           rut_con_guion});

                        double monto = item.montoTotal * -1;
                        auxli_dummy.CuentaFinanciera = item.CuentaFinanciera;
                        auxli_dummy.montoTotal = monto;
                        auxli_dummy.codigodecargo = item.codigodecargo;
                        auxli_dummy.nombreOFFER = item.nombreOFFER;
                        compartidos_duplicados.Add(auxli_dummy);

                    }
                    else
                    {
                        dt2.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
                                                           long.Parse(item.recveivercustomer),
                                                           item.tipocargo,
                                                           item.codigodecargo,
                                                           long.Parse(item.offwer),
                                                           item.nombreOFFER,
                                                           item.promo,
                                                           item.description,
                                                           item.montoCargos,
                                                           item.montoDescuentos,
                                                           item.montoTotal ,
                                                           item.tipodeCobro,
                                                           long.Parse(item.PCS),
                                                           long.Parse(item.secuenciadeCiclo),
                                                           item.prorrateo,
                                                           rut_con_guion});
                    }

                }

                var query = compartidos_duplicados.GroupBy(d => d.CuentaFinanciera)
                .Select(
                        g => new
                        {
                            Key = g.Key,
                            Monto_Total = g.Sum(s => s.montoTotal),
                            CodigoCarga = g.First().codigodecargo,
                            nombreOFFER = g.First().nombreOFFER
                        });

                foreach (var item in query)
                {
                    dt3.Rows.Add(new Object[] { long.Parse(item.Key),
                                                           Math.Round(item.Monto_Total),
                                                           item.CodigoCarga,
                                                           item.nombreOFFER,
                                                           ""});
                }
                //despues crear cargos compartidos
                oSLDocument.ImportDataTable(1, 1, dt1, true);
                oSLDocument.RenameWorksheet("Sheet1", "SIN PCS");//buscar los sin pcs por sice y pablo
                oSLDocument.AddWorksheet("CON PCS"); //algo de este esta mal y debemos usar culture info para los montos y las letras
                oSLDocument.ImportDataTable(1, 1, dt2, true);

                oSLDocument2.ImportDataTable(1, 1, dt3, true);
                oSLDocument2.RenameWorksheet("Sheet1", "AJUSTE");

                var dt = new DataTable("ORIGINAL");//NOMBRE POR AHORA
                dt.Columns.Add("CuentaFinanciera", typeof(long));
                dt.Columns.Add("recveivercustomer", typeof(long));
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
                dt.Columns.Add("secuenciadeCiclo", typeof(long));
                dt.Columns.Add("prorrateo", typeof(string));
                dt.Columns.Add("RUT", typeof(string));

                foreach (var item in dummies_original)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    if (item.PCS.Equals(""))
                    {
                        dt.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
                                                           long.Parse(item.recveivercustomer),
                                                           item.tipocargo,
                                                           item.codigodecargo,
                                                           item.offwer,
                                                           item.nombreOFFER,
                                                           item.promo,
                                                           item.description,
                                                           item.montoCargos,
                                                           item.montoDescuentos,
                                                           item.montoTotal ,
                                                           item.tipodeCobro,
                                                           0,
                                                           long.Parse(item.secuenciadeCiclo),
                                                           item.prorrateo,
                                                           rut_con_guion});
                    }
                    else
                    {
                        dt.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
                                                           long.Parse(item.recveivercustomer),
                                                           item.tipocargo,
                                                           item.codigodecargo,
                                                           item.offwer,
                                                           item.nombreOFFER,
                                                           item.promo,
                                                           item.description,
                                                           item.montoCargos,
                                                           item.montoDescuentos,
                                                           item.montoTotal ,
                                                           item.tipodeCobro,
                                                           long.Parse(item.PCS),
                                                           long.Parse(item.secuenciadeCiclo),
                                                           item.prorrateo,
                                                           rut_con_guion});
                    }

                }
                oSLDocument.AddWorksheet("ORIGINAL");
                oSLDocument.ImportDataTable(1, 1, dt, true);

                oSLDocument2.SaveAs(@aux_path + @"\Ciclo" + @"\Ruta" + @"\AJUSTES CARGOS COMPARTIDOS.xlsx");
                oSLDocument3.SaveAs(@aux_path + @"\Ciclo" + @"\DUMMYS EMPRESAS _CICLO_X_X.xlsx");
                oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\Extraidas" + @"\DUMMYS _CICLO_X_X.xlsx"); //nos quedamos sin memoria hay que dividirlo


                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("error" + ex);
                return false;
            }
        }




    }
}
