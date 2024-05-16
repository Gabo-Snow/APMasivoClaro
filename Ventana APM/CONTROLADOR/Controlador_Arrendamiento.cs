using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.CONTROLADOR
{
    public class Controlador_Arrendamiento
    {
        //DESDE AQUI CARGOS DE ARRENDAMIENTO BIEN HECHOS
        //----------------------------colecciones--------------------------------------------------------
        Coleccion_c_ar coleccion_C_Ar = new Coleccion_c_ar();
        Coleccion_ajustes coleccion_Ajustes = new Coleccion_ajustes();
        Coleccion_EComerce Coleccion_EComerce = new Coleccion_EComerce();
        //-----------------------------------------------------------------------------------------------
        StringBuilder sb = new System.Text.StringBuilder();//aqui esta el controlador de mensaje de errores
        //---------------------------------Listas--------------------------------------------------------
        List<Cartera_Empresarial> cartera_Empresarials_sice;
        List<Cartera> carteras_pablo;
        List<Cargos_Arrendamiento> cargos;
       // List<Cargos_Arrendamiento> cargos_limpios;
        List<Cargos_Arrendamiento> cargos_empresas = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> cargos_pymes = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> ajustes_especiales = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> E_COMERCE = new List<Cargos_Arrendamiento>();
        List<string> tipos_ = new List<string>();
        List<string> cargos_menores = new List<string>();
        List<string> cargos_mayores = new List<string>();
        List<string> PCS_UNICO = new List<string>();
        List<Ajustes> ajustes;
        List<eComerce> eComerces_list = new List<eComerce>();
        string aux_path = string.Empty;
       
        ////-----------------------------------------------------------------------------------------------

        //---------------------------------------Excel--------------------------------------------------
        SLDocument oSLDocument = new SLDocument();
        SLDocument oSLDocument2 = new SLDocument();
        SLDocument oSLDocument3 = new SLDocument();
        //---------------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------------
        public Controlador_Arrendamiento(string path_cargos, string path_ajustes,string ecomer, string path_salida, List<Cartera_Empresarial> cartera_Empresarials, List<Cartera> carteras, List<string> tipos_x)
        {
            cartera_Empresarials_sice = cartera_Empresarials;// _Cartera_Empresarial.GenerarListado(@"C:\Users\Marcelo\Desktop\Ahora si final final final final final\Como crear excel con varias hojas\Ciclo 18 entrada\Cartera actualizada 23-02-21.xlsx");
            cargos = coleccion_C_Ar.GenerarListado(@path_cargos);//GenerarListado(@"C:\Users\Marcelo\Desktop\Ahora si final final final final final\Como crear excel con varias hojas\Ciclo 18 entrada\CARGOS_ARRENDAMIENTOS.txt");
            Console.Beep();
            carteras_pablo = carteras;//coleccion_Cartera.GenerarListado(@"C:\Users\Marcelo\Desktop\Ahora si final final final final final\Como crear excel con varias hojas\Ciclo 18 entrada\Cartera 08.02.2021 _PABLO.xlsx");
            ajustes = coleccion_Ajustes.GenerarListado(@path_ajustes); //controlar la excepcion hoy todo lo de esta ventana tiene que quedar hoy finiquitado se fini
            eComerces_list = Coleccion_EComerce.GenerarListado(@ecomer);//GenerarListado(@"C:\Users\Marcelo\Desktop\Ahora si final final final final final\Como crear excel con varias hojas\Ciclo 18 entrada\C_ARRENDAMIENTO AJUSTES ESPECIALES C18.01.2021.xlsx");
            tipos_ = tipos_x;
            foreach (var item in cargos)
            {
                if (item.segmento.Equals(""))
                {
                    item.segmento = "N/A";
                }
                if (item.prorrateo.Equals(""))
                {
                    item.prorrateo = "C";
                }
            }
            aux_path = path_salida;
            Console.Beep();
        }

        public bool corroboracion_de_colecciones_empresas_0()
        {
            return true;
            try
            {

            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public bool corroboracion_de_colecciones_empresas_pablo_0()
        {
            return true;
            try
            {

            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public bool corroboracion_de_colecciones_cargos_0()
        {
            return true;
            try
            {

            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public bool corroboracion_de_colecciones_ajustes_0()
        {
            return true;
            try
            {

            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public bool corroboracion_de_colecciones_ecomerce_0()
        {
            return true;
            try
            {

            }
            catch (Exception ex)
            {

                throw;
            }
        }
        public bool corroboracion_de_colecciones_0()//tiene que estar en una forma de corroborar uno por uno para todos
        {
            string erro = "";
            
            if (cartera_Empresarials_sice.Count() == 0)
            {
                erro = "la cartera empresarial esta vacia!";
                sb.AppendLine(erro);
            }
            if (carteras_pablo.Count() == 0)
            {
                erro = "la cartera actualizada esta vacia!";
                sb.AppendLine(erro);
            }
            if (cargos.Count() == 0)
            {
                erro = "los cargos de arrendamiento estan vacios!";
                sb.AppendLine(erro);
            }
            if (ajustes.Count() == 0)
            {
                erro = "los ajustes estan vacios!";
                sb.AppendLine(erro);
            }
            if (eComerces_list.Count() == 0)
            {
                erro = "eComerce esta vacio!";
                sb.AppendLine(erro);
            }

            if (sb.Length == 0)
            {
                return false;

            }
            else
            {
                MessageBox.Show("\n" + sb, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Console.Error.WriteLine();
                return true;
            }



            
        }
        public bool proceso_arrendamiento_4_5()
        {
            try
            {
            var dt = new DataTable("ORIGINAL");
            dt.Columns.Add("CuentaFinanciera", typeof(long));
            dt.Columns.Add("recveivercustomer", typeof(long));
            dt.Columns.Add("tipocargo", typeof(string));
            dt.Columns.Add("codigodecargo", typeof(string));
            dt.Columns.Add("offwer", typeof(long));
            dt.Columns.Add("nombreOFFER", typeof(string));
            dt.Columns.Add("promo", typeof(string));
            dt.Columns.Add("description", typeof(string));
            dt.Columns.Add("montoCargos", typeof(double));
            dt.Columns.Add("montoDescuentos", typeof(double));
            dt.Columns.Add("montoTotal", typeof(double));
            dt.Columns.Add("tipodeCobro", typeof(string));
            dt.Columns.Add("PCS", typeof(long));
            dt.Columns.Add("secuenciadeCiclo", typeof(long));
            dt.Columns.Add("TIPO", typeof(string));
            dt.Columns.Add("RUT", typeof(string));
            dt.Columns.Add("CARTERA", typeof(string));

            


            foreach (var item in cargos)
            {
                int contador_para_guion = item.rut.Length - 1;
                string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                dt.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               rut_con_guion,
                                               item.segmento});
            }



            oSLDocument.AddWorksheet("SIN ASIGNAR");
            oSLDocument.ImportDataTable(1, 1, dt, true);


                // TODO PERFECT LOS DE ARRENDAMIENTO
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex);
                return false;
            }
        }//PERFECTO //en ves tipo deberia decir cartera y en vez de prorrateo tipo
        public bool proceso_arrendamiento_empresas_1()
        {
            try
            {
                var cargos_linq_empresas = from arrendamiento in cargos
                                           join sice in cartera_Empresarials_sice
                                                on arrendamiento.rut equals sice.RUT
                                           select new Cargos_Arrendamiento
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
                                                 select new Cargos_Arrendamiento
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
                List<Cargos_Arrendamiento> auxiliar_empresas_1 = cargos_linq_empresas_aux_1.ToList();

                HashSet<string> rut_empresas_empresarial = new HashSet<string>(cargos_linq_empresas_aux_1.Select(x => x.rut));

                cargos.RemoveAll(x => rut_empresas_empresarial.Contains(x.rut));
                //algo esta malo o ingresaron mal los datos esa es la cuestion...
                //chum chumchumchum
                List<Cargos_Arrendamiento> auxiliar_pyme_1 = cargos_linq_empresas.ToList();
                HashSet<string> rut_pyme_empresarial = new HashSet<string>(cargos_linq_empresas.Select(x => x.rut));
                cargos.RemoveAll(x => rut_pyme_empresarial.Contains(x.rut));
                int count = cargos_linq_empresas.Count();

                var cargos_linq_pablo = from arrendamiento in cargos
                                        join pablo in carteras_pablo
                                             on arrendamiento.CuentaFinanciera equals pablo.BA_NO
                                        select new Cargos_Arrendamiento
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
                                            segmento = pablo.SEGMENTO,
                                            rut = arrendamiento.rut
                                        };

                var cargos_linq_empresas_aux_2 = from arrendamiento in cargos_linq_pablo
                                                 join tipo in tipos_
                                                on arrendamiento.segmento equals tipo
                                                 select new Cargos_Arrendamiento
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
                List<Cargos_Arrendamiento> auxiliar_empresas_2 = cargos_linq_empresas_aux_2.ToList();
                HashSet<string> cuenta_financiera_empresas_empresarial = new HashSet<string>(cargos_linq_empresas_aux_2.Select(x => x.CuentaFinanciera));
                cargos.RemoveAll(x => cuenta_financiera_empresas_empresarial.Contains(x.CuentaFinanciera));

                List<Cargos_Arrendamiento> auxiliar_pyme_2 = cargos_linq_pablo.ToList();
                HashSet<string> rut_pyme_pablo = new HashSet<string>(cargos_linq_pablo.Select(x => x.CuentaFinanciera));
                cargos.RemoveAll(x => rut_pyme_pablo.Contains(x.CuentaFinanciera));
                foreach (var item in auxiliar_empresas_2)
                {
                    auxiliar_empresas_1.Add(item);
                }
                foreach (var item in auxiliar_pyme_2)
                {
                    auxiliar_pyme_1.Add(item);
                }
                foreach (var item in auxiliar_pyme_1)
                {
                    cargos.Add(item);
                }
                Console.Beep();

                int count1 = cargos_linq_empresas.Count(); //primero va pablo porque el que domina es empresas
                int count2 = cargos_linq_pablo.Count();
                var dt1 = new DataTable("EMPRESAS");
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
                dt1.Columns.Add("PCS", typeof(long));
                dt1.Columns.Add("secuenciadeCiclo", typeof(long));
                dt1.Columns.Add("TIPO", typeof(string));
                dt1.Columns.Add("RUT", typeof(string));
                dt1.Columns.Add("CARTERA", typeof(string));

                foreach (var item in auxiliar_empresas_1)//sigue dando 0..... arreglarlo
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
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
                                               long.Parse(item.PCS),
                                               long.Parse(item.secuenciadeCiclo),
                                               item.prorrateo,
                                               rut_con_guion,
                                               item.segmento});
                }
                oSLDocument.ImportDataTable(1, 1, dt1, true);
                oSLDocument.RenameWorksheet("Sheet1", "EMPRESAS");
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        } //perfecto
        public bool proceso_arrendamiento_neteados_2()
        {
            try
            {
                List<Cargos_Arrendamiento> cargos_neteados = new List<Cargos_Arrendamiento>();
                List<Cargos_Arrendamiento> cargos_sumados = new List<Cargos_Arrendamiento>();

                var arrendamiento_sumados = cargos.GroupBy(d => d.PCS)
                            .Select(
                                    g => new
                                    {
                                        Key = g.Key,
                                        Monto_Total = g.Sum(s => s.montoTotal),
                                        Rut = g.First().rut,
                                        Pcs = g.First().PCS
                                    });
                int contador_tabla_dinamica = arrendamiento_sumados.Count();


                foreach (var item_1 in arrendamiento_sumados)
                {
                    if (item_1.Monto_Total <= 100.0)//
                    {
                        cargos_menores.Add(item_1.Pcs);
                    }
                    else
                    {
                        cargos_mayores.Add(item_1.Pcs);
                    }
                }

                foreach (var item_1 in cargos_menores)
                {
                    foreach (var item_2 in cargos)
                    {
                        if (item_1.Equals(item_2.PCS))
                        {
                            cargos_neteados.Add(item_2);
                        }
                    }
                }

                HashSet<string> pcs_suma = new HashSet<string>(cargos_neteados.Select(x => x.PCS));
                cargos.RemoveAll(x => pcs_suma.Contains(x.PCS));


                var dt2 = new DataTable("NETEADOS");
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
                dt2.Columns.Add("RUT", typeof(string));
                dt2.Columns.Add("TIPO", typeof(string));
                //----------

                foreach (var item in cargos_neteados)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
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
                                               rut_con_guion,
                                               "NETEADOS"});
                }
                oSLDocument.AddWorksheet("NETEADOS");
                oSLDocument.ImportDataTable(1, 1, dt2, true);
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        } //PERFECTO
        public bool proceso_arrendamiento_ajustes_3()
        {
            try
            {
                var ajustes_linq_arrendamiento = from carg in cargos
                                                 join ajust in ajustes
                                                      on carg.PCS equals ajust.PCS
                                                 select new Cargos_Arrendamiento
                                                 {
                                                     CuentaFinanciera = carg.CuentaFinanciera,
                                                     recveivercustomer = carg.recveivercustomer,
                                                     tipocargo = carg.tipocargo,
                                                     codigodecargo = carg.codigodecargo,
                                                     offwer = carg.offwer,
                                                     nombreOFFER = carg.nombreOFFER,
                                                     promo = carg.promo,
                                                     description = carg.description,
                                                     montoCargos = carg.montoCargos,
                                                     montoDescuentos = carg.montoDescuentos,
                                                     montoTotal = carg.montoTotal,
                                                     tipodeCobro = carg.tipodeCobro,
                                                     PCS = carg.PCS,
                                                     secuenciadeCiclo = carg.secuenciadeCiclo,
                                                     prorrateo = carg.prorrateo,
                                                     segmento = carg.segmento,
                                                     rut = carg.rut
                                                 };

                ajustes_especiales = ajustes_linq_arrendamiento.ToList();

                HashSet<string> pcs_ajustes = new HashSet<string>(ajustes_linq_arrendamiento.Select(x => x.PCS));
                cargos.RemoveAll(x => pcs_ajustes.Contains(x.PCS));

                //shopppi pium pium
                var dt3 = new DataTable("AJUSTES");
                dt3.Columns.Add("CuentaFinanciera", typeof(long));
                dt3.Columns.Add("recveivercustomer", typeof(long));
                dt3.Columns.Add("tipocargo", typeof(string));
                dt3.Columns.Add("codigodecargo", typeof(string));
                dt3.Columns.Add("offwer", typeof(long));
                dt3.Columns.Add("nombreOFFER", typeof(string));
                dt3.Columns.Add("promo", typeof(string));
                dt3.Columns.Add("description", typeof(string));
                dt3.Columns.Add("montoCargos", typeof(double));
                dt3.Columns.Add("montoDescuentos", typeof(double));
                dt3.Columns.Add("montoTotal", typeof(double));
                dt3.Columns.Add("tipodeCobro", typeof(string));
                dt3.Columns.Add("PCS", typeof(long));
                dt3.Columns.Add("secuenciadeCiclo", typeof(long));
                dt3.Columns.Add("TIPO", typeof(string));
                dt3.Columns.Add("RUT", typeof(string));
                dt3.Columns.Add("CARTERA", typeof(string));


                foreach (var item in ajustes_especiales)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    dt3.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               rut_con_guion,
                                               item.segmento});
                }
                var suma_ajustes = ajustes_especiales.GroupBy(d => d.PCS)
                    .Select(
                            g => new
                            {
                                Key = g.Key,
                                Monto_Total = g.Sum(s => s.montoTotal),
                                CuentaFinanciera = g.First().CuentaFinanciera,
                                Pcs = g.First().PCS,
                                CodigoCarga = g.First().codigodecargo
                            });

                var dt4 = new DataTable("AJUSTES");
                dt4.Columns.Add("PCS", typeof(long));
                dt4.Columns.Add("CUENTA FINANCIERA", typeof(long));
                dt4.Columns.Add("CODIGO DE CARGO", typeof(string));
                dt4.Columns.Add("AMOUNT", typeof(double));
                dt4.Columns.Add("CICLO", typeof(string));


                foreach (var item in suma_ajustes)
                {
                    if (item.Monto_Total > 0.0)
                    {
                        dt4.Rows.Add(new Object[] { long.Parse(item.Pcs), long.Parse(item.CuentaFinanciera), item.CodigoCarga, "-" + Math.Round(item.Monto_Total), "XX" });
                    }
                }

                oSLDocument2.ImportDataTable(1, 1, dt4, true);
                oSLDocument2.RenameWorksheet("Sheet1", "AJUSTES"); 
                Console.Beep();
                // SaveExcel.BuildExcel(dt4, @"C:\Users\Marcelo\Desktop\" + @"\AJUSTES ESPECIALES_CICLO_X_X.xlsx");
                oSLDocument.AddWorksheet("AJUSTES ESPECIALES");
                oSLDocument.ImportDataTable(1, 1, dt3, true);
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }//perfecto
        public bool proceso_arrendamiento_ecomerce_4()//esto falla pero porque falla
        {
            try
            {
                var comerce_linq_arrendamiento = from carg in cargos
                                                 join ajust in eComerces_list
                                                      on carg.PCS equals ajust.PCS
                                                 select new Cargos_Arrendamiento
                                                 {
                                                     CuentaFinanciera = carg.CuentaFinanciera,
                                                     recveivercustomer = carg.recveivercustomer,
                                                     tipocargo = carg.tipocargo,
                                                     codigodecargo = carg.codigodecargo,
                                                     offwer = carg.offwer,
                                                     nombreOFFER = carg.nombreOFFER,
                                                     promo = carg.promo,
                                                     description = carg.description,
                                                     montoCargos = carg.montoCargos,
                                                     montoDescuentos = carg.montoDescuentos,
                                                     montoTotal = carg.montoTotal,
                                                     tipodeCobro = carg.tipodeCobro,
                                                     PCS = carg.PCS,
                                                     secuenciadeCiclo = carg.secuenciadeCiclo,
                                                     prorrateo = carg.prorrateo,
                                                     segmento = carg.segmento,
                                                     rut = carg.rut
                                                 };

                E_COMERCE = comerce_linq_arrendamiento.ToList();

                HashSet<string> pcs_ajustes = new HashSet<string>(comerce_linq_arrendamiento.Select(x => x.PCS));
                cargos.RemoveAll(x => pcs_ajustes.Contains(x.PCS));

                var dt3 = new DataTable("ECOMERCE");
                dt3.Columns.Add("CuentaFinanciera", typeof(long));
                dt3.Columns.Add("recveivercustomer", typeof(long));
                dt3.Columns.Add("tipocargo", typeof(string));
                dt3.Columns.Add("codigodecargo", typeof(string));
                dt3.Columns.Add("offwer", typeof(long));
                dt3.Columns.Add("nombreOFFER", typeof(string));
                dt3.Columns.Add("promo", typeof(string));
                dt3.Columns.Add("description", typeof(string));
                dt3.Columns.Add("montoCargos", typeof(double));
                dt3.Columns.Add("montoDescuentos", typeof(double));
                dt3.Columns.Add("montoTotal", typeof(double));
                dt3.Columns.Add("tipodeCobro", typeof(string));
                dt3.Columns.Add("PCS", typeof(long));
                dt3.Columns.Add("secuenciadeCiclo", typeof(long));
                dt3.Columns.Add("TIPO", typeof(string));
                dt3.Columns.Add("RUT", typeof(string));
                dt3.Columns.Add("CARTERA", typeof(string));

                var e_comerce_no_repetidos = E_COMERCE.Distinct();
                //bueh
                foreach (var item in e_comerce_no_repetidos.Distinct())//esto deberia servir para que no se dupliquen
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    dt3.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               rut_con_guion,
                                               item.segmento});
                }

                oSLDocument.AddWorksheet("E COMERCE");
                oSLDocument.ImportDataTable(1, 1, dt3, true);
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        } //perfecto
        public bool proceso_arrendamiento_duplicados_5()
        {
            try
            {
                List<Cargos_Arrendamiento> cargos_duplicados = new List<Cargos_Arrendamiento>();
                int repetidos = 0;
                foreach (var item1 in cargos)
                {
                    repetidos = 0;
                    foreach (var item2 in cargos)
                    {
                        if (item1.PCS.Equals(item2.PCS))
                        {
                            repetidos++;
                        }
                    }
                    if (repetidos >= 2)
                    {
                        item1.duplicado = "ES DULPICADO";
                        cargos_duplicados.Add(item1); //esta bien pero puede ser mejor
                    }
                }

                //var duplicados_linq = cargos.GroupBy(x => new { x.PCS })
                //                   .Where(x => x.Skip(1).Any());
                //foreach (var item1 in cargos)
                //{
                //    foreach (var item2 in duplicados_linq)
                //    {
                //        if (item2.Key.PCS.Equals(item1.PCS))
                //        {
                //            item1.duplicado = "ES DULPICADO";
                //            cargos_duplicados.Add(item1);
                //            break;
                //        }

                //    }
                //}
               //int duplicados =  duplicados_linq.Count();


                HashSet<string> pcs_duplicados = new HashSet<string>(cargos_duplicados.Select(x => x.PCS)); //revisar los duplicados.
                cargos.RemoveAll(x => pcs_duplicados.Contains(x.PCS));

                foreach (var item_1 in cargos)
                {
                    if (item_1.montoTotal >= 14000.0)
                    {
                        cargos_duplicados.Add(item_1);
                    }
                }

                HashSet<string> pcs_mayores_14 = new HashSet<string>(cargos_duplicados.Select(x => x.PCS));
                cargos.RemoveAll(x => pcs_mayores_14.Contains(x.PCS));

                var dt5 = new DataTable("DUPLICADOS");
                dt5.Columns.Add("CuentaFinanciera", typeof(long));
                dt5.Columns.Add("recveivercustomer", typeof(long));
                dt5.Columns.Add("tipocargo", typeof(string));
                dt5.Columns.Add("codigodecargo", typeof(string));
                dt5.Columns.Add("offwer", typeof(long));
                dt5.Columns.Add("nombreOFFER", typeof(string));
                dt5.Columns.Add("promo", typeof(string));
                dt5.Columns.Add("description", typeof(string));
                dt5.Columns.Add("montoCargos", typeof(double));
                dt5.Columns.Add("montoDescuentos", typeof(double));
                dt5.Columns.Add("montoTotal", typeof(double));
                dt5.Columns.Add("tipodeCobro", typeof(string));
                dt5.Columns.Add("PCS", typeof(long));
                dt5.Columns.Add("secuenciadeCiclo", typeof(long));
                dt5.Columns.Add("TIPO", typeof(string));
                dt5.Columns.Add("RUT", typeof(string));
                dt5.Columns.Add("CARTERA", typeof(string));
                dt5.Columns.Add("DUPLICADO", typeof(string));

                var duplicados_ordenados_pcs = cargos_duplicados.OrderBy(q => q.duplicado).ToList();

                foreach (var item in duplicados_ordenados_pcs)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    dt5.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               rut_con_guion,
                                               item.segmento,
                                               item.duplicado});
                }

                oSLDocument.AddWorksheet("CARGOS DE ARRENDAMIENTO");//LOS DUPLICADOS ESTAN CORRECTOS
                oSLDocument.ImportDataTable(1, 1, dt5, true);

                var dt6 = new DataTable("CARGOS DE ARRENDAMIENTO");//NOMBRE POR AHORA
                dt6.Columns.Add("ASIGNAR", typeof(string));
                dt6.Columns.Add("CuentaFinanciera", typeof(long));
                dt6.Columns.Add("recveivercustomer", typeof(long));
                dt6.Columns.Add("tipocargo", typeof(string));
                dt6.Columns.Add("codigodecargo", typeof(string));
                dt6.Columns.Add("offwer", typeof(long));
                dt6.Columns.Add("nombreOFFER", typeof(string));
                dt6.Columns.Add("promo", typeof(string));
                dt6.Columns.Add("description", typeof(string));
                dt6.Columns.Add("montoCargos", typeof(double));
                dt6.Columns.Add("montoDescuentos", typeof(double));
                dt6.Columns.Add("montoTotal", typeof(double));
                dt6.Columns.Add("tipodeCobro", typeof(string));
                dt6.Columns.Add("PCS", typeof(long));
                dt6.Columns.Add("secuenciadeCiclo", typeof(long));
                dt6.Columns.Add("TIPO", typeof(string));
                dt6.Columns.Add("RUT", typeof(string));
                dt6.Columns.Add("CARTERA", typeof(string));
                dt6.Columns.Add("DUPLICADO", typeof(string));
                dt6.Columns.Add("COBRAR", typeof(string));
                dt6.Columns.Add("MONTO", typeof(string));
                dt6.Columns.Add("CUOTA", typeof(string));
                dt6.Columns.Add("OBSERVACION", typeof(string));

                foreach (var item in duplicados_ordenados_pcs)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    dt6.Rows.Add(new Object[] { "",
                                               long.Parse(item.CuentaFinanciera),
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
                                               rut_con_guion,
                                               item.segmento,
                                               item.duplicado,
                                               "",
                                               "",
                                               "",
                                               ""});
                }

                oSLDocument3.ImportDataTable(1, 1, dt6, true);
                oSLDocument3.RenameWorksheet("Sheet1", "CARGOS DE ARRENDAMIENTO");
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }//PERFECTO para amar para amar tienes tratar de todo entregar ._. 
        public bool proceso_arrendamiento_guardando_6()
        {
            try
            {
                //oSLDocument.SaveAs(@aux_path   + @"\Ciclo\Ruta\CARGOS DE ARRENDAMIENTO CICLO XX.xlsx");
                //oSLDocument2.SaveAs(@aux_path  + @"\Ciclo\Ruta\ARRENDAMIENTO AJUSTES ESPECIALES_CICLO_X_X.xlsx");
                //oSLDocument3.SaveAs(@aux_path  + @"\Ciclo\Extraidas\ASIGNAR_CARGOS DE ARRENDAMIENTO _CICLO_X_X.xlsx");
                //System.Media.SystemSounds.Exclamation.Play();
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }//PERFECTO


       
    }
}
