using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using Ventana_APM.Auxiliares;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.CONTROLADOR
{
    public class Controlador_Activacion
    {
        string aux_path = string.Empty;
        Coleccion_Activacion coleccion_Activacion = new Coleccion_Activacion();
        Coleccion_PlanSolidario Coleccion_PlanSolidario = new Coleccion_PlanSolidario();
        List<PlanSolidario> planSolidarios;
        List<Portabilidad> portabilidads;
        Coleccion_Portabilidad coleccion_Portabilidad = new Coleccion_Portabilidad();
        List<FVM> fVMs;
        Coleccion_FVM coleccion_FVM = new Coleccion_FVM();
        List<Activacion> activacion_neteados = new List<Activacion>();
        List<HABILITACIONES> hABILITACIONEs;
        Coleccion_HABILITACIONES coleccion_HABILITACIONES = new Coleccion_HABILITACIONES();
        List<Consolidacion_Validacion> consolidacion_Validacions = new List<Consolidacion_Validacion>();

        List<Cartera_Empresarial> cartera_Empresarials_sice;
        List<Cartera> carteras_pablo;
        List<Activacion> activacions;
        List<Activacion> activacions_asignados = new List<Activacion>();
        List<Activacion> activacions_asignados_portados = new List<Activacion>();
        List<Activacion> activacions_revisados_asignados = new List<Activacion>();
        //------------------
        List<PCS_SOBRANTES> pCS_SOBRANTEs;
        Colecccion_Sobrantes Colecccion_Sobrantes = new Colecccion_Sobrantes();
        List<Activacion> activacion_portados = new List<Activacion>();
        //------------------
        List<string> tipos_ = new List<string>();
        string path_portados = string.Empty;
        string path_fvms = string.Empty;
        string path_habili = string.Empty;
        string path_solidario = string.Empty;
        string path_conso = string.Empty;
        string carpeta_respaldo = string.Empty;
        string carpeta_asignar = string.Empty;
        string carpeta_reportar = string.Empty;
        SLDocument oSLDocument = new SLDocument();
        SLDocument oSLDocument2 = new SLDocument();
        SLDocument oSLDocument3 = new SLDocument();
        SLDocument oSLDocument4 = new SLDocument();
        SLDocument oSLDocument5 = new SLDocument();
        SLDocument oSLDocument6 = new SLDocument();//^~--._.--~^
        SLDocument oSLDocument7 = new SLDocument();//^~--._.--~^
        double monto_filtrado;
        long offerPorDefecto = 1013;
        long pcsPorDefecto = 0;


        public Controlador_Activacion(string path_activacion,string path_portabilidad, string path_fvm, string path_habilitaciones,string path_soli, string path_salida,string path_consolidado, List<Cartera_Empresarial> cartera_Empresarials, List<Cartera> carteras, List<string> filtro,double monto_para_filtrar)
        {

            cartera_Empresarials_sice = cartera_Empresarials;

            carteras_pablo = carteras;

            tipos_ = filtro;

            activacions = coleccion_Activacion.GenerarListado(@path_activacion);

            path_portados = path_portabilidad;

            aux_path = path_salida;

            path_fvms = path_fvm;

            path_habili = path_habilitaciones;

            path_solidario = path_soli;

            monto_filtrado = monto_para_filtrar;

            path_conso = path_consolidado;

            


        }

        public bool proceso_activacion_empresas_1()
        {
            try
            {
                var cargos_linq_empresas = from arrendamiento in activacions
                                           join sice in cartera_Empresarials_sice
                                                on arrendamiento.rut equals sice.RUT
                                           select new Activacion
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
                                                 select new Activacion
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
                List<Activacion> auxiliar_empresas_1 = cargos_linq_empresas_aux_1.ToList();

                HashSet<string> rut_empresas_empresarial = new HashSet<string>(cargos_linq_empresas_aux_1.Select(x => x.rut));

                activacions.RemoveAll(x => rut_empresas_empresarial.Contains(x.rut));

                List<Activacion> auxiliar_pyme_1 = cargos_linq_empresas.ToList();
                HashSet<string> rut_pyme_empresarial = new HashSet<string>(cargos_linq_empresas.Select(x => x.rut));
                activacions.RemoveAll(x => rut_pyme_empresarial.Contains(x.rut));//masca_ int char list wabaluvdob
                
                int count = cargos_linq_empresas.Count();

                var cargos_linq_pablo = from arrendamiento in activacions
                                        join pablo in carteras_pablo
                                             on arrendamiento.CuentaFinanciera equals pablo.BA_NO
                                        select new Activacion
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
                                            rut = arrendamiento.rut//fusroda fus ro da f u s r o d a dobakin tururu turu ru turu ru tu tu ru tu ru ru turu ru tu ru ru tu tu tu
                                        };
                 
                var cargos_linq_empresas_aux_2 = from arrendamiento in cargos_linq_pablo
                                                 join tipo in tipos_
                                                on arrendamiento.segmento equals tipo
                                                 select new Activacion
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
                List<Activacion> auxiliar_empresas_2 = cargos_linq_empresas_aux_2.ToList();
                HashSet<string> cuenta_financiera_empresas_empresarial = new HashSet<string>(cargos_linq_empresas_aux_2.Select(x => x.CuentaFinanciera));
                activacions.RemoveAll(x => cuenta_financiera_empresas_empresarial.Contains(x.CuentaFinanciera));

                List<Activacion> auxiliar_pyme_2 = cargos_linq_pablo.ToList();
                HashSet<string> rut_pyme_pablo = new HashSet<string>(cargos_linq_pablo.Select(x => x.CuentaFinanciera));
                activacions.RemoveAll(x => rut_pyme_pablo.Contains(x.CuentaFinanciera));

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
                    activacions.Add(item);
                }
                

                int count1 = cargos_linq_empresas.Count(); //HASTA AQUI BORRA BIEN
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
                dt1.Columns.Add("prorrateo", typeof(string));
                dt1.Columns.Add("RUT", typeof(string));
                dt1.Columns.Add("CARTERA", typeof(string));


                foreach (var item in auxiliar_empresas_1)//sigue dando 0..... arreglarlo
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }

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
                int counte = auxiliar_empresas_1.Count();
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }  //TODO PERFECTO
        public bool proceso_activacion_portados_2()
        {
            try
            {
                portabilidads = coleccion_Portabilidad.GenerarListado(@path_portados);

                foreach (var item1 in activacions)
                {
                    string TIPO = string.Empty;
                    foreach (var item2 in portabilidads)
                    {
                        if (item1.PCS.Equals(item2.NRO_DE_PCS))
                        {
                            TIPO = item2.MODALIDAD;
                        }
                    }
                    item1.TIPO = TIPO;
                    if (TIPO.Equals(""))
                    {
                        item1.TIPO = "#N/A";
                    }
                }
                int contador_1 = 0;
                int contador_2 = 0;
                int contador_3 = 0;

                foreach (var item in activacions)
                {
                    if (item.TIPO.Equals("PRE A POST"))
                    {
                        contador_1++;
                    }
                    else if (item.TIPO.Equals("PORTADO"))
                    {
                        contador_2++;
                    }
                    else if (item.TIPO.Equals("#N/A"))
                    {
                        contador_3++;
                    }
                }

                //boca lo mah grande 
                var query = activacions.GroupBy(d => d.PCS)
                                            .Select(
                                                    g => new
                                                    {
                                                        Key = g.Key,
                                                        Monto_Total = g.Sum(s => s.montoTotal),
                                                        CuentaFinanciera = g.First().CuentaFinanciera,
                                                        Pcs = g.First().PCS,
                                                        CodigoCarga = g.First().codigodecargo
                                                    }).OrderBy(g => g.Monto_Total);

                int contador_4 = 0;
                int contador_de_zeros = 0;
                int cuantos_del_sum = query.Count();

                foreach (var item1 in query)
                {
                    if (item1.Monto_Total == 0 || item1.Monto_Total == 0.0)
                    {
                        contador_de_zeros++;
                    }
                }
                foreach (var item1 in query)
                {
                    if (item1.Monto_Total == 0 || item1.Monto_Total == 0.0)
                    {
                        foreach (var item2 in activacions)
                        {
                            if (item1.Pcs.Equals(item2.PCS))
                            {
                                item2.TIPO = "NETEADO";
                                contador_4++;
                            }
                        }
                    }

                }
                //AQUI CREAR C_ACTIVACION

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
                dt2.Columns.Add("prorrateo", typeof(string));
                dt2.Columns.Add("RUT", typeof(string));
                dt2.Columns.Add("TIPO", typeof(string));

                
                int CONTADOR_AHORA = activacions.Count();

                try
                {
                    foreach (var item in activacions)
                    {
                        int contador_para_guion = item.rut.Length - 1;
                        string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                        bool succes = long.TryParse(item.offwer, out long numero);
                        bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                        if (succes)
                        {
                            item.offwer = numero.ToString();
                        }
                        else
                        {
                            item.offwer = offerPorDefecto.ToString();
                        }
                        if (pcsNoNull)
                        {
                            item.PCS = pcsNoNulo.ToString();
                        }
                        else
                        {
                            item.PCS = pcsPorDefecto.ToString();
                        }

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
                                               rut_con_guion,
                                               item.TIPO});
                    }
                    oSLDocument2.ImportDataTable(1, 1, dt2, true);
                    oSLDocument2.RenameWorksheet("Sheet1", "activacion");




                    foreach (var item1 in activacions)
                    {
                        if (item1.TIPO.Equals("NETEADO"))
                        {
                            activacion_neteados.Add(item1);
                        }
                    }

                }
                catch (Exception ex)
                {
                    
                    MessageBox.Show(ex.Message.ToString());
                    

                    throw;
                }
                


               // List<Activacion> auxiliar_pyme_2 = cargos_linq_pablo.ToList();
                //HashSet<string> rut_pyme_pablo = new HashSet<string>(activacion_neteados.Select(x => x.PCS));
                //activacions.RemoveAll(x => rut_pyme_pablo.Contains(x.PCS));                                                                                                  //VA EN RUTA

                foreach (var item in activacion_neteados)
                {
                    activacions.Remove(item);
                }


                var dt3 = new DataTable("NETEADOS");
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
                dt3.Columns.Add("prorrateo", typeof(string));
                dt3.Columns.Add("RUT", typeof(string));
                dt3.Columns.Add("TIPO", typeof(string));
                dt3.Columns.Add("COBRAR");
                dt3.Columns.Add("MONTO");
                dt3.Columns.Add("CUOTA");
                dt3.Columns.Add("OBSERVACION");
                //
                //
                foreach (var item in activacion_neteados)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }
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
                                               long.Parse(item.secuenciadeCiclo), //pim pam pum
                                               item.prorrateo,
                                               rut_con_guion,
                                               item.TIPO,
                                               "","","",""});
                    //ku ku ku ku
                }
                oSLDocument.AddWorksheet("NETEADOS");
                oSLDocument.ImportDataTable(1, 1, dt3, true);
                return true;
            }
            catch (Exception ex) 
            {
                MessageBox.Show("pasaste la segunda");
                return false;
            }
        }//TODO PERFECTO, PERO SE PUEDE AFINAR AUN MÁS
        public bool proceso_activacion_solidarios_3()
        {
            try
            {
                planSolidarios = Coleccion_PlanSolidario.GenerarListado(@path_solidario);
                var cargos_linq_solidarios = from arrendamiento in activacions
                                           join sice in planSolidarios
                                                on arrendamiento.PCS equals sice.PCS
                                           select new Activacion
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
                                               TIPO = arrendamiento.TIPO,
                                               rut = arrendamiento.rut
                                           };

                List<Activacion> auxiliar_solidarios = cargos_linq_solidarios.ToList();

                HashSet<string> rut_empresas_empresarial = new HashSet<string>(auxiliar_solidarios.Select(x => x.PCS));

                activacions.RemoveAll(x => rut_empresas_empresarial.Contains(x.PCS));

                var dt10 = new DataTable("SOLIDARIOS");
                dt10.Columns.Add("CuentaFinanciera", typeof(long));
                dt10.Columns.Add("recveivercustomer", typeof(long));
                dt10.Columns.Add("tipocargo", typeof(string));
                dt10.Columns.Add("codigodecargo", typeof(string));
                dt10.Columns.Add("offwer", typeof(long));
                dt10.Columns.Add("nombreOFFER", typeof(string));
                dt10.Columns.Add("promo", typeof(string));
                dt10.Columns.Add("description", typeof(string));
                dt10.Columns.Add("montoCargos", typeof(double));
                dt10.Columns.Add("montoDescuentos", typeof(double));
                dt10.Columns.Add("montoTotal", typeof(double));
                dt10.Columns.Add("tipodeCobro", typeof(string));
                dt10.Columns.Add("PCS", typeof(long));
                dt10.Columns.Add("secuenciadeCiclo", typeof(long));
                dt10.Columns.Add("prorrateo", typeof(string));
                dt10.Columns.Add("RUT", typeof(string));
                dt10.Columns.Add("TIPO", typeof(string));
                dt10.Columns.Add("COBRAR");
                dt10.Columns.Add("MONTO");
                dt10.Columns.Add("CUOTA");
                dt10.Columns.Add("OBSERVACION");

                foreach (var item in auxiliar_solidarios) //falta el cruce por las otras dos hojas
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }
                    dt10.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               item.TIPO,
                                               ""
                                               ,"","",""});
                }

                oSLDocument.AddWorksheet("SOLIDARIOS");
                oSLDocument.ImportDataTable(1, 1, dt10, true);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hay campos obligatorios en blanco");
                return false;
            }
        }
        public bool proceso_activacion_fvm_4()
        {
            try
            {
                fVMs = coleccion_FVM.GenerarListado(@path_fvms);
                int contador_vacios = 0;
                foreach (var item in fVMs)
                {
                    if (item.CCOCSUBEQPCOM.Equals(""))
                    {
                        contador_vacios++;
                    }
                }

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

                List<Activacion> auxiliar_activacion_fvm = activacion_fvm.ToList();

                HashSet<string> pc_activacion_fvm = new HashSet<string>(auxiliar_activacion_fvm.Select(x => x.PCS));

                activacions.RemoveAll(x => pc_activacion_fvm.Contains(x.PCS));
                //foreach (var item in activacion_fvm)
                //{
                //    activacions.Remove(item);
                //}


                List<FVM> auxiliar_fvm = cruce_con_activacion.ToList();

                HashSet<string> pc_fvm = new HashSet<string>(auxiliar_fvm.Select(x => x.PCS));

                fVMs.RemoveAll(x => pc_fvm.Contains(x.PCS));
                //foreach (var item in cruce_con_activacion)
                //{
                //    fVMs.Remove(item);
                //}

                //var miTable5 = new DataTable("FVMS");
                var dt4 = new DataTable("FVMS");
                dt4.Columns.Add("CuentaFinanciera", typeof(long));
                dt4.Columns.Add("recveivercustomer", typeof(long));
                dt4.Columns.Add("tipocargo", typeof(string));
                dt4.Columns.Add("codigodecargo", typeof(string));
                dt4.Columns.Add("offwer", typeof(long));
                dt4.Columns.Add("nombreOFFER", typeof(string));
                dt4.Columns.Add("promo", typeof(string));
                dt4.Columns.Add("description", typeof(string));
                dt4.Columns.Add("montoCargos", typeof(double));
                dt4.Columns.Add("montoDescuentos", typeof(double));
                dt4.Columns.Add("montoTotal", typeof(double));
                dt4.Columns.Add("tipodeCobro", typeof(string));
                dt4.Columns.Add("PCS", typeof(long));
                dt4.Columns.Add("secuenciadeCiclo", typeof(long));
                dt4.Columns.Add("prorrateo", typeof(string));
                dt4.Columns.Add("RUT", typeof(string));
                dt4.Columns.Add("TIPO", typeof(string));
                dt4.Columns.Add("COBRAR",typeof(double));
                dt4.Columns.Add("MONTO");
                dt4.Columns.Add("CUOTA");
                dt4.Columns.Add("OBSERVACION");

                foreach (var item in activacion_fvm)
                {
                    int contador_para_guion = item.rut.Length - 1;
                    string rut_con_guion = item.rut.Insert(contador_para_guion, "-");
                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }
                    dt4.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               item.TIPO,
                                               item.COBRAR
                                               ,"","",""});
                }

                oSLDocument.AddWorksheet("FVM");
                oSLDocument.ImportDataTable(1, 1, dt4, true); //porque mierda no agrega esta hoja


                var linq_fvmxfvmsactivacion = from fvm in activacion_fvm
                                              join neteados in activacion_fvm
                                             on fvm.PCS equals neteados.PCS
                                        select new FVM
                                        {
                                            CCOCSUBEQPCOM = string.Empty,
                                            MONTO_SIN_IMPUESTO = 0.0,
                                            PCS = string.Empty,
                                            INGRESO_VTA = string.Empty,
                                            FA = string.Empty,
                                            CBP = string.Empty,
                                            FECHA_ACTIVACION = string.Empty
                                        };
                var dt5 = new DataTable("FVM");

                dt5.Columns.Add("CCOCSUBEQPCOM", typeof(string));
                dt5.Columns.Add("PCS", typeof(long));
                dt5.Columns.Add("MONTO_SIN_IMPUESTO", typeof(double));
                dt5.Columns.Add("INGRESO_VTA", typeof(string));
                dt5.Columns.Add("FA", typeof(long));
                dt5.Columns.Add("CBP", typeof(long));
                dt5.Columns.Add("FECHA_ACTIVACION", typeof(string));
                int contador_nocruza = linq_fvmxfvmsactivacion.Count();

                foreach (var item in fVMs)
                {
                    //DateTime oDate = DateTime.Parse(item.FECHA_ACTIVACION);
                    //string fecha = oDate.Day + "-" + oDate.Month + "-" + oDate.Year;
                    if (item.CCOCSUBEQPCOM.Equals(""))
                    {

                    }
                    else
                    {
                        dt5.Rows.Add(new Object[] { item.CCOCSUBEQPCOM,
                                                long.Parse(item.PCS),
                                                Math.Round(item.MONTO_SIN_IMPUESTO),
                                                item.INGRESO_VTA,
                                                long.Parse(item.FA),
                                                long.Parse(item.CBP),
                                                item.FECHA_ACTIVACION});
                    }

                }

                oSLDocument3.ImportDataTable(1, 1, dt5, true);
                oSLDocument3.RenameWorksheet("Sheet1", "FVM");
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Aqui es donde se pudre todo");
                return false;
            }
        }
        public bool proceso_activacion_habilitaciones_5()
        {
            try
            {
                hABILITACIONEs = coleccion_HABILITACIONES.GenerarListado(@path_habili);
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


                List<Activacion> auxiliar_habilitaciones_fvm = activacion_habilitaciones.ToList();

                HashSet<string> pc_activacion_habilitaciones = new HashSet<string>(auxiliar_habilitaciones_fvm.Select(x => x.PCS));

                activacions.RemoveAll(x => pc_activacion_habilitaciones.Contains(x.PCS));
                //foreach (var item in activacion_habilitaciones)
                //{
                //    activacions.Remove(item);
                //}

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


                var dt5 = new DataTable("HABILITACIONES");
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
                dt5.Columns.Add("COBRAR");
                dt5.Columns.Add("MONTO", typeof(double));
                dt5.Columns.Add("CUOTA",typeof(string));
                dt5.Columns.Add("OBSERVACION");

                
                foreach (var item in activacion_habilitaciones)
                {
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
                                               item.TIPO,
                                               item.COBRAR_SI
                                               ,item.COBRAR
                                               ,item.CUOTA,
                                                ""});
                    
                }

                

                oSLDocument.AddWorksheet("HABILITACIONES");
                oSLDocument.ImportDataTable(1, 1, dt5, true);
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Hay un error con las habilitaciones, si pinchas continuar, el proceso se completará normalmente");
                
                return false;
            }
        }

        public bool proceso_activacion_duplicados_6() //aquie sta lo complicado
        {
            try
            {
                List<Activacion> activacion_duplicados = new List<Activacion>();
                int repetidos = 0;
                foreach (var item1 in activacions)
                {
                    repetidos = 0;
                    foreach (var item2 in activacions)
                    {
                        if (item1.PCS.Equals(item2.PCS))
                        {
                            repetidos++;
                        }
                    }
                    if (repetidos >= 2)
                    {
                        item1.duplicado = "ES DULPICADO";
                        activacion_duplicados.Add(item1); //esta bien pero puede ser mejor
                    }
                }


                foreach (var item in activacion_duplicados)//HASTA AQUI TODO CORRECTO Y YO QUE ME ALEGRO
                {
                    activacions.Remove(item);
                }
                int contador_verdaderos = 0;
                int contador_falsos = 0;
                foreach (var item in activacion_duplicados)
                {
                    if (item.montoTotal > 0)
                    {
                        item.verdaredo_falso = "VERDADERO";
                        contador_verdaderos++;
                    }
                    else
                    {
                        item.verdaredo_falso = "FALSO";
                        contador_falsos++;
                    }
                }

                List<Activacion> activacion_duplicados_verdaderos_repetidos = new List<Activacion>();

                foreach (var item1 in activacion_duplicados)
                {
                    repetidos = 0;
                    foreach (var item2 in activacion_duplicados)
                    {
                        if (item1.PCS.Equals(item2.PCS) && item1.verdaredo_falso.Equals("VERDADERO") && item1.verdaredo_falso.Equals(item2.verdaredo_falso))
                        {
                            repetidos++;
                        }
                    }
                    if (repetidos >= 2)
                    {
                        item1.duplicado = "ES DULPICADO";//
                        
                        activacion_duplicados_verdaderos_repetidos.Add(item1); //esta bien pero puede ser mejor, creo que lo mejore pero no me acuerdo
                    }
                }

                foreach (var item in activacion_duplicados_verdaderos_repetidos)
                {
                    activacion_duplicados.Remove(item);
                }

                List<Activacion> activacion_duplicados_verdaderos_falsos = new List<Activacion>();
                foreach (var item1 in activacion_duplicados)
                {
                    foreach (var item2 in activacion_duplicados_verdaderos_repetidos)
                    {
                        if (item1.PCS.Equals(item2.PCS))
                        {
                            activacion_duplicados_verdaderos_falsos.Add(item1);
                            break;
                        }
                    }
                }

                foreach (var item in activacion_duplicados_verdaderos_falsos)
                {
                    activacion_duplicados_verdaderos_repetidos.Add(item);
                }
                Console.WriteLine("aqui me detengo");

                foreach (var item in activacions)
                {
                    if (item.montoTotal > monto_filtrado)
                    {
                        item.verdaredo_falso = "";
                        activacion_duplicados_verdaderos_repetidos.Add(item);
                    }
                }

                foreach (var item in activacion_duplicados_verdaderos_repetidos)
                {
                    activacion_duplicados.Remove(item);
                }

                

                foreach (var item in activacion_duplicados_verdaderos_repetidos)
                {
                    if (item.TIPO.Equals("PORTADO"))
                    {
                        activacion_portados.Add(item);
                    }
                }

                var dt7 = new DataTable("DUPLICADOS");
                dt7.Columns.Add("CuentaFinanciera", typeof(long));
                dt7.Columns.Add("recveivercustomer", typeof(long));
                dt7.Columns.Add("tipocargo", typeof(string));
                dt7.Columns.Add("codigodecargo", typeof(string));
                dt7.Columns.Add("offwer", typeof(long));
                dt7.Columns.Add("nombreOFFER", typeof(string));
                dt7.Columns.Add("promo", typeof(string));
                dt7.Columns.Add("description", typeof(string));
                dt7.Columns.Add("montoCargos", typeof(double));
                dt7.Columns.Add("montoDescuentos", typeof(double));
                dt7.Columns.Add("montoTotal", typeof(double));
                dt7.Columns.Add("tipodeCobro", typeof(string));
                dt7.Columns.Add("PCS", typeof(long));
                dt7.Columns.Add("secuenciadeCiclo", typeof(long));
                dt7.Columns.Add("TIPO", typeof(string));
                dt7.Columns.Add("COBRAR");
                dt7.Columns.Add("MONTO");
                dt7.Columns.Add("CUOTA");
                dt7.Columns.Add("OBSERVACION");


                List<Activacion> auxiliar_portados_activacion = activacion_portados.ToList();

                HashSet<string> pc_portados_activacion = new HashSet<string>(auxiliar_portados_activacion.Select(x => x.PCS));

                activacions.RemoveAll(x => pc_portados_activacion.Contains(x.PCS));


                foreach (var item in activacion_portados)
                {
                    dt7.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               item.TIPO,
                                               ""
                                               ,""
                                               ,"",
                                                ""});
                    activacions_asignados_portados.Add(item);
                }
                oSLDocument5.ImportDataTable(1, 1, dt7, true);
                oSLDocument5.RenameWorksheet("Sheet1", "PORTADOS");

                foreach (var item in activacion_portados)
                {
                    activacion_duplicados_verdaderos_repetidos.Remove(item);
                }

                var dt8 = new DataTable("CARGOS DE ACTIVACION");
                dt8.Columns.Add("ASIGNACION"); //esto se cruza con lo nuevo pom pom pooooooooooom
                dt8.Columns.Add("CuentaFinanciera", typeof(long));
                dt8.Columns.Add("recveivercustomer", typeof(long));
                dt8.Columns.Add("tipocargo", typeof(string));
                dt8.Columns.Add("codigodecargo", typeof(string));
                dt8.Columns.Add("offwer", typeof(long));
                dt8.Columns.Add("nombreOFFER", typeof(string));
                dt8.Columns.Add("promo", typeof(string));
                dt8.Columns.Add("description", typeof(string));
                dt8.Columns.Add("montoCargos", typeof(double));
                dt8.Columns.Add("montoDescuentos", typeof(double));
                dt8.Columns.Add("montoTotal", typeof(double));
                dt8.Columns.Add("VERDADERO O FALSO", typeof(string));
                dt8.Columns.Add("DUPLICADO", typeof(string));
                dt8.Columns.Add("tipodeCobro", typeof(string));
                dt8.Columns.Add("PCS", typeof(long));
                dt8.Columns.Add("secuenciadeCiclo", typeof(long));
                dt8.Columns.Add("TIPO", typeof(string));
                dt8.Columns.Add("COBRAR");
                dt8.Columns.Add("MONTO");
                dt8.Columns.Add("CUOTA");
                dt8.Columns.Add("OBSERVACION");

                List<Activacion> auxiliar_verdaderos = activacion_duplicados_verdaderos_repetidos.ToList();

                HashSet<string> pc_verdaderos = new HashSet<string>(activacion_duplicados_verdaderos_repetidos.Select(x => x.PCS));

                activacions.RemoveAll(x => pc_verdaderos.Contains(x.PCS));

                //--------------------------------------------------------------------------------------------------------------

                //var auxiliaron_asignados = aux_Activacion.GenerarListado_asignados(@aux_path);--------------------------------

                //--------------------------------------------------------------------------------------------------------------
                //esto tomo para lo de ahora
                foreach (var item in activacion_duplicados_verdaderos_repetidos.OrderBy(i => i.duplicado))
                {
                    dt8.Rows.Add(new Object[] { "",
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
                                               item.verdaredo_falso,
                                               item.duplicado,
                                               item.tipodeCobro,
                                               long.Parse(item.PCS),
                                               long.Parse(item.secuenciadeCiclo),
                                               item.TIPO,
                                               ""
                                               ,""
                                               ,"",
                                                ""});

                    activacions_asignados.Add(item);
                }
                //
                //------------------------------------------------------------------------------------------------------------------
                oSLDocument6.ImportDataTable(1, 1, dt8, true);
                oSLDocument6.RenameWorksheet("Sheet1", "CARGOS DE ACTIVACION");//estos son los datos que despues se utilizaran

                //foreach (var item in activacion_duplicados)
                //{
                //    if (item.TIPO.Equals("PORTADO"))
                //    {
                //        activacion_portados.Add(item);
                //    }
                //}//eureka! datatable
                //

                var dt6 = new DataTable("DUPLICADOS");
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
                dt6.Columns.Add("COBRAR");
                dt6.Columns.Add("MONTO");
                dt6.Columns.Add("CUOTA");
                dt6.Columns.Add("OBSERVACION");


                foreach (var item in activacion_duplicados)
                {
                    dt6.Rows.Add(new Object[] { long.Parse(item.CuentaFinanciera),
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
                                               item.TIPO,
                                               ""
                                               ,""
                                               ,"",
                                                ""});
                }
                ////esto va al nuevo excel
                oSLDocument.AddWorksheet("DUPLICADOS");
                oSLDocument.ImportDataTable(1, 1, dt6, true);
                
                //////SaveExcel.BuildExcel(miTable17, @quedaran_aqui_txt.Text + @"\Ciclo" + @"\Cargos de Activacion.xlsx");// este es el excel nuevo





                // SaveExcel.BuildExcel(miTable16, @quedaran_aqui_txt.Text + @"\Activacion CARGOS DE ACTIVACION.xlsx");

                var dt9 = new DataTable("ORIGINAL");
                dt9.Columns.Add("CuentaFinanciera", typeof(long));
                dt9.Columns.Add("recveivercustomer", typeof(long));
                dt9.Columns.Add("tipocargo", typeof(string));
                dt9.Columns.Add("codigodecargo", typeof(string));
                dt9.Columns.Add("offwer", typeof(long));
                dt9.Columns.Add("nombreOFFER", typeof(string));
                dt9.Columns.Add("promo", typeof(string));
                dt9.Columns.Add("description", typeof(string));
                dt9.Columns.Add("montoCargos", typeof(double));
                dt9.Columns.Add("montoDescuentos", typeof(double));
                dt9.Columns.Add("montoTotal", typeof(double));
                dt9.Columns.Add("tipodeCobro", typeof(string));
                dt9.Columns.Add("PCS", typeof(long));
                dt9.Columns.Add("secuenciadeCiclo", typeof(long));
                dt9.Columns.Add("TIPO", typeof(string));
                dt9.Columns.Add("COBRAR");
                dt9.Columns.Add("MONTO");
                dt9.Columns.Add("CUOTA");
                dt9.Columns.Add("OBSERVACION");

                //...
                foreach (var item in activacions)
                {
                    dt9.Rows.Add(new Object[] {
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
                                               item.montoTotal,
                                               item.tipodeCobro,
                                               long.Parse(item.PCS),
                                               long.Parse(item.secuenciadeCiclo),//a_a e_e i_i o_o u_u w_w z_z x_x c_c v_v b_b n_n m_m s_s d_d
                                               item.TIPO,
                                               ""
                                               ,""
                                               ,"",
                                                ""});
                }

                oSLDocument.AddWorksheet("SIN ASIGNAR"); 
                oSLDocument.ImportDataTable(1, 1, dt9, true);
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }
        //---.
        public bool proceso_consolidacion_6_5()//
        {
            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            int contador_ = 0;
            try
            {//
                Coleccion_Consolidacion_Validacion coleccion_conso = new Coleccion_Consolidacion_Validacion();
                consolidacion_Validacions = coleccion_conso.GenerarListado(@path_conso);

                List<Consolidacion_Validacion> auxiliar_consolidado_ajuste_0 = new List<Consolidacion_Validacion>();
                List<Consolidacion_Validacion> auxiliar_consolidado_ajuste = new List<Consolidacion_Validacion>();
                
                List<Activacion> activacion_revisados    = new List<Activacion>();
                List<Activacion> activacion_no_revisados = new List<Activacion>();

                List<Activacion> activacion_duplex = new List<Activacion>();

                List<string> pcs_ajuste_conso = new List<string>();

                foreach (var item in consolidacion_Validacions)
                {
                    contador_++;
                    if (item.pcs.Equals(""))
                    {

                    }
                    else
                    {
                        double ajuste = 0.0;
                        int tiene_punto = item.monto_ajuste.IndexOf('.');
                        if (tiene_punto > 1)
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_punto);

                        }
                        else
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_coma);
                        }
                        if (ajuste == 0)
                        {

                        }
                        else
                        {
                            pcs_ajuste_conso.Add(item.pcs);
                            item.hay_ajuste = true;
                        }
                    }


                }

                foreach (var item1 in pcs_ajuste_conso.Distinct())
                {
                    foreach (var item2 in consolidacion_Validacions)
                    {
                        if (item1.Equals(item2.pcs))
                        {
                            item2.hay_ajuste = true;
                        }
                    }
                }



                //activacions_asignados esta es la lista a ocupar

                var cargos_linq_validacion = from asignados in activacions_asignados
                                             join validacion in consolidacion_Validacions
                                                  on asignados.PCS equals validacion.pcs
                                             select new Activacion
                                             {
                                                 CuentaFinanciera = asignados.CuentaFinanciera,
                                                 recveivercustomer = asignados.recveivercustomer,
                                                 tipocargo = asignados.tipocargo,
                                                 codigodecargo = asignados.codigodecargo,
                                                 offwer = asignados.offwer,
                                                 nombreOFFER = asignados.nombreOFFER,
                                                 promo = asignados.promo,
                                                 description = asignados.description,
                                                 montoCargos = asignados.montoCargos,
                                                 montoDescuentos = asignados.montoDescuentos,
                                                 montoTotal = asignados.montoTotal,
                                                 tipodeCobro = asignados.tipodeCobro,
                                                 PCS = asignados.PCS,
                                                 secuenciadeCiclo = asignados.secuenciadeCiclo,
                                                 prorrateo = asignados.prorrateo,
                                                 TIPO = asignados.TIPO,
                                                 rut = asignados.rut,
                                                 verdaredo_falso = asignados.verdaredo_falso,
                                                 duplicado = asignados.duplicado,
                                                 monto_consolidado = validacion.monto_validador,
                                                 monto_ajuste = validacion.monto_ajuste,
                                                 hay_ajuste = validacion.hay_ajuste

                                             };

                int contadorinio = 0;

                foreach (var item in consolidacion_Validacions)
                {
                    double ajuste = 0.0;
                    int tiene_punto = item.monto_ajuste.IndexOf('.');
                    if (item.monto_ajuste.Equals(""))
                    {

                    }
                    else
                    {
                        if (item.hay_ajuste)
                        {
                            auxiliar_consolidado_ajuste.Add(item);
                        }
                        else
                        {
                            auxiliar_consolidado_ajuste_0.Add(item);
                        }
                    }
                }
                


                var dt9 = new DataTable("CARGOS DE ACTIVACION");
                dt9.Columns.Add("ASIGNACION"); 
                dt9.Columns.Add("CuentaFinanciera", typeof(long));
                dt9.Columns.Add("recveivercustomer", typeof(long));
                dt9.Columns.Add("tipocargo", typeof(string));
                dt9.Columns.Add("codigodecargo", typeof(string));
                dt9.Columns.Add("offwer", typeof(long));
                dt9.Columns.Add("nombreOFFER", typeof(string));
                dt9.Columns.Add("promo", typeof(string));
                dt9.Columns.Add("description", typeof(string));
                dt9.Columns.Add("montoCargos", typeof(double));
                dt9.Columns.Add("montoDescuentos", typeof(double));
                dt9.Columns.Add("montoTotal", typeof(double));
                dt9.Columns.Add("tipodeCobro", typeof(string));
                dt9.Columns.Add("PCS", typeof(long));
                dt9.Columns.Add("secuenciadeCiclo", typeof(long));
                dt9.Columns.Add("TIPO", typeof(string));
                dt9.Columns.Add("COBRAR");
                dt9.Columns.Add("MONTO");
                dt9.Columns.Add("CUOTA");
                dt9.Columns.Add("OBSERVACION");


                var dt10 = new DataTable("CARGOS DE ACTIVACION");
                dt10.Columns.Add("ASIGNACION"); 
                dt10.Columns.Add("CuentaFinanciera", typeof(long));
                dt10.Columns.Add("recveivercustomer", typeof(long));
                dt10.Columns.Add("tipocargo", typeof(string));
                dt10.Columns.Add("codigodecargo", typeof(string));
                dt10.Columns.Add("offwer", typeof(long));
                dt10.Columns.Add("nombreOFFER", typeof(string));
                dt10.Columns.Add("promo", typeof(string));
                dt10.Columns.Add("description", typeof(string));
                dt10.Columns.Add("montoCargos", typeof(double));
                dt10.Columns.Add("montoDescuentos", typeof(double));
                dt10.Columns.Add("montoTotal", typeof(double));
                dt10.Columns.Add("VERDADERO O FALSO", typeof(string));
                dt10.Columns.Add("DUPLICADO", typeof(string));
                dt10.Columns.Add("tipodeCobro", typeof(string));
                dt10.Columns.Add("PCS", typeof(long));
                dt10.Columns.Add("secuenciadeCiclo", typeof(long));
                dt10.Columns.Add("TIPO", typeof(string));
                dt10.Columns.Add("COBRAR");
                dt10.Columns.Add("MONTO", typeof(double));
                dt10.Columns.Add("CUOTA");
                dt10.Columns.Add("OBSERVACION");


                
                int cuantos_me_dan = cargos_linq_validacion.Count();
                //int contador_nuevo = 0;

                foreach (var item in cargos_linq_validacion)
                {
                    double monto = 0.0;
                    double ajuste = 0.0;
                    if (item.monto_consolidado.Equals(""))
                    {

                    }
                    else
                    {
                        int tiene_punto = item.monto_consolidado.IndexOf('.');
                        if (tiene_punto > 1)
                        {
                            monto = double.Parse(item.monto_consolidado, culture_punto);

                        }
                        else
                        {
                            monto = double.Parse(item.monto_consolidado, culture_coma);
                        }
                        //monto = double.Parse(item.monto_consolidado);
                    }
                    if (item.monto_ajuste.Equals(""))
                    {

                    }
                    else
                    {
                        int tiene_punto = item.monto_ajuste.IndexOf('.');
                        if (tiene_punto > 1)
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_punto);

                        }
                        else
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_coma);
                        }
                        //monto = double.Parse(item.monto_consolidado);
                    }

                    if (item.duplicado.Equals("")) 
                    {
                        if (item.hay_ajuste)
                        {

                        }
                        else
                        {
                            contadorinio++;
                            if (monto == item.montoTotal)
                            {
                                activacion_revisados.Add(item);

                            }
                        }
 
                    }
                    else
                    {
                        activacion_duplex.Add(item);
                    }
                }

                List<int> donde_esta = new List<int>();
                //aqui hay que volver a cruzar
                var cargos_linq_validacion_2 = from asignados in activacion_revisados
                                               join validacion in auxiliar_consolidado_ajuste
                                                  on asignados.PCS equals validacion.pcs
                                             select new Activacion
                                             {
                                                 CuentaFinanciera = asignados.CuentaFinanciera,
                                                 recveivercustomer = asignados.recveivercustomer,
                                                 tipocargo = asignados.tipocargo,
                                                 codigodecargo = asignados.codigodecargo,
                                                 offwer = asignados.offwer,
                                                 nombreOFFER = asignados.nombreOFFER,
                                                 promo = asignados.promo,
                                                 description = asignados.description,
                                                 montoCargos = asignados.montoCargos,
                                                 montoDescuentos = asignados.montoDescuentos,
                                                 montoTotal = asignados.montoTotal,
                                                 tipodeCobro = asignados.tipodeCobro,
                                                 PCS = asignados.PCS,
                                                 secuenciadeCiclo = asignados.secuenciadeCiclo,
                                                 prorrateo = asignados.prorrateo,
                                                 TIPO = asignados.TIPO,
                                                 rut = asignados.rut,
                                                 verdaredo_falso = asignados.verdaredo_falso,
                                                 duplicado = asignados.duplicado,
                                                 monto_consolidado = validacion.monto_validador,
                                                 monto_ajuste = validacion.monto_ajuste

                                             };
                 
                cuantos_me_dan = cargos_linq_validacion_2.Count();//con esto ya capturamos los 2 sobrantes, entonces ahora se debe eliminar estos 2 pcs de activacion revisados

                for (int i = 0; i < activacion_revisados.Count()-1; i++)
                {

                    bool se_encontro = false;
                    foreach (var item in cargos_linq_validacion_2)
                    {
                        if (activacion_revisados[i].PCS.Equals(item.PCS))
                        {
                            se_encontro = true;
                            break;
                        }
                    }

                    if (se_encontro)
                    {
                        donde_esta.Add(i);
                    }
                    
                }

                foreach (var item in donde_esta)
                {
                    activacion_no_revisados.Add(activacion_revisados[item]);
                    activacion_revisados.RemoveAt(item); //esto no se deberia hacer asi, esta malo
                }

                foreach (var item in activacion_revisados)
                {
                    activacions_revisados_asignados.Add(item);
                    //dt9.Rows.Add(new Object[] { "",
                    //                   long.Parse(item.CuentaFinanciera),
                    //                   long.Parse(item.recveivercustomer),
                    //                   item.tipocargo,
                    //                   item.codigodecargo,
                    //                   long.Parse(item.offwer),
                    //                   item.nombreOFFER,
                    //                   item.promo,
                    //                   item.description,
                    //                   item.montoCargos,
                    //                   item.montoDescuentos,
                    //                   item.montoTotal ,
                    //                   item.tipodeCobro,
                    //                   long.Parse(item.PCS),
                    //                   long.Parse(item.secuenciadeCiclo),
                    //                   item.TIPO,
                    //                   ""
                    //                   ,""
                    //                   ,"",
                    //                    "REVISADO"});
                } //aqui quedan todos los revisados
                //oSLDocument4.AddWorksheet("REVISADO"); //aqui vamos a crear el primer foreach para crear la hoja revisados
                //oSLDocument4.ImportDataTable(1, 1, dt9, true);//esto debe moverse a el respaldo

                List<Activacion> auxiliar_asignados = activacions_asignados;
                HashSet<string> pcs_activacion_revisados = new HashSet<string>(activacion_revisados.Select(x => x.PCS));

                auxiliar_asignados.RemoveAll(x => pcs_activacion_revisados.Contains(x.PCS));

                List<Activacion> auxiliar_verdadero = new List<Activacion>();
                List<Activacion> auxiliar_falso = new List<Activacion>();

                foreach (var item in activacions_asignados)
                {
                    if (item.verdaredo_falso.Equals("VERDADERO"))
                    {
                        auxiliar_verdadero.Add(item);
                    }
                    else if(item.verdaredo_falso.Equals("FALSO"))
                    {
                        auxiliar_falso.Add(item);
                    }
                } // hacer la misma separacion que hize en el excel pero como.... separar los que solo son verdaderos de los que oslo son falsos

                List<int> donde_verdadero = new List<int>();

                for (int i = 0; i < auxiliar_verdadero.Count() - 1; i++)
                {
                    bool se_encontro = false;
                    foreach (var item in auxiliar_falso)
                    {
                        if (auxiliar_verdadero[i].PCS.Equals(item.PCS))
                        {
                            se_encontro = true;
                            break;
                        }
                    }

                    if (se_encontro)
                    {
                        donde_verdadero.Add(i);
                    }

                }//esta todo correcto hasta aqui marcare--------------------------------------------------------------------------

                foreach (var item in donde_verdadero)
                {
                    auxiliar_falso.Add(auxiliar_verdadero[item]);
                    //auxiliar_verdadero.RemoveAt(item); esto esta malo no descomentar
                }

                foreach (var item in auxiliar_falso)//
                {
                    auxiliar_verdadero.Remove(item);
                }

                //esta un poco molesto por no saber, 

                foreach (var item in auxiliar_verdadero) //esto se agregara algun dia 
                {
                    //dt10.Rows.Add(new Object[] { "",
                    //                   long.Parse(item.CuentaFinanciera),
                    //                   long.Parse(item.recveivercustomer),
                    //                   item.tipocargo,
                    //                   item.codigodecargo,
                    //                   long.Parse(item.offwer),
                    //                   item.nombreOFFER,
                    //                   item.promo,
                    //                   item.description,
                    //                   item.montoCargos,
                    //                   item.montoDescuentos,
                    //                   item.montoTotal ,
                    //                   item.verdaredo_falso,
                    //                   item.duplicado,
                    //                   item.tipodeCobro,
                    //                   long.Parse(item.PCS),
                    //                   long.Parse(item.secuenciadeCiclo),
                    //                   item.TIPO,
                    //                   ""
                    //                   ,""
                    //                   ,"",
                    //                    ""});
                }

                //
                //oSLDocument4.AddWorksheet("DUPLICADO"); //aqui vamos a crear el segundo foreach para crear la hoja duplicado
                //oSLDocument4.ImportDataTable(1, 1, dt10, true); //say something or i give in up on u sorry that i tarara rara ra ra

                var dt11 = new DataTable("CARGOS DE ACTIVACION");
                dt11.Columns.Add("ASIGNACION");
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
                dt11.Columns.Add("VERDADERO O FALSO", typeof(string));
                dt11.Columns.Add("DUPLICADO", typeof(string));
                dt11.Columns.Add("tipodeCobro", typeof(string));
                dt11.Columns.Add("PCS", typeof(long));
                dt11.Columns.Add("secuenciadeCiclo", typeof(long));
                dt11.Columns.Add("TIPO", typeof(string));
                dt11.Columns.Add("COBRAR");
                dt11.Columns.Add("MONTO", typeof(double));
                dt11.Columns.Add("CUOTA");
                dt11.Columns.Add("OBSERVACION");

                foreach (var item in activacions_asignados)
                {
                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }


                    dt11.Rows.Add(new Object[] { "",
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
                                       item.verdaredo_falso,
                                       item.duplicado,
                                       item.tipodeCobro,
                                       long.Parse(item.PCS),
                                       long.Parse(item.secuenciadeCiclo),
                                       item.TIPO,
                                       ""
                                       ,item.montoTotal
                                       ,"",
                                        ""});
                }


                oSLDocument4.AddWorksheet("NO REVISADOS"); //aqui vamos a crear el segundo foreach para crear la hoja duplicado
                oSLDocument4.ImportDataTable(1, 1, dt11, true);

                

                

                cuantos_me_dan = cargos_linq_validacion.Count();
                //PRIMERO TOMAMOS LOS PCS DE ASIGNACION Y LOS CRUZAMOS CON CONSOLIDADO TODOS LOS QUE CRUZEN SE DEBE COMPROBAR SI LOS MONTOS SON IGUALES Y NO SON DUPLICADOS. ESOS SON LOS PRIMEROS PASOS DE LOS PASOS.
                //PASO UNO CAPTURAR EL EXCEL DE CONSOLIDADO(SUS PCS Y MONTO)
                //PASO DOS COMPARAR PCS QUE SE ASIGNARAN CON CONSOLIDADO
                //PASO TRES TODO LO QUE CRUZA SE LE DEBERA COMPARAR EL MONTO
                //PASO CUATRO SI SON IGUALES LOS MONTOS SE DEBE VERIFICAR SI NO SON DUPLICADOS
                //Y AHI ESTAN LOS REVISADOS DE LA HOJA
                //CODIGO CCOCSUBEQPCOMM, ESTE CODIGO ES LO QUE HACE ALUSION AL CARGO DE ACTIVACION
                //
                //primero revisar ajuste y despues comparar montos.

                //------------------------------------------------------------------------------------------
                //DESPUES EMPEZAMOS A SACAR LOS DUPLICADOS
                //LOS DE E COMERCE NO SE DE DONDE SE SACAN Y AJUSTADOS TAMPOCO, PERO ES COSA DE IR DESGLOZANDO COMO LO HICE AHORA C:
                //agregar ajustes adicionales


                return true;

            }
            catch (Exception ex)
            {

                return false;
                
            }
        }
        public bool proceso_consolidacion_6_6_portados()//
        {
            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            int contador_ = 0;
            try
            {//
                Coleccion_Consolidacion_Validacion coleccion_conso = new Coleccion_Consolidacion_Validacion();
                consolidacion_Validacions = coleccion_conso.GenerarListado(@path_conso);

                List<Consolidacion_Validacion> auxiliar_consolidado_ajuste_0 = new List<Consolidacion_Validacion>();
                List<Consolidacion_Validacion> auxiliar_consolidado_ajuste = new List<Consolidacion_Validacion>();

                List<Activacion> activacion_revisados = new List<Activacion>();
                List<Activacion> activacion_no_revisados = new List<Activacion>();

                List<Activacion> activacion_duplex = new List<Activacion>();

                List<string> pcs_ajuste_conso = new List<string>();

                foreach (var item in consolidacion_Validacions)
                {
                    contador_++;
                    if (item.pcs.Equals(""))
                    {

                    }
                    else
                    {
                        double ajuste = 0.0;
                        int tiene_punto = item.monto_ajuste.IndexOf('.');
                        if (tiene_punto > 1)
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_punto);

                        }
                        else
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_coma);
                        }
                        if (ajuste == 0)
                        {

                        }
                        else
                        {
                            pcs_ajuste_conso.Add(item.pcs);
                            item.hay_ajuste = true;
                        }
                    }


                }

                foreach (var item1 in pcs_ajuste_conso.Distinct())
                {
                    foreach (var item2 in consolidacion_Validacions)
                    {
                        if (item1.Equals(item2.pcs))
                        {
                            item2.hay_ajuste = true;
                        }
                    }
                }

                //activacions_asignados esta es la lista a ocupar

                
                var cargos_linq_validacion = from asignados in activacion_portados
                                             join validacion in consolidacion_Validacions
                                                  on asignados.PCS equals validacion.pcs
                                             select new Activacion
                                             {
                                                 CuentaFinanciera = asignados.CuentaFinanciera,
                                                 recveivercustomer = asignados.recveivercustomer,
                                                 tipocargo = asignados.tipocargo,
                                                 codigodecargo = asignados.codigodecargo,
                                                 offwer = asignados.offwer,
                                                 nombreOFFER = asignados.nombreOFFER,
                                                 promo = asignados.promo,
                                                 description = asignados.description,
                                                 montoCargos = asignados.montoCargos,
                                                 montoDescuentos = asignados.montoDescuentos,
                                                 montoTotal = asignados.montoTotal,
                                                 tipodeCobro = asignados.tipodeCobro,
                                                 PCS = asignados.PCS,
                                                 secuenciadeCiclo = asignados.secuenciadeCiclo,
                                                 prorrateo = asignados.prorrateo,
                                                 TIPO = asignados.TIPO,
                                                 rut = asignados.rut,
                                                 verdaredo_falso = asignados.verdaredo_falso,
                                                 duplicado = asignados.duplicado,
                                                 monto_consolidado = validacion.monto_validador,
                                                 monto_ajuste = validacion.monto_ajuste,
                                                 hay_ajuste = validacion.hay_ajuste

                                             };

                int contadorinio = 0;

                foreach (var item in consolidacion_Validacions)
                {
                    double ajuste = 0.0;
                    int tiene_punto = item.monto_ajuste.IndexOf('.');
                    if (item.monto_ajuste.Equals(""))
                    {

                    }
                    else
                    {
                        if (item.hay_ajuste)
                        {
                            auxiliar_consolidado_ajuste.Add(item);
                        }
                        else
                        {
                            auxiliar_consolidado_ajuste_0.Add(item);
                        }
                    }
                }



                var dt9 = new DataTable("CARGOS DE ACTIVACION");
                dt9.Columns.Add("ASIGNACION");
                dt9.Columns.Add("CuentaFinanciera", typeof(long));
                dt9.Columns.Add("recveivercustomer", typeof(long));
                dt9.Columns.Add("tipocargo", typeof(string));
                dt9.Columns.Add("codigodecargo", typeof(string));
                dt9.Columns.Add("offwer", typeof(long));
                dt9.Columns.Add("nombreOFFER", typeof(string));
                dt9.Columns.Add("promo", typeof(string));
                dt9.Columns.Add("description", typeof(string));
                dt9.Columns.Add("montoCargos", typeof(double));
                dt9.Columns.Add("montoDescuentos", typeof(double));
                dt9.Columns.Add("montoTotal", typeof(double));
                dt9.Columns.Add("tipodeCobro", typeof(string));
                dt9.Columns.Add("PCS", typeof(long));
                dt9.Columns.Add("secuenciadeCiclo", typeof(long));
                dt9.Columns.Add("TIPO", typeof(string));
                dt9.Columns.Add("COBRAR");
                dt9.Columns.Add("MONTO", typeof(double));
                dt9.Columns.Add("CUOTA");
                dt9.Columns.Add("OBSERVACION");


                var dt10 = new DataTable("CARGOS DE ACTIVACION");
                dt10.Columns.Add("ASIGNACION");
                dt10.Columns.Add("CuentaFinanciera", typeof(long));
                dt10.Columns.Add("recveivercustomer", typeof(long));
                dt10.Columns.Add("tipocargo", typeof(string));
                dt10.Columns.Add("codigodecargo", typeof(string));
                dt10.Columns.Add("offwer", typeof(long));
                dt10.Columns.Add("nombreOFFER", typeof(string));
                dt10.Columns.Add("promo", typeof(string));
                dt10.Columns.Add("description", typeof(string));
                dt10.Columns.Add("montoCargos", typeof(double));
                dt10.Columns.Add("montoDescuentos", typeof(double));
                dt10.Columns.Add("montoTotal", typeof(double));
                dt10.Columns.Add("VERDADERO O FALSO", typeof(string));
                dt10.Columns.Add("DUPLICADO", typeof(string));
                dt10.Columns.Add("tipodeCobro", typeof(string));
                dt10.Columns.Add("PCS", typeof(long));
                dt10.Columns.Add("secuenciadeCiclo", typeof(long));
                dt10.Columns.Add("TIPO", typeof(string));
                dt10.Columns.Add("COBRAR");
                dt10.Columns.Add("MONTO", typeof(double));
                dt10.Columns.Add("CUOTA");
                dt10.Columns.Add("OBSERVACION");



                int cuantos_me_dan = cargos_linq_validacion.Count();
                //int contador_nuevo = 0;

                foreach (var item in cargos_linq_validacion)
                {
                    double monto = 0.0;
                    double ajuste = 0.0;
                    if (item.monto_consolidado.Equals(""))
                    {

                    }
                    else
                    {
                        int tiene_punto = item.monto_consolidado.IndexOf('.');
                        if (tiene_punto > 1)
                        {
                            monto = double.Parse(item.monto_consolidado, culture_punto);

                        }
                        else
                        {
                            monto = double.Parse(item.monto_consolidado, culture_coma);
                        }
                        //monto = double.Parse(item.monto_consolidado);
                    }
                    if (item.monto_ajuste.Equals(""))
                    {

                    }
                    else
                    {
                        int tiene_punto = item.monto_ajuste.IndexOf('.');
                        if (tiene_punto > 1)
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_punto);

                        }
                        else
                        {
                            ajuste = double.Parse(item.monto_ajuste, culture_coma);
                        }
                        //monto = double.Parse(item.monto_consolidado);
                    }

                    if (item.duplicado.Equals(""))
                    {
                        if (item.hay_ajuste)
                        {

                        }
                        else
                        {
                            contadorinio++;
                            if (monto == item.montoTotal)
                            {
                                activacion_revisados.Add(item);

                            }
                        }

                    }
                    else
                    {
                        activacion_duplex.Add(item);
                    }
                }

                List<int> donde_esta = new List<int>();
                //aqui hay que volver a cruzar
                var cargos_linq_validacion_2 = from asignados in activacion_revisados
                                               join validacion in auxiliar_consolidado_ajuste
                                                  on asignados.PCS equals validacion.pcs
                                               select new Activacion
                                               {
                                                   CuentaFinanciera = asignados.CuentaFinanciera,
                                                   recveivercustomer = asignados.recveivercustomer,
                                                   tipocargo = asignados.tipocargo,
                                                   codigodecargo = asignados.codigodecargo,
                                                   offwer = asignados.offwer,
                                                   nombreOFFER = asignados.nombreOFFER,
                                                   promo = asignados.promo,
                                                   description = asignados.description,
                                                   montoCargos = asignados.montoCargos,
                                                   montoDescuentos = asignados.montoDescuentos,
                                                   montoTotal = asignados.montoTotal,
                                                   tipodeCobro = asignados.tipodeCobro,
                                                   PCS = asignados.PCS,
                                                   secuenciadeCiclo = asignados.secuenciadeCiclo,
                                                   prorrateo = asignados.prorrateo,
                                                   TIPO = asignados.TIPO,
                                                   rut = asignados.rut,
                                                   verdaredo_falso = asignados.verdaredo_falso,
                                                   duplicado = asignados.duplicado,
                                                   monto_consolidado = validacion.monto_validador,
                                                   monto_ajuste = validacion.monto_ajuste

                                               };

                cuantos_me_dan = cargos_linq_validacion_2.Count();//con esto ya capturamos los 2 sobrantes, entonces ahora se debe eliminar estos 2 pcs de activacion revisados

                for (int i = 0; i < activacion_revisados.Count() - 1; i++)
                {

                    bool se_encontro = false;
                    foreach (var item in cargos_linq_validacion_2)
                    {
                        if (activacion_revisados[i].PCS.Equals(item.PCS))
                        {
                            se_encontro = true;
                            break;
                        }
                    }

                    if (se_encontro)
                    {
                        donde_esta.Add(i);
                    }

                }

                foreach (var item in donde_esta)
                {
                    activacion_no_revisados.Add(activacion_revisados[item]);
                    activacion_revisados.RemoveAt(item); //esto no se deberia hacer asi, esta malo
                }

                foreach (var item in activacion_revisados)
                {
                    activacions_revisados_asignados.Add(item);

                }

                foreach (var item in activacions_revisados_asignados)
                {

                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }

                    dt9.Rows.Add(new Object[] { "",
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
                                       item.TIPO,
                                       "SI"
                                       ,item.montoTotal
                                       ,"",
                                        "REVISADO"});
                }//aqui quedan todos los revisados
                oSLDocument.AddWorksheet("REVISADO"); //aqui vamos a crear el primer foreach para crear la hoja revisados
                oSLDocument.ImportDataTable(1, 1, dt9, true); //esto tiene que ir a la hoja revisado

                List<Activacion> auxiliar_asignados = activacions_asignados_portados;
                HashSet<string> pcs_activacion_revisados = new HashSet<string>(activacion_revisados.Select(x => x.PCS));

                auxiliar_asignados.RemoveAll(x => pcs_activacion_revisados.Contains(x.PCS));

                List<Activacion> auxiliar_verdadero = new List<Activacion>();
                List<Activacion> auxiliar_falso = new List<Activacion>();

                foreach (var item in activacions_asignados_portados)
                {
                    if (item.verdaredo_falso.Equals("VERDADERO"))
                    {
                        auxiliar_verdadero.Add(item);
                    }
                    else if (item.verdaredo_falso.Equals("FALSO"))
                    {
                        auxiliar_falso.Add(item);
                    }
                } // hacer la misma separacion que hize en el excel pero como.... separar los que solo son verdaderos de los que oslo son falsos

                List<int> donde_verdadero = new List<int>();

                for (int i = 0; i < auxiliar_verdadero.Count() - 1; i++)
                {
                    bool se_encontro = false;
                    foreach (var item in auxiliar_falso)
                    {
                        if (auxiliar_verdadero[i].PCS.Equals(item.PCS))
                        {
                            se_encontro = true;
                            break;
                        }
                    }

                    if (se_encontro)
                    {
                        donde_verdadero.Add(i);
                    }

                }//esta todo correcto hasta aqui marcare--------------------------------------------------------------------------

                foreach (var item in donde_verdadero)
                {
                    auxiliar_falso.Add(auxiliar_verdadero[item]);
                    //auxiliar_verdadero.RemoveAt(item); esto esta malo no descomentar
                }

                foreach (var item in auxiliar_falso)//
                {
                    auxiliar_verdadero.Remove(item);
                }

                //esta un poco molesto por no saber, 

                foreach (var item in auxiliar_verdadero) //esto se agregara algun dia 
                {
                    dt10.Rows.Add(new Object[] { "",
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
                                       item.verdaredo_falso,
                                       item.duplicado,
                                       item.tipodeCobro,
                                       long.Parse(item.PCS),
                                       long.Parse(item.secuenciadeCiclo),
                                       item.TIPO,
                                       ""
                                       ,item.montoTotal
                                       ,"",
                                        ""});
                }

                //
                //oSLDocument4.AddWorksheet("DUPLICADO"); //aqui vamos a crear el segundo foreach para crear la hoja duplicado
                //oSLDocument4.ImportDataTable(1, 1, dt10, true); //say something or i give in up on u sorry that i tarara rara ra ra

                var dt11 = new DataTable("CARGOS DE ACTIVACION");
                dt11.Columns.Add("ASIGNACION");
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
                dt11.Columns.Add("VERDADERO O FALSO", typeof(string));
                dt11.Columns.Add("DUPLICADO", typeof(string));
                dt11.Columns.Add("tipodeCobro", typeof(string));
                dt11.Columns.Add("PCS", typeof(long));
                dt11.Columns.Add("secuenciadeCiclo", typeof(long));
                dt11.Columns.Add("TIPO", typeof(string));
                dt11.Columns.Add("COBRAR");
                dt11.Columns.Add("MONTO", typeof(double));
                dt11.Columns.Add("CUOTA");
                dt11.Columns.Add("OBSERVACION");

                foreach (var item in activacions_asignados_portados)
                {
                    bool succes = long.TryParse(item.offwer, out long numero);
                    bool pcsNoNull = long.TryParse(item.PCS, out long pcsNoNulo);


                    if (succes)
                    {
                        item.offwer = numero.ToString();
                    }
                    else
                    {
                        item.offwer = offerPorDefecto.ToString();
                    }
                    if (pcsNoNull)
                    {
                        item.PCS = pcsNoNulo.ToString();
                    }
                    else
                    {
                        item.PCS = pcsPorDefecto.ToString();
                    }

                    dt11.Rows.Add(new Object[] { "",
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
                                       item.verdaredo_falso,
                                       item.duplicado,
                                       item.tipodeCobro,
                                       long.Parse(item.PCS),
                                       long.Parse(item.secuenciadeCiclo),
                                       item.TIPO,
                                       ""
                                       ,item.montoTotal
                                       ,"",
                                        ""});
                }


                oSLDocument7.AddWorksheet("NO REVISADOS"); //aqui vamos a crear el segundo foreach para crear la hoja duplicado
                oSLDocument7.ImportDataTable(1, 1, dt11, true);





                cuantos_me_dan = cargos_linq_validacion.Count();
                //PRIMERO TOMAMOS LOS PCS DE ASIGNACION Y LOS CRUZAMOS CON CONSOLIDADO TODOS LOS QUE CRUZEN SE DEBE COMPROBAR SI LOS MONTOS SON IGUALES Y NO SON DUPLICADOS. ESOS SON LOS PRIMEROS PASOS DE LOS PASOS.
                //PASO UNO CAPTURAR EL EXCEL DE CONSOLIDADO(SUS PCS Y MONTO)
                //PASO DOS COMPARAR PCS QUE SE ASIGNARAN CON CONSOLIDADO
                //PASO TRES TODO LO QUE CRUZA SE LE DEBERA COMPARAR EL MONTO
                //PASO CUATRO SI SON IGUALES LOS MONTOS SE DEBE VERIFICAR SI NO SON DUPLICADOS
                //Y AHI ESTAN LOS REVISADOS DE LA HOJA
                //CODIGO CCOCSUBEQPCOMM, ESTE CODIGO ES LO QUE HACE ALUSION AL CARGO DE ACTIVACION
                //
                //primero revisar ajuste y despues comparar montos.

                //------------------------------------------------------------------------------------------
                //DESPUES EMPEZAMOS A SACAR LOS DUPLICADOS
                //LOS DE E COMERCE NO SE DE DONDE SE SACAN Y AJUSTADOS TAMPOCO, PERO ES COSA DE IR DESGLOZANDO COMO LO HICE AHORA C:
                //agregar ajustes adicionales


                return true;

            }
            catch (Exception ex)
            {

                return false;

            }
        }
        public bool proceso_activacion_guardando_7()
        {
            try
            {
                oSLDocument.SaveAs(@aux_path + @"\Ciclo" + @"\Extraidas" + @"\EXTRAIDA DE CARGOS DE ACTIVACION CICLO XX.xlsx");
                oSLDocument2.SaveAs(@aux_path +@"\Ciclo" + @"\Ruta\" + @"\C_ACTIVACION.xlsx");
                oSLDocument3.SaveAs(@aux_path + @"\Ciclo" + @"\NO CRUZA CON CARGOS DE ACTIVACION CICLO FVM.xlsx");
                oSLDocument4.SaveAs(@aux_path + @"\Ciclo" + @"\Extraidas" + @"\CARGOS DE ACTIVACION_PARA_ASIGNAR.xlsx");
                oSLDocument7.SaveAs(@aux_path + @"\Ciclo" + @"\Extraidas" + @"\PORTADOS CICLO XX_MES_ANIO.xlsx");
                //oSLDocument5.SaveAs(@aux_path + @"\Ciclo" + @"\Extraidas" + @"\PORTADOS CICLO XX_MES_ANIO.xlsx");
                //oSLDocument6.SaveAs(@aux_path + @"\Ciclo" + @"\Extraidas" + @"\CARGOS DE ACTIVACION_PARA_ASIGNAR.xlsx"); //esto debe cambiar para eliminar el consolidado

                


                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }
    }
}
