using Proyecto_Automatizacion.MODELO.CLASES;
using Proyecto_Automatizacion.MODELO.COLECCIONES;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Proyecto_Automatizacion.ESTRUCTURA
{
   public class Estructura_3
    {
        //anticuota cruzar por portados
        string path_ruta = string.Empty;
        List<Anticuota> Anticuotas = new List<Anticuota>();
        List<Portabilidad> Portabilidads = new List<Portabilidad>();
        List<Anticuota> Anticuotas2 = new List<Anticuota>();
        List<PlanSolidario> planSolidarios = new List<PlanSolidario>();
        List<FVM> fVMs = new List<FVM>();
        List<HABILITACIONES> hABILITACIONEs = new List<HABILITACIONES>();
        public void Fase_3(string aux)
        {
            path_ruta = @aux;
            AlmacenarColecciones();
            CrucePortabilidadxAnticuota();
            CruceAnticuotaxCartera();

        }

        private void AlmacenarColecciones()
        {
            Coleccion_Portabilidad coleccion_Portabilidad = new Coleccion_Portabilidad();
            Portabilidads = coleccion_Portabilidad.GenerarListado(@path_ruta);

            Coleccion_Anticuota coleccion_Anticuota = new Coleccion_Anticuota();
            Anticuotas = coleccion_Anticuota.GenerarListado(path_ruta);

            Coleccion_PlanSolidario coleccion_PlanSolidario = new Coleccion_PlanSolidario();
            planSolidarios = coleccion_PlanSolidario.GenerarListado(@path_ruta);

            Coleccion_FVM coleccion_FVM = new Coleccion_FVM();
            fVMs = coleccion_FVM.GenerarListado(@path_ruta);

            Coleccion_HABILITACIONES coleccion_HABILITACIONES = new Coleccion_HABILITACIONES();
            hABILITACIONEs = coleccion_HABILITACIONES.GenerarListado(@path_ruta);

        }

        public void CrucePortabilidadxAnticuota()
        {
            var query = from x1 in Portabilidads
                        join x2 in Anticuotas on x1.NRO_DE_PCS equals x2.PCS
                        select new Anticuota // tiene que ser anticuota y extraerle modalidad y sumar los montos de los duplicados
                        {
                            CuentaFinanciera   = x2.CuentaFinanciera,
                            recveivercustomer  = x2.recveivercustomer,
                            tipocargo          = x2.tipocargo,        
                            codigodecargo      = x2.codigodecargo,    
                            offwer             = x2.offwer,           
                            nombreOFFER        = x2.nombreOFFER,      
                            promo              = x2.promo,            
                            description        = x2.description,      
                            montoCargos        = x2.montoCargos,      
                            montoDescuentos    = x2.montoDescuentos,  
                            montoTotal         = x2.montoTotal,       
                            tipodeCobro        = x2.tipodeCobro,      
                            PCS                = x2.PCS,              
                            secuenciadeCiclo   = x2.secuenciadeCiclo, 
                            prorrateo          = x2.prorrateo,        
                            segmento           = x2.segmento,         
                            rut                = "",              
                            TIPO               = x1.MODALIDAD             
                        };
            List<Anticuota> anticuotas_sumando = query.ToList();
            List<Anticuota> anticuotas_distinto = anticuotas_sumando.GroupBy(p => p.PCS)
                   .Select(grp => grp.First())
                   .ToList();

            //Console.WriteLine("acticuota distintos : {0}", anticuotas_distinto.Count());

            for (int i = 0; i < anticuotas_distinto.Count; i++)
            {
                Anticuota anticuota = new Anticuota();
                for (int j = 0; j < anticuotas_sumando.Count; j++)
                {
                    if (anticuotas_distinto[i].PCS.Equals(anticuotas_sumando[j].PCS))
                    {

                        anticuota.CuentaFinanciera = anticuotas_sumando[j].CuentaFinanciera;
                        anticuota.recveivercustomer = anticuotas_sumando[j].recveivercustomer;
                        anticuota.tipocargo = anticuotas_sumando[j].tipocargo;
                        anticuota.codigodecargo = anticuotas_sumando[j].codigodecargo;
                        anticuota.offwer = anticuotas_sumando[j].offwer;
                        anticuota.nombreOFFER = anticuotas_sumando[j].nombreOFFER;
                        anticuota.promo = anticuotas_sumando[j].promo;
                        anticuota.description = anticuotas_sumando[j].description;
                        anticuota.montoCargos = anticuotas_sumando[j].montoCargos;
                        anticuota.montoDescuentos = anticuotas_sumando[j].montoDescuentos;
                        anticuota.montoTotal = anticuotas_sumando[j].montoTotal+ anticuota.montoTotal;
                        anticuota.tipodeCobro = anticuotas_sumando[j].tipodeCobro;
                        anticuota.PCS = anticuotas_sumando[j].PCS;
                        anticuota.secuenciadeCiclo = anticuotas_sumando[j].secuenciadeCiclo;
                        anticuota.prorrateo = anticuotas_sumando[j].prorrateo;
                        anticuota.segmento = anticuotas_sumando[j].segmento;
                        anticuota.rut = "";
                        anticuota.TIPO = anticuotas_sumando[j].TIPO;
                    }
                }
                if (anticuota.montoTotal == 0)
                {
                    anticuota.DATO_MONTO = "NETEADO";
                    Anticuotas2.Add(anticuota);
                }
                else
                {
                    anticuota.DATO_MONTO = "NO NETEADO";
                    Anticuotas2.Add(anticuota);
                }
                


            }




        }
        public void CruceAnticuotaxCartera()
        {
            Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
            List<Cartera> carteras = coleccion_Cartera.GenerarListado(path_ruta);
            Console.WriteLine("---------------------------------------------------------------");
            var miTable = new DataTable("Portabilidad Personas");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("MONTO TOTAL");
            miTable.Columns.Add("DATO MONTO");

            var query_1 = (from x1 in Anticuotas2
                         join x2 in carteras on x1.CuentaFinanciera equals x2.BA_NO
                         where x2.SEGMENTO.Equals("Empresas") || x2.SEGMENTO.Equals("Corporaciones") || x2.SEGMENTO.Equals("Gobierno") || x2.SEGMENTO.Equals("Pruebas")
                         select new Anticuota
                         {
                             CuentaFinanciera = x1.CuentaFinanciera,
                             recveivercustomer = x1.recveivercustomer,
                             tipocargo = x1.tipocargo,
                             codigodecargo = x1.codigodecargo,
                             offwer = x1.offwer,
                             nombreOFFER = x1.nombreOFFER,
                             promo = x1.promo,
                             description = x1.description,
                             montoCargos = x1.montoCargos,
                             montoDescuentos = x1.montoDescuentos,
                             montoTotal = x1.montoTotal,
                             tipodeCobro = x1.tipodeCobro,
                             PCS = x1.PCS,
                             secuenciadeCiclo = x1.secuenciadeCiclo,
                             prorrateo = x1.prorrateo,
                             segmento = x2.SEGMENTO,
                             rut = x1.rut,
                             DATO_MONTO = x1.DATO_MONTO
                         });

            var query_2 = (from x1 in Anticuotas2
                         join x2 in carteras on x1.CuentaFinanciera equals x2.BA_NO
                         where x2.SEGMENTO != "Empresas" || x2.SEGMENTO != "Corporaciones" || x2.SEGMENTO != "Gobierno" || x2.SEGMENTO != "Pruebas"
                         select new Anticuota
                         {
                             CuentaFinanciera = x1.CuentaFinanciera,
                             recveivercustomer = x1.recveivercustomer,
                             tipocargo = x1.tipocargo,
                             codigodecargo = x1.codigodecargo,
                             offwer = x1.offwer,
                             nombreOFFER = x1.nombreOFFER,
                             promo = x1.promo,
                             description = x1.description,
                             montoCargos = x1.montoCargos,
                             montoDescuentos = x1.montoDescuentos,
                             montoTotal = x1.montoTotal,
                             tipodeCobro = x1.tipodeCobro,
                             PCS = x1.PCS,
                             secuenciadeCiclo = x1.secuenciadeCiclo,
                             prorrateo = x1.prorrateo,
                             segmento = x2.SEGMENTO,
                             rut = x1.rut
                         });

            Console.WriteLine("Empresas {0}: ---------------------------------------------------", query_1.Count());


            foreach (var item in query_1)
            {
                miTable.Rows.Add(new Object[] { item.PCS, item.montoTotal, item.DATO_MONTO });
                //Console.WriteLine("PCS {0}| TIPO {1}",item.NRO_DE_PCS,item.MODALIDAD);
            }

            SaveExcel.BuildExcel(miTable, @path_ruta + @"\ARCHIVOS DE RESPALDO" + @"\CA_CORPYEMP.xlsx");

            Console.WriteLine("Pymes o varios {0}: ---------------------------------------------",query_2.Count());
            foreach (var item in query_2)
            {
                miTable.Rows.Add(new Object[] { item.PCS, item.montoTotal, item.DATO_MONTO });
                //Console.WriteLine("PCS {0}| TIPO {1}",item.NRO_DE_PCS,item.MODALIDAD);
            }

            SaveExcel.BuildExcel(miTable, @path_ruta + @"\ARCHIVOS INTERMEDIOS" + @"\CA_PERSONAS.xlsx");

            CrucePlanSolidarioxCa_Personas(query_2.ToList());


        }

        public void CrucePlanSolidarioxCa_Personas(List<Anticuota> anticuotas)
        {
            var query = (from x1 in anticuotas
                         join x2 in fVMs on x1.PCS equals x2.PCS
                         select new Cargos_Arrendamiento
                         {
                             CuentaFinanciera = x1.CuentaFinanciera,
                             recveivercustomer = x1.recveivercustomer,
                             tipocargo = x1.tipocargo,
                             codigodecargo = x1.codigodecargo,
                             offwer = x1.offwer,
                             nombreOFFER = x1.nombreOFFER,
                             promo = x1.promo,
                             description = x1.description,
                             montoCargos = x1.montoCargos,
                             montoDescuentos = x1.montoDescuentos,
                             montoTotal = x1.montoTotal,
                             tipodeCobro = x1.tipodeCobro,
                             PCS = x1.PCS,
                             secuenciadeCiclo = x1.secuenciadeCiclo,
                             prorrateo = x1.prorrateo
                         });


            var miTable = new DataTable("Portabilidad Personas");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("MONTO TOTAL");
            miTable.Columns.Add("DATO MONTO");

            Console.WriteLine("Plan Solidario cruce {0}: ---------------------------------------------", planSolidarios.Count());
            foreach (var item in query)
            {
                miTable.Rows.Add(new Object[] { item.PCS, item.montoTotal,item.CuentaFinanciera});
                //Console.WriteLine("PCS {0}| TIPO {1}",item.NRO_DE_PCS,item.MODALIDAD);
            }

            SaveExcel.BuildExcel(miTable, @path_ruta + @"\ARCHIVOS INTERMEDIOS" + @"\CA_PERSONAS_PLAN SOLIDARIO.xlsx");

            var miTable2 = new DataTable("Solidario no");
            miTable2.Columns.Add("PCS");
            miTable2.Columns.Add("MONTO TOTAL");
            miTable2.Columns.Add("DATO MONTO");
            List<Anticuota> noes_plan_solidario = new List<Anticuota>(); 


            foreach (var item in anticuotas)
            {
                Anticuota anticuota = new Anticuota();
                foreach (var item2 in planSolidarios)
                {
                    if (item.PCS != item2.PCS)
                    {
                        miTable2.Rows.Add(new Object[] { item.PCS, item.montoTotal, item.CuentaFinanciera });
                        anticuota = item;
                    }
                }

                noes_plan_solidario.Add(anticuota);
            }

            SaveExcel.BuildExcel(miTable2, @path_ruta + @"\ARCHIVOS INTERMEDIOS" + @"\CA_PERSONAS_NO_PLAN SOLIDARIO.xlsx");
            CruceFvmxCa_Personas(noes_plan_solidario);
        }

        public void CruceFvmxCa_Personas(List<Anticuota> anticuotas)
        {
            var query = (from x1 in anticuotas
                         join x2 in fVMs on x1.PCS equals x2.PCS
                         select new Cargos_Arrendamiento
                         {
                             CuentaFinanciera = x1.CuentaFinanciera,
                             recveivercustomer = x1.recveivercustomer,
                             tipocargo = x1.tipocargo,
                             codigodecargo = x1.codigodecargo,
                             offwer = x1.offwer,
                             nombreOFFER = x1.nombreOFFER,
                             promo = x1.promo,
                             description = x1.description,
                             montoCargos = x1.montoCargos,
                             montoDescuentos = x1.montoDescuentos,
                             montoTotal = x1.montoTotal,
                             tipodeCobro = x1.tipodeCobro,
                             PCS = x1.PCS,
                             secuenciadeCiclo = x1.secuenciadeCiclo,
                             prorrateo = x1.prorrateo
                         });

            var miTable = new DataTable("Portabilidad Personas");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("MONTO TOTAL");
            miTable.Columns.Add("DATO MONTO");

            Console.WriteLine("FVM cruce {0}: ---------------------------------------------", planSolidarios.Count());
            foreach (var item in query)
            {
                miTable.Rows.Add(new Object[] { item.PCS, item.montoTotal, item.CuentaFinanciera });
                //Console.WriteLine("PCS {0}| TIPO {1}",item.NRO_DE_PCS,item.MODALIDAD);
            }

            SaveExcel.BuildExcel(miTable, @path_ruta + @"\ARCHIVOS INTERMEDIOS" + @"\CA_PERSONA_NO PLAN SOLIDARIO_FVM.xlsx");

            var miTable2 = new DataTable("Solidario no");
            miTable2.Columns.Add("PCS");
            miTable2.Columns.Add("MONTO TOTAL");
            miTable2.Columns.Add("DATO MONTO");
            List<Anticuota> noes_plan_solidario = new List<Anticuota>();


            foreach (var item in anticuotas)
            {
                Anticuota anticuota = new Anticuota();
                foreach (var item2 in fVMs)
                {
                    if (item.PCS != item2.PCS)
                    {
                        miTable2.Rows.Add(new Object[] { item.PCS, item.montoTotal, item.CuentaFinanciera });
                        anticuota = item;
                    }
                }

                noes_plan_solidario.Add(anticuota);
            }

            SaveExcel.BuildExcel(miTable2, @path_ruta + @"\ARCHIVOS INTERMEDIOS" + @"\CA_PERSONA_NO PLAN SOLIDARIO_NO FVM.xlsx");

            var query_2 = (from x1 in anticuotas
                           join x2 in hABILITACIONEs on x1.PCS equals x2.PCS
                            select new Cargos_Arrendamiento
                            {
                             CuentaFinanciera = x1.CuentaFinanciera,
                             recveivercustomer = x1.recveivercustomer,
                             tipocargo = x1.tipocargo,
                             codigodecargo = x1.codigodecargo,
                             offwer = x1.offwer,
                             nombreOFFER = x1.nombreOFFER,
                             promo = x1.promo,
                             description = x1.description,
                             montoCargos = x1.montoCargos,
                             montoDescuentos = x1.montoDescuentos,
                             montoTotal = x1.montoTotal,
                             tipodeCobro = x1.tipodeCobro,
                             PCS = x1.PCS,
                             secuenciadeCiclo = x1.secuenciadeCiclo,
                             prorrateo = x1.prorrateo
                            });

            
        }

    }
}
    
