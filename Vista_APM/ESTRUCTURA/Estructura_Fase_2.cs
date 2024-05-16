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
   public class Estructura_Fase_2
    {
        List<Cargos_Arrendamiento> cargos_arr_personas1;
        List<Ajustes> ajustes;
        List<Cargos_Arrendamiento> especiales_ajustes_arr = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> especiales_ajustes_arr_2 = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> especiales_ajustes_arr_dupli = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> especiales_ajustes_arr_NOdupli1 = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> especiales_ajustes_arr_NOdupli2 = new List<Cargos_Arrendamiento>();
        List<Cargos_Arrendamiento> especiales_ajustes_arr_sumado = new List<Cargos_Arrendamiento>();
        string path_ruta_intermedio = string.Empty;
        public void Fase_2(string aux)
        {
            path_ruta_intermedio = @aux;
            AlmacenarColecciones();

        }

        public void AlmacenarColecciones()
        {
            Coleccion_c_arr_personas1_ arr_Personas1_ = new Coleccion_c_arr_personas1_();
            Coleccion_ajustes coleccion_Ajustes = new Coleccion_ajustes();
            cargos_arr_personas1 = arr_Personas1_.GenerarListado(@path_ruta_intermedio);
            ajustes = coleccion_Ajustes.GenerarListado(@path_ruta_intermedio);
            ajustes.Count();
            cargos_arr_personas1.Count();
            CruceArrendamientoAjustesPcs();

        }

        public void CruceArrendamientoAjustesPcs()
        {
            var query = (from x1 in cargos_arr_personas1
                         join x2 in ajustes on x1.PCS equals x2.PCS
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
            especiales_ajustes_arr = query.ToList();
            especiales_ajustes_arr.Count();
            CrearC_ARR_PERSONA_AJESP_1();
        }
        public void PCSDuplicadosArrendamiento()
        {
            int contador = 0;
            foreach (var item1 in cargos_arr_personas1)
            {
                contador = 0;
                Cargos_Arrendamiento cargos_ = new Cargos_Arrendamiento();
                foreach (var item2 in especiales_ajustes_arr)
                {
                    if (item1.PCS.Equals(item2.PCS))
                    {
                        contador++;
                    }
                }
                if (contador == 0 && item1.PCS != "")
                {

                   cargos_.CuentaFinanciera   =  item1.CuentaFinanciera   ;
                   cargos_.recveivercustomer  =  item1.recveivercustomer  ;
                   cargos_.tipocargo          =  item1.tipocargo          ;
                   cargos_.codigodecargo      =  item1.codigodecargo      ;
                   cargos_.offwer             =  item1.offwer             ;
                   cargos_.nombreOFFER        =  item1.nombreOFFER        ;
                   cargos_.promo              =  item1.promo              ;
                   cargos_.description        =  item1.description        ;
                   cargos_.montoCargos        =  item1.montoCargos        ;
                   cargos_.montoDescuentos    =  item1.montoDescuentos    ;
                   cargos_.montoTotal         =  item1.montoTotal         ;
                   cargos_.tipodeCobro        =  item1.tipodeCobro        ;
                   cargos_.PCS                =  item1.PCS                ;
                   cargos_.secuenciadeCiclo   =  item1.secuenciadeCiclo   ;
                   cargos_.prorrateo          =  item1.prorrateo          ;



                    especiales_ajustes_arr_2.Add(cargos_);
                }
            }

            foreach (var item1 in especiales_ajustes_arr_2)
            {
                contador = 0;
                Cargos_Arrendamiento cargos_ = new Cargos_Arrendamiento();
                foreach (var item2 in especiales_ajustes_arr_2)
                {
                    if (item1.PCS.Equals(item2.PCS))
                    {
                        contador++;
                    }
                }
                if (contador >= 2)
                {
                    cargos_.CuentaFinanciera = item1.CuentaFinanciera;
                    cargos_.recveivercustomer = item1.recveivercustomer;
                    cargos_.tipocargo = item1.tipocargo;
                    cargos_.codigodecargo = item1.codigodecargo;
                    cargos_.offwer = item1.offwer;
                    cargos_.nombreOFFER = item1.nombreOFFER;
                    cargos_.promo = item1.promo;
                    cargos_.description = item1.description;
                    cargos_.montoCargos = item1.montoCargos;
                    cargos_.montoDescuentos = item1.montoDescuentos;
                    cargos_.montoTotal = item1.montoTotal;
                    cargos_.tipodeCobro = item1.tipodeCobro;
                    cargos_.PCS = item1.PCS;
                    cargos_.secuenciadeCiclo = item1.secuenciadeCiclo;
                    cargos_.prorrateo = item1.prorrateo;

                    especiales_ajustes_arr_dupli.Add(cargos_);
                }
                if (contador == 1 && item1.montoTotal >= 16000)
                {
                    cargos_.CuentaFinanciera = item1.CuentaFinanciera;
                    cargos_.recveivercustomer = item1.recveivercustomer;
                    cargos_.tipocargo = item1.tipocargo;
                    cargos_.codigodecargo = item1.codigodecargo;
                    cargos_.offwer = item1.offwer;
                    cargos_.nombreOFFER = item1.nombreOFFER;
                    cargos_.promo = item1.promo;
                    cargos_.description = item1.description;
                    cargos_.montoCargos = item1.montoCargos;
                    cargos_.montoDescuentos = item1.montoDescuentos;
                    cargos_.montoTotal = item1.montoTotal;
                    cargos_.tipodeCobro = item1.tipodeCobro;
                    cargos_.PCS = item1.PCS;
                    cargos_.secuenciadeCiclo = item1.secuenciadeCiclo;
                    cargos_.prorrateo = item1.prorrateo;
                    especiales_ajustes_arr_NOdupli1.Add(cargos_);
                }
                if (contador == 1 && item1.montoTotal < 16000)
                {
                    cargos_.CuentaFinanciera = item1.CuentaFinanciera;
                    cargos_.recveivercustomer = item1.recveivercustomer;
                    cargos_.tipocargo = item1.tipocargo;
                    cargos_.codigodecargo = item1.codigodecargo;
                    cargos_.offwer = item1.offwer;
                    cargos_.nombreOFFER = item1.nombreOFFER;
                    cargos_.promo = item1.promo;
                    cargos_.description = item1.description;
                    cargos_.montoCargos = item1.montoCargos;
                    cargos_.montoDescuentos = item1.montoDescuentos;
                    cargos_.montoTotal = item1.montoTotal;
                    cargos_.tipodeCobro = item1.tipodeCobro;
                    cargos_.PCS = item1.PCS;
                    cargos_.secuenciadeCiclo = item1.secuenciadeCiclo;
                    cargos_.prorrateo = item1.prorrateo;
                    especiales_ajustes_arr_NOdupli2.Add(cargos_);
                }

            }

            cargos_arr_personas1.Count();
            especiales_ajustes_arr.Count();
            especiales_ajustes_arr_2.Count();
            CrearC_ARR_PERSONA_AJESP_2();
        }

        public void CrearC_ARR_PERSONA_AJESP_1()
        {
            var miTable = new DataTable("ESPECIALES");
            miTable.Columns.Add("CuentaFinanciera");
            miTable.Columns.Add("recveivercustomer");
            miTable.Columns.Add("tipocargo");
            miTable.Columns.Add("codigodecargo");
            miTable.Columns.Add("offwer");
            miTable.Columns.Add("nombreOFFER");
            miTable.Columns.Add("promo");
            miTable.Columns.Add("description");
            miTable.Columns.Add("montoCargos");
            miTable.Columns.Add("montoDescuentos");
            miTable.Columns.Add("montoTotal");
            miTable.Columns.Add("tipodeCobro");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("secuenciadeCiclo");
            miTable.Columns.Add("prorrateo");


            foreach (var item in especiales_ajustes_arr)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ARCHIVOS INTERMEDIOS" + @"\C_ARR_PERSONA_AJESP_1.xlsx");
            PCSDuplicadosArrendamiento();
        }
        public void CrearC_ARR_PERSONA_AJESP_2()
        {
            var miTable = new DataTable("DUPLICADOS");
            miTable.Columns.Add("CuentaFinanciera");
            miTable.Columns.Add("recveivercustomer");
            miTable.Columns.Add("tipocargo");
            miTable.Columns.Add("codigodecargo");
            miTable.Columns.Add("offwer");
            miTable.Columns.Add("nombreOFFER");
            miTable.Columns.Add("promo");
            miTable.Columns.Add("description");
            miTable.Columns.Add("montoCargos");
            miTable.Columns.Add("montoDescuentos");
            miTable.Columns.Add("montoTotal");
            miTable.Columns.Add("tipodeCobro");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("secuenciadeCiclo");
            miTable.Columns.Add("prorrateo");


            foreach (var item in especiales_ajustes_arr_dupli)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ARCHIVOS INTERMEDIOS" + @"\C_ARR_PERSONA_AJESP_2.xlsx");
            CrearC_ARR_PERSONA_AJESP_2_ASIG();


        }
        public void CrearC_ARR_PERSONA_AJESP_2_ASIG()
        {

            var miTable = new DataTable("ASIGNADOS ARRENDAMIENTO");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("cobrar");
            miTable.Columns.Add("monto correcto");
            miTable.Columns.Add("cuota");
            miTable.Columns.Add("observacion");
            miTable.Columns.Add("procediendo a asignarlo");


            foreach (var item in especiales_ajustes_arr_dupli)
            {
                miTable.Rows.Add(new Object[] {item.PCS, "","","","",""});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ASIGNACION" + @"\C_ARR_PERSONA_AJESP_2_ASIG.xlsx");
            if (especiales_ajustes_arr_NOdupli1.Count() >= 1)
            {
                CrearC_ARR_PERSONA_AJESP_3();
            }
            

        }


        public void CrearC_ARR_PERSONA_AJESP_3()
        {
            var miTable = new DataTable("MONTO_MAYOR16");
            miTable.Columns.Add("CuentaFinanciera");
            miTable.Columns.Add("recveivercustomer");
            miTable.Columns.Add("tipocargo");
            miTable.Columns.Add("codigodecargo");
            miTable.Columns.Add("offwer");
            miTable.Columns.Add("nombreOFFER");
            miTable.Columns.Add("promo");
            miTable.Columns.Add("description");
            miTable.Columns.Add("montoCargos");
            miTable.Columns.Add("montoDescuentos");
            miTable.Columns.Add("montoTotal");
            miTable.Columns.Add("tipodeCobro");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("secuenciadeCiclo");
            miTable.Columns.Add("prorrateo");


            foreach (var item in especiales_ajustes_arr_NOdupli1)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ARCHIVOS INTERMEDIOS" + @"\C_ARR_PERSONA_AJESP_3.xlsx");
            if (especiales_ajustes_arr_NOdupli1.Count() >= 1)
            {
                CrearC_ARR_PERSONA_AJESPMENOR16000();
            }


        }

        public void CrearC_ARR_PERSONA_AJESPMENOR16000()
        {
            var miTable = new DataTable("MONTO_MMENOR16");
            miTable.Columns.Add("CuentaFinanciera");
            miTable.Columns.Add("recveivercustomer");
            miTable.Columns.Add("tipocargo");
            miTable.Columns.Add("codigodecargo");
            miTable.Columns.Add("offwer");
            miTable.Columns.Add("nombreOFFER");
            miTable.Columns.Add("promo");
            miTable.Columns.Add("description");
            miTable.Columns.Add("montoCargos");
            miTable.Columns.Add("montoDescuentos");
            miTable.Columns.Add("montoTotal");
            miTable.Columns.Add("tipodeCobro");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("secuenciadeCiclo");
            miTable.Columns.Add("prorrateo");


            foreach (var item in especiales_ajustes_arr_NOdupli2)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ARCHIVOS DE RESPALDO" + @"\C_ARR_PERSONA_AJESPMENOR16000.xlsx");
            CrearC_ARR_PERSONA_AJESP_3_ASIG();

        }

        public void CrearC_ARR_PERSONA_AJESP_3_ASIG()
        {

            var miTable = new DataTable("ASIGNADOS ARRENDAMIENTO");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("cobrar");
            miTable.Columns.Add("monto correcto");
            miTable.Columns.Add("cuota");
            miTable.Columns.Add("observacion");
            miTable.Columns.Add("procediendo a asignarlo");


            foreach (var item in especiales_ajustes_arr_NOdupli1)
            {
                miTable.Rows.Add(new Object[] { item.PCS, "", "", "", "", "" });
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ASIGNACION" + @"\C_ARR_PERSONA_AJESP_3_ASIG.xlsx");

            CrearC_ARR_PERSONA_AJESP_FINAL();



        }

        public void CrearC_ARR_PERSONA_AJESP_FINAL()
        {
            var miTable = new DataTable("ESPECIALES");
            miTable.Columns.Add("CuentaFinanciera");
            miTable.Columns.Add("recveivercustomer");
            miTable.Columns.Add("tipocargo");
            miTable.Columns.Add("codigodecargo");
            miTable.Columns.Add("offwer");
            miTable.Columns.Add("nombreOFFER");
            miTable.Columns.Add("promo");
            miTable.Columns.Add("description");
            miTable.Columns.Add("montoCargos");
            miTable.Columns.Add("montoDescuentos");
            miTable.Columns.Add("montoTotal");
            miTable.Columns.Add("tipodeCobro");
            miTable.Columns.Add("PCS");
            miTable.Columns.Add("secuenciadeCiclo");
            miTable.Columns.Add("prorrateo");

           

            especiales_ajustes_arr_sumado = especiales_ajustes_arr
                           .GroupBy(x => x.PCS)
                           .Select(x => new Cargos_Arrendamiento { PCS = x.Key, CuentaFinanciera = x.First().CuentaFinanciera, recveivercustomer = x.First().recveivercustomer, tipocargo = x.First().tipocargo, codigodecargo = x.First().codigodecargo,
                               offwer = x.First().offwer, nombreOFFER = x.First().nombreOFFER, promo = x.First().promo, description = x.First().description,montoCargos = x.First().montoCargos, montoDescuentos = x.First().montoDescuentos,
                               tipodeCobro = x.First().tipodeCobro, secuenciadeCiclo = x.First().secuenciadeCiclo,prorrateo = x.First().prorrateo, montoTotal = x.Sum(y => y.montoTotal) })
                           .ToList();

            


            foreach (var item in especiales_ajustes_arr_sumado)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\ARCHIVOS SALIDA" + @"\C_ARR_PERSONA_AJESP_FINAL.xlsx");
        }

        













    }
}
