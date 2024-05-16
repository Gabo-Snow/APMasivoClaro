using APM;
using Proyecto_Automatizacion.MODELO.CLASES;
using Proyecto_Automatizacion.MODELO.COLECCIONES;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Proyecto_Automatizacion.ESTRUCTURA
{
    public class Estructura_Inicio
    {
        List<Cargos_Arrendamiento> cargos;
        List<Cartera>              carteras;
        List<Cargos_Arrendamiento> cargos_empresas;
        List<Cargos_Arrendamiento> cargos_0;
        List<Cargos_Arrendamiento> cargos_distinto_0;
        List<Cargos_Arrendamiento> cargos___ = new List<Cargos_Arrendamiento>();


        string path_ruta_inicial = string.Empty;
        string path_ruta_entrada = string.Empty;
        string path_ruta_respaldo = string.Empty;
        string path_ruta_intermedio = string.Empty;
        string path_ruta_salida = string.Empty;
        string path_ruta_asignacionl = string.Empty;
        public void InicioCiclo()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            path_ruta_inicial = fbd.SelectedPath.ToString();
            CreacionCarpetas();
            AlmacenarColecciones();
        }

        public void AlmacenarColecciones()
        {
            Coleccion_c_ar c_Ar     = new Coleccion_c_ar();
            Coleccion_cartera cart  = new Coleccion_cartera();
            cargos   = c_Ar.GenerarListado(path_ruta_inicial);
            carteras = cart.GenerarListado(path_ruta_inicial);
            var query = cargos.GroupJoin(carteras, k1 => k1.CuentaFinanciera, k2 => k2.BA_NO, (k1, k2) => 
            new Cargos_Arrendamiento
             {
                 segmento = k2.SingleOrDefault() == null ? k1.segmento : k2.SingleOrDefault().SEGMENTO,
                 CuentaFinanciera = k1.CuentaFinanciera,
                 recveivercustomer = k1.recveivercustomer,
                 tipocargo = k1.tipocargo,
                 codigodecargo = k1.codigodecargo,
                 offwer = k1.offwer,
                 nombreOFFER = k1.nombreOFFER,
                 promo = k1.promo,
                 description = k1.description,
                 montoCargos = k1.montoCargos,
                 montoDescuentos = k1.montoDescuentos,
                 montoTotal = k1.montoTotal,
                 tipodeCobro = k1.tipodeCobro,
                 PCS = k1.PCS,
                 secuenciadeCiclo = k1.secuenciadeCiclo,
                 prorrateo = k1.prorrateo,
             });
            foreach (var item in query)
            {
                Cargos_Arrendamiento cargos_ = new Cargos_Arrendamiento();
                if (item.segmento.Equals(""))
                {
                    cargos_.CuentaFinanciera = item.CuentaFinanciera;
                    cargos_.recveivercustomer = item.recveivercustomer;
                    cargos_.tipocargo = item.tipocargo;
                    cargos_.codigodecargo = item.codigodecargo;
                    cargos_.offwer = item.offwer;
                    cargos_.nombreOFFER = item.nombreOFFER;
                    cargos_.promo = item.promo;
                    cargos_.description = item.description;
                    cargos_.montoCargos = item.montoCargos;
                    cargos_.montoDescuentos = item.montoDescuentos;
                    cargos_.montoTotal = item.montoTotal;
                    cargos_.tipodeCobro = item.tipodeCobro;
                    cargos_.PCS = item.PCS;
                    cargos_.secuenciadeCiclo = item.secuenciadeCiclo;
                    cargos_.prorrateo = item.prorrateo;
                    cargos_.segmento = "N/A";
                    cargos___.Add(cargos_);
                }
                else
                {
                    cargos_.CuentaFinanciera = item.CuentaFinanciera;
                    cargos_.recveivercustomer = item.recveivercustomer;
                    cargos_.tipocargo = item.tipocargo;
                    cargos_.codigodecargo = item.codigodecargo;
                    cargos_.offwer = item.offwer;           
                    cargos_.nombreOFFER = item.nombreOFFER;
                    cargos_.promo = item.promo;
                    cargos_.description = item.description;
                    cargos_.montoCargos = item.montoCargos;
                    cargos_.montoDescuentos = item.montoDescuentos;
                    cargos_.montoTotal = item.montoTotal;
                    cargos_.tipodeCobro = item.tipodeCobro;
                    cargos_.PCS = item.PCS;
                    cargos_.secuenciadeCiclo = item.secuenciadeCiclo;
                    cargos_.prorrateo = item.prorrateo;
                    cargos_.segmento = item.segmento;
                    cargos___.Add(cargos_);
                }


            }

            CruceCarteraCargosEmpresa();
        }

        public void CruceCarteraCargosEmpresa()
        {
            var query = (from x1 in cargos
                         join x2 in carteras on x1.CuentaFinanciera equals x2.BA_NO
                         where x2.SEGMENTO.Equals("Empresas") || x2.SEGMENTO.Equals("Corporaciones") || x2.SEGMENTO.Equals("Gobierno") || x2.SEGMENTO.Equals("Pruebas")
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
                             prorrateo = x1.prorrateo,
                             segmento = x2.SEGMENTO,
                             rut = x1.rut
                         }) ;

            ///Console.WriteLine(cargos_empresas.Count());
            cargos_empresas = query.ToList();
            CrearC_ARR_CARTERA_EMPRESAS();
        }

        public void CruceCarteraCargosMonto0()
        {
            var query = (from x1 in cargos___
                                 where x1.montoTotal == 0 && x1.segmento != "Empresas" && x1.segmento != ("Corporaciones") && x1.segmento != ("Gobierno") && x1.segmento != ("Pruebas")
                                 select x1);
            cargos_0 = query.ToList();
            Console.WriteLine("los con 0 : "+ cargos_0.Count());
            CrearC_ARR_PERSONA_0();
        }
        public void CruceCarteraCargosMontosDistintos0()
        {
            var query = (from x1 in cargos___
                         where x1.montoTotal != 0 && x1.segmento != "Empresas" && x1.segmento != ("Corporaciones") && x1.segmento != ("Gobierno") && x1.segmento != ("Pruebas")
                         select x1);
            cargos_distinto_0 = query.ToList();
            Console.WriteLine("los con distinto a 0 : " + cargos_distinto_0.Count());
            CrearC_ARR_PERSONA_1();
        }
        public void CreacionCarpetas()
        {
            Console.WriteLine("Creando el directorio: {0}", @path_ruta_inicial);
            path_ruta_respaldo      = @path_ruta_inicial + @"\ARCHIVOS DE RESPALDO";
            DirectoryInfo di3       = Directory.CreateDirectory(@path_ruta_inicial +@"\ARCHIVOS DE RESPALDO");
            path_ruta_intermedio    = @path_ruta_inicial + @"\ARCHIVOS INTERMEDIOS";
            DirectoryInfo di4       = Directory.CreateDirectory(@path_ruta_inicial + @"\ARCHIVOS INTERMEDIOS");
            path_ruta_intermedio    = @path_ruta_inicial + @"\ARCHIVOS INTERMEDIOS";
            DirectoryInfo di5       = Directory.CreateDirectory(@path_ruta_inicial + @"\ASIGNACION");
            path_ruta_asignacionl   = @path_ruta_inicial + @"\ARCHIVOS SALIDA";
            DirectoryInfo di6       = Directory.CreateDirectory(@path_ruta_inicial + @"\ARCHIVOS SALIDA");
         
        }
        public void CrearC_ARR_CARTERA_EMPRESAS() //paso 1
        {
            var miTable = new DataTable("EMPRESAS");
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
            miTable.Columns.Add("SEGMENTO");
            miTable.Columns.Add("prorrateo");


            foreach (var item in cargos_empresas)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_respaldo + @"\C_ARR_CARTERA EMPRESAS.xlsx");
            CruceCarteraCargosMonto0();
        }

        public void CrearC_ARR_PERSONA_0()
        {
            var miTable = new DataTable("ARR_PERSONA0");
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
            miTable.Columns.Add("SEGMENTO");
            miTable.Columns.Add("prorrateo");


            foreach (var item in cargos_0)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_respaldo + @"\C_ARR_PERSONA_0.xlsx");
            CruceCarteraCargosMontosDistintos0();
        }
        public void CrearC_ARR_PERSONA_1()
        {
            var miTable = new DataTable("ARR_PERSONA1");
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
            miTable.Columns.Add("SEGMENTO");
            miTable.Columns.Add("prorrateo");


            foreach (var item in cargos_distinto_0)
            {
                miTable.Rows.Add(new Object[] { item.CuentaFinanciera, item.recveivercustomer, item.tipocargo,item.codigodecargo, item.offwer,
                    item.nombreOFFER, item.promo, item.description, item.montoCargos, item.montoDescuentos, item.montoTotal , item.tipodeCobro
                    , item.PCS, item.secuenciadeCiclo,item.segmento, item.prorrateo});
            }

            SaveExcel.BuildExcel(miTable, @path_ruta_intermedio + @"\C_ARR_PERSONA_1.xlsx");
        }
        public string Enviar_Path()
        {
            cargos = null;
            carteras = null;
            cargos_empresas = null;
            cargos_0 = null;
            cargos_distinto_0 = null;
            return @path_ruta_inicial;
        }


    }

        
        
        
        
        
        
        
        
        
        
        
        
        
}
