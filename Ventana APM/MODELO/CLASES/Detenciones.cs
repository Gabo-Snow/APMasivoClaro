using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
   public class Detenciones
    {
                                        
        public string RUT { get; set; }
        public string RAZON_SOCIAL { get; set; }
        public string EDAC { get; set; }
        public string SUPERVISOR { get; set; }
        public string CUENTA_FINANCIERA { get; set; }
        public string NOMBRE_CLIENTE { get; set; }
        public string CICLO { get; set; }
        public string ACCION { get; set; }
        public string JustificaciOn_no_emisiones { get; set; }
        public string CANTIDAD_DE_LINEAS { get; set; }
        public string USO_BACK { get; set; }
        public string Plan_Q_de_líneas { get; set; }
        public string CF_Neto { get; set; }
        public string Cargo_de_activacion_NETO { get; set; }
        public string Observacion_CUOTA { get; set; }
        public string Remanentes_u_opcion_de_compra_NETO { get; set; }
        
        public Detenciones()
        {
            this.Init();
        }



        private void Init()
        {
            RAZON_SOCIAL = string.Empty;
            EDAC = string.Empty;
            SUPERVISOR = string.Empty;
            CUENTA_FINANCIERA = string.Empty;
            NOMBRE_CLIENTE = string.Empty;
            CICLO = string.Empty;
            RUT = string.Empty;
            ACCION = string.Empty;
            JustificaciOn_no_emisiones = string.Empty;
            CANTIDAD_DE_LINEAS = string.Empty;
            USO_BACK = string.Empty;
            Plan_Q_de_líneas = string.Empty;
            CF_Neto = string.Empty;
            Cargo_de_activacion_NETO = string.Empty;
            Observacion_CUOTA = string.Empty;
            Remanentes_u_opcion_de_compra_NETO = string.Empty;
        }
    }
}
