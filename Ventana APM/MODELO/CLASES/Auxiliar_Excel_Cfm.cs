using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
   public class Auxiliar_Excel_Cfm
    {
        public string PCS { get; set; }
        public string CuentaFinanciera { get; set; }
        public string amount { get; set; }
        public string cf_mas_iva { get; set; }
        public string cargos { get; set; }
        public string ajustes { get; set; }
        public string campana { get; set; }


        public Auxiliar_Excel_Cfm()
        {
            this.Init();
        }



        private void Init()
        {
            CuentaFinanciera = string.Empty;
            PCS = string.Empty;
            amount = string.Empty;
            cf_mas_iva = string.Empty;
            cargos = string.Empty;
            ajustes = string.Empty;
            campana = string.Empty;
        }
    }
}
