using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Cargos_Cuenta_Corporaciones
    {
                         
        //2021	   17	     14	          5301900	Comp SOC DE TRANSPORTES STEUER HOLMGREN Y COMPANIA LTDA	77785440-2	CALL_DETAILS Impresión Detalle de Llamadas CLP	0	0	0
        public string CYCLE_YEAR { get; set; }
        public string PERIOD_KEY { get; set; }
        public string CuentaFinanciera { get; set; }
        public string CYCLE_CODE { get; set; }
        public string BA_NO { get; set; }
        public string NOMBRE_CLIENTE { get; set; }
        public string RUT { get; set; }
        public string CHARGE_CODE { get; set; }
        public string DESCRIPTION { get; set; }
        public string AMOUNT_CURRENCY { get; set; }
        public string MONTO_CARGOS { get; set; }
        public string MONTO_DESCUENTOS { get; set; }
        public string MONTO_TOTAL { get; set; }

        public Cargos_Cuenta_Corporaciones()
        {
            this.Init();
        }



        private void Init()
        {
            CYCLE_YEAR = string.Empty;
            PERIOD_KEY = string.Empty;
            CuentaFinanciera = string.Empty;
            CYCLE_CODE = string.Empty;
            BA_NO = string.Empty;
            NOMBRE_CLIENTE = string.Empty;
            RUT = string.Empty;
            CHARGE_CODE = string.Empty;
            DESCRIPTION = string.Empty;
            AMOUNT_CURRENCY = string.Empty;
            MONTO_CARGOS = string.Empty;
            MONTO_DESCUENTOS = string.Empty;
            MONTO_TOTAL = string.Empty;
        }
    }
}
