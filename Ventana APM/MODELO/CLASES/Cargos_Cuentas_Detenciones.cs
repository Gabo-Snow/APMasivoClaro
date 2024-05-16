using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pruebas.MODELO.CLASES
{
   public class Cargos_Cuentas_Detenciones
    {
        public string CuentaFinanciera          { get; set; }
        public string RECEIVER_CUSTOMER         { get; set; }
        public string TIPO_CARGO                { get; set; }
        public string CodigodeCargo             { get; set; }
        public string DESCRIPTION               { get; set; }
        public double MONTO_CARGOS              { get; set; }
        public double MONTO_DESCUENTOS          { get; set; }
        public double MONTO_TOTAL               { get; set; }
        public string MOTIVO                    { get; set; }
        public string OBSERVACION               { get; set; }

        public Cargos_Cuentas_Detenciones()
        {
            this.Init();
        }



        private void Init()
        {
            CuentaFinanciera    = string.Empty;
            RECEIVER_CUSTOMER   = string.Empty;
            TIPO_CARGO = string.Empty;
            CodigodeCargo = string.Empty;
            DESCRIPTION = string.Empty;
            MONTO_CARGOS = 0.0;
            MONTO_DESCUENTOS = 0.0;
            MONTO_TOTAL = 0.0;               
        }
    }
}

