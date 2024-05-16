using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Ajustes
    {
        public string PCS { get; set; }
        public string CuentaFinanciera { get; set; }
        public string codigodecargo { get; set; }
        public double amount { get; set; }
        public string ciclo { get; set; }

        public Ajustes()
        {
            this.Init();
        }



        private void Init()
        {
            CuentaFinanciera = string.Empty;
            codigodecargo = string.Empty;
            ciclo = string.Empty;
            amount = 0.0;
            PCS = string.Empty;
        }
    }
}
