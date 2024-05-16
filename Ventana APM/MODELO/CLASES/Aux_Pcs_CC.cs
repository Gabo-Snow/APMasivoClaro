using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Aux_Pcs_CC
    {
        public string CuentaFinanciera { get; set; }
        public string PCS { get; set; }

        public Aux_Pcs_CC()
        {
            this.Init();
        }



        private void Init()
        {
            CuentaFinanciera = string.Empty;
            PCS = string.Empty;

        }
    }
}
