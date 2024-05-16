using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class PlanSolidario
    {
        public string PCS { get; set; }
        public string CUENTA { get; set; }
        public string CICLO { get; set; }
        public double MONTO { get; set; }
        public string APLICA_CDP { get; set; }
        public string CAMPANA { get; set; }


        public PlanSolidario()
        {
            this.Init();
        }



        private void Init()
        {
            PCS = string.Empty;
            CUENTA = string.Empty;
            CICLO = string.Empty;
            MONTO = 0.0;
            APLICA_CDP = string.Empty;
            CAMPANA = string.Empty;
        }
    }
}
