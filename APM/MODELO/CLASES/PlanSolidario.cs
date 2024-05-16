using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APM.MODELO.CLASES
{
    public class PlanSolidario
    {
        public string PCS { get; set; }
        public string CUENTA { get; set; }
        public string CICLO { get; set; }
        public string MONTO { get; set; }
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
            MONTO = string.Empty;
            APLICA_CDP = string.Empty;
            CAMPANA = string.Empty;
        }
    }
}
