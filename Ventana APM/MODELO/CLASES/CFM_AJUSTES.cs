using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class CFM_AJUSTES
    {
        public string PCS { get; set; }
        public string CUENTA { get; set; }
        public string CICLO { get; set; }
        public double MONTO { get; set; }
        public double CF_MAS_IVA { get; set; }
        public double CARGOS { get; set; }
        public double AJUSTE { get; set; }
        public string APLICA_CDP { get; set; }
        public string CAMPANA { get; set; }
        public CFM_AJUSTES()
        {
            this.Init();
        }

        private void Init()
        {
            PCS = string.Empty;
            CUENTA = string.Empty;
            CICLO = string.Empty;
            MONTO = 0.0;
            CF_MAS_IVA = 0.0;
            CARGOS = 0.0;
            AJUSTE = 0.0;
            APLICA_CDP = string.Empty;
            CAMPANA = string.Empty;

        }

    }
}
