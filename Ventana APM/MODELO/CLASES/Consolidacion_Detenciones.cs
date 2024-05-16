using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pruebas.MODELO.CLASES
{
    public class Consolidacion_Detenciones
    {
        public string CuentaFinanciera { get; set; }
        public string CicloFacturacion { get; set; }
        public string AcccionRealizar { get; set; }
        public string Observacion { get; set; }

        public Consolidacion_Detenciones()
        {
            this.Init();
        }



        private void Init()
        {
            CuentaFinanciera = string.Empty;
            CicloFacturacion = string.Empty;
            AcccionRealizar = string.Empty;
            Observacion     = string.Empty;
        }

    }
}
