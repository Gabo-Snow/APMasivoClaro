using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Cuentas_Asignadas
    {
        public string PCS { get; set; }

        public Cuentas_Asignadas()
        {
            this.Init();
        }



        private void Init()
        {
            PCS = string.Empty;
        }
    }
}
