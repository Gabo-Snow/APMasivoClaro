using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Consolidacion_Validacion
    {
        public string pcs { get; set; }
        public string monto_validador { get; set; }
        public string monto_ajuste { get; set; }
        public bool hay_ajuste { get; set; }
        public string codigo { get; set; }

        public Consolidacion_Validacion()
        {
            this.Init();
        }
        private void Init()
        {
            pcs = string.Empty;
            monto_validador = string.Empty;
            monto_ajuste = string.Empty;
            hay_ajuste = false;
            codigo = string.Empty;
        }

    }
}
