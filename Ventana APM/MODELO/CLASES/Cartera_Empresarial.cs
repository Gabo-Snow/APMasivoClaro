using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Cartera_Empresarial
    {
        public string RUT { get; set; }
        public string SEGMENTO { get; set; }

        public Cartera_Empresarial()
        {
            this.Init();
        }


        private void Init()
        {
            RUT = string.Empty;
            SEGMENTO = string.Empty;
        }
    }
}
