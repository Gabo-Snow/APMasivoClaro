using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public  class PCS_SOBRANTES
    {
        public string PCS { get; set; }
        public PCS_SOBRANTES()
        {
            this.Init();
        }



        private void Init()
        {
     
            PCS = string.Empty;
        }
    }
}
