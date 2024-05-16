using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.Auxiliares
{
    public class Aux_Downgrade
    {
            

        public string PCS { get; set; }
        public string CUENTA { get; set; }
        public string CARGO { get; set; }
        public double AMOUNT { get; set; }

        public Aux_Downgrade()
        {
            this.Init();
        }



        private void Init()
        {
            PCS = string.Empty;
            CUENTA = string.Empty;
            CARGO = string.Empty;
            AMOUNT = 0.0;
        }
    }
}
