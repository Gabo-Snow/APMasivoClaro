using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APM.MODELO.CLASES
{
   public class FVM
    {
               
        public string CCOCSUBEQPCOM { get; set; }
        public string MONTO_SIN_IMPUESTO { get; set; }
        public string PCS { get; set; }
        public string INGRESO_VTA { get; set; }
        public string FA { get; set; }
        public string CBP { get; set; }
        public string FECHA_ACTIVACION { get; set; }
        public FVM()
        {
            this.Init();
        }

        private void Init()
        {
            CCOCSUBEQPCOM = string.Empty;
            MONTO_SIN_IMPUESTO = string.Empty;
            PCS = string.Empty;
            INGRESO_VTA = string.Empty;
            FA = string.Empty;
            CBP = string.Empty;
            FECHA_ACTIVACION = string.Empty;
        }
    }
}
