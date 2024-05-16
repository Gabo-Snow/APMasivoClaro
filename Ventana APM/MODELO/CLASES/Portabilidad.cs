using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
   public class Portabilidad
    {

        public string TIPO { get; set; }
        public string ENVIO_SOLICITUD { get; set; }
        public string NRO_DE_PCS { get; set; }
        public string TIPO_SERVICIO { get; set; }
        public string MODALIDAD { get; set; }
        public string RECEPTOR { get; set; }
        public string DONANTE { get; set; }
        public DateTime FECHA_PORTABILIDAD { get; set; }
        public DateTime SYSDATE { get; set; }

        public Portabilidad()
        {
            this.Init();
        }



        private void Init()
        {
            TIPO = string.Empty;
            ENVIO_SOLICITUD = string.Empty;
            NRO_DE_PCS = string.Empty;
            TIPO_SERVICIO = string.Empty;
            MODALIDAD = string.Empty;
            RECEPTOR = string.Empty;
            DONANTE = string.Empty;
            FECHA_PORTABILIDAD = DateTime.Now;
            SYSDATE = DateTime.Now;
        }
    }
}
