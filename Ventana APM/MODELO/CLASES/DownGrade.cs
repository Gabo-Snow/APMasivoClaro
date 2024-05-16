using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class DownGrade
    {
        public string RUT { get; set; }
        public string ID_ORDEN_ACT { get; set; }
        public string ID_ORDEN { get; set; }
        public string PLAN_ANTERIOR { get; set; }
        public string PLAN_ACTUAL { get; set; }
        public string ID { get; set; }
        public string TIPO_DE_ORDEN { get; set; }
        public string ID_USUARIO { get; set; }
        public string TIPO_USER { get; set; }
        public string TIPO_CLIENTE { get; set; }
        public string SUCURSAL { get; set; }
        public string ESTADO_ID_ORDEN_REFERENCE { get; set; }
        public string ESTADO_ORDEN_ACT { get; set; }
        public string PCS { get; set; }
        public string CUENTA { get; set; }
        public string CICLO { get; set; }

        public DownGrade()
        {
            this.Init();
        }



        private void Init()
        {
            RUT = string.Empty;
            ID_ORDEN_ACT = string.Empty;
            ID_ORDEN = string.Empty;
            PLAN_ANTERIOR = string.Empty;
            PLAN_ACTUAL = string.Empty;
            ID = string.Empty;
            TIPO_DE_ORDEN = string.Empty;
            ID_USUARIO = string.Empty;
            TIPO_USER = string.Empty;
            TIPO_CLIENTE = string.Empty;
            SUCURSAL = string.Empty;
            ESTADO_ID_ORDEN_REFERENCE = string.Empty;
            ESTADO_ORDEN_ACT = string.Empty;
            PCS = string.Empty;
            CUENTA = string.Empty;
            CICLO = string.Empty;

        }

    }
}
