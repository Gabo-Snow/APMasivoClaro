using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APM.MODELO.CLASES
{
    public class Asignados_arr
    {

        


        public string ASIGNACIÓN { get; set; }
        public string Cuenta_Financiera { get; set; }
        public string RECEIVER_CUSTOMER { get; set; }
        public string TIPO_CARGO { get; set; }
        public string Codigo_de_Cargo { get; set; }
        public string OFFER { get; set; }
        public string Nombre_OFFER { get; set; }
        public string PROMO { get; set; }
        public string DESCRIPTION { get; set; }
        public double MONTO_CARGOS { get; set; }
        public double MONTO_DESCUENTOS { get; set; }
        public double MONTO_TOTAL { get; set; }
        public string Tipo_de_Cobro { get; set; }
        public string PCS { get; set; }
        public string Secuencia_de_Ciclo { get; set; }
        public string COBRAR { get; set; }
        public double MONTO { get; set; }
        public string CUOTA { get; set; }
        public string OBSERVACIÓN { get; set; }

        public Asignados_arr()
        {
            this.Init();
        }

        private void Init()
        {
            ASIGNACIÓN = string.Empty;
            Cuenta_Financiera = string.Empty;
            RECEIVER_CUSTOMER = string.Empty;
            TIPO_CARGO = string.Empty;
            Codigo_de_Cargo = string.Empty;
            OFFER = string.Empty;
            Nombre_OFFER = string.Empty;
            PROMO = string.Empty;
            DESCRIPTION = string.Empty;
            MONTO_CARGOS = 0.0;
            MONTO_DESCUENTOS = 0.0;
            MONTO_TOTAL = 0.0;
            Tipo_de_Cobro = string.Empty;
            PCS = string.Empty;
            Secuencia_de_Ciclo = string.Empty;
            COBRAR = string.Empty;
            MONTO = 0.0;
            CUOTA = string.Empty;
            OBSERVACIÓN = string.Empty;
        }
    }
}