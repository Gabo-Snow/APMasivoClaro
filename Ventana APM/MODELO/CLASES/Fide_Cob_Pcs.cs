using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Fide_Cob_Pcs
    {
        public string CuentaFinanciera { get; set; }
        public string recveivercustomer { get; set; }
        public string tipocargo { get; set; }
        public string codigodecargo { get; set; }
        public string offwer { get; set; }
        public string nombreOFFER { get; set; }
        public string promo { get; set; }
        public string description { get; set; }
        public double montoCargos { get; set; }
        public double montoDescuentos { get; set; }
        public double montoTotal { get; set; }
        public string tipodeCobro { get; set; }
        public string PCS { get; set; }
        public string secuenciadeCiclo { get; set; }
        public string prorrateo { get; set; }
        public string segmento { get; set; }
        public string rut { get; set; }
        public string TIPO { get; set; }
        public string DATO_MONTO { get; set; }
        public string MODALIDAD { get; set; }
        public double COBRAR { get; set; }
        public string COBRAR_SI { get; set; }
        public string CUOTA { get; set; }
        public string verdaredo_falso { get; set; }

        public Fide_Cob_Pcs()
        {
            this.Init();
        }



        private void Init()
        {
            CuentaFinanciera = string.Empty;
            recveivercustomer = string.Empty;
            tipocargo = string.Empty;
            codigodecargo = string.Empty;
            offwer = string.Empty;
            nombreOFFER = string.Empty;
            promo = string.Empty;
            description = string.Empty;
            montoCargos = 0.0;
            montoDescuentos = 0.0;
            montoTotal = 0.0;
            tipodeCobro = string.Empty;
            PCS = string.Empty;
            secuenciadeCiclo = string.Empty;
            prorrateo = string.Empty;
            segmento = string.Empty;
            rut = string.Empty;
            TIPO = string.Empty;
            DATO_MONTO = string.Empty;
            MODALIDAD = string.Empty;
            COBRAR = 0.0;
            COBRAR_SI = string.Empty;
            CUOTA = string.Empty;
            verdaredo_falso = string.Empty;
        }
    }
}
