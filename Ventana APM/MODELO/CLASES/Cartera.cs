using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ventana_APM.MODELO.CLASES
{
    public class Cartera
    {
        public string L7_RUT_ID_VALUE { get; set; }
        public string BA_NO { get; set; }
        public string SEGMENTO { get; set; }

        public Cartera()
        {
            this.Init();
        }

        //como mierda agregamos la otra cartera intedezante =/
        //mira weon no se te ocurre nada estamso estancados, corta.
        //Mmmmmmmmmmmmmm realmente creo que ;


        private void Init()
        {
            L7_RUT_ID_VALUE = string.Empty;
            BA_NO = string.Empty;
            SEGMENTO = string.Empty;
        }
    }
}
