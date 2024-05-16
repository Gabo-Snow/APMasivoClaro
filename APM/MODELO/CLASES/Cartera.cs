using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APM.MODELO.CLASES
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



        private void Init()
        {
            L7_RUT_ID_VALUE = string.Empty;
            BA_NO = string.Empty;
            SEGMENTO = string.Empty;
        }
    }
}
