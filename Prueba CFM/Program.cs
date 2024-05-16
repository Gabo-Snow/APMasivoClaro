using System;
using System.Collections.Generic;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;
using APM;
using Pruebas.MODELO.CLASES;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Prueba_CFM
{
    class Program
    {
        static void Main(string[] args)
        {
            List<FVM> fVMs = new List<FVM>();
            Coleccion_FVM coleccion_FVM = new Coleccion_FVM();
            fVMs = coleccion_FVM.GenerarListado(@"C:\Users\Marcelo\Desktop\Ciclo 18 entrada\CICLOS FVM.xlsx");
        }
    }
}
