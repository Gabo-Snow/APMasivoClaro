using APM;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using Ventana_APM.MODELO.COLECCIONES;


namespace Cargos_de_Activacion
{
    class Program
    {
        static void Main(string[] args)
        {
            //Coleccion_Activacion coleccion_Activacion = new Coleccion_Activacion();
            //coleccion_Activacion.GenerarListado(@"D:\marcelo.cona\Desktop\Cargos de Activacion pruebas\Activación 08.08.txt");

            //Coleccion_Consolidacion_Validacion coleccion = new Coleccion_Consolidacion_Validacion();
            //coleccion.GenerarListado(@"D:\marcelo.cona\Desktop\Cargos de Activacion pruebas\Activación 08.08.txt");



            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");

            string numero_coma = "3000,4";
            string numero_punto = "3000.4";
            string numero_sin_punto_ni_coma = "30000";
            int tiene_punto_monto_cargos = numero_coma.IndexOf('.');
            int tiene_punto_montoDescuentos = numero_punto.IndexOf('.');
            int tiene_punto_montoTotal = numero_sin_punto_ni_coma.IndexOf('.');
            List<double> numeros = new List<double>();
            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
            if (tiene_punto_monto_cargos > 1)
            {
                numeros.Add(double.Parse(numero_coma, culture_punto));
            }
            else
            {
                numeros.Add(double.Parse(numero_coma, culture_coma));
            }
            if (tiene_punto_montoDescuentos > 1)
            {
                numeros.Add(double.Parse(numero_punto, culture_punto));
            }
            else
            {
                numeros.Add(double.Parse(numero_punto, culture_coma));
            }
            if (tiene_punto_montoTotal > 1)
            {
                numeros.Add(double.Parse(numero_sin_punto_ni_coma, culture_punto));
            }
            else
            {
                numeros.Add(double.Parse(numero_sin_punto_ni_coma, culture_coma));
            }
            //se me borro todo lo trabajado aaaaaaaaaaaaaaaah
            //---------------------------------------------------------------------------------------------------------------------
            //List<FVM> fVMs = new List<FVM>();
            //Coleccion_FVM coleccion_FVM = new Coleccion_FVM();
            //fVMs = coleccion_FVM.GenerarListado(@"C:\Users\Marcelo\Desktop\Ciclo 18 entrada\CICLOS FVM.xlsx");
            //Coleccion_Cartera_Empresarial coleccion_Cartera_Empresarial = new Coleccion_Cartera_Empresarial();
            //Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
            //List<Cartera_Empresarial> cartera_Empresarials = new List<Cartera_Empresarial>();

            //List<Cartera> carteras = new List<Cartera>();

            //carteras = coleccion_Cartera.GenerarListado(@"C:\Users\Marcelo\Desktop\Ciclo 18 entrada\Cartera 08.02.2021 _PABLO.xlsx");
            //cartera_Empresarials = coleccion_Cartera_Empresarial.GenerarListado(@"C:\Users\Marcelo\Desktop\Ciclo 18 entrada\Cartera actualizada 23-02-21.xlsx");
            int numeros_random = 0;
            
            foreach (var item in numeros)
            {
                Console.WriteLine("Numeros {0}",item);
            }

            //
            //


            Console.Beep();//sonidos para saber que termino
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();

            Console.WriteLine("termine"); 

        }
    }
}

