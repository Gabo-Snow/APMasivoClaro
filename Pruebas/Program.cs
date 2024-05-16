using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ventana_APM.MODELO.COLECCIONES;
using System.Data;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Ventana_APM.MODELO.CLASES;
using APM;
using System.Collections;
using System.Diagnostics;
using Spire.Xls;
using Ventana_APM.CONTROLADOR;

namespace Pruebas
{
    class Program
    {
        static void Main(string[] args)
        {

            //desde ahora todo lo venga en excel sera leido de esta forma//
            // se transformara en txt y de txt pasara a ser una coleccion e dicho caso cerrao
            //de esta forma si pueden existir columnas en blanco
            Coleccion_CargosxPcs coleccion_CargosxPcs = new Coleccion_CargosxPcs();
            coleccion_CargosxPcs.GenerarListado(@"D:\marcelo.cona\Desktop\Ciclo 16 2021\CICLO 16\cargosxpcs2.txt");
            Controlador_Hbo_Paramount controlador_Hbo_Paramount = new Controlador_Hbo_Paramount(@"D:\marcelo.cona\Desktop\Paramount y Hbo\HBO 14.12.txt", "path paramount", @"D:\marcelo.cona\Desktop\Paramount y Hbo");

            controlador_Hbo_Paramount.hbo_31();

            //hagamos que se suba
            int currentMonth = DateTime.Now.Month;
            int currentYear = DateTime.Now.Year;
            int days = DateTime.DaysInMonth(currentYear -1, 2);
            //double valor = Math.Round(6900 / 1.19);


            Workbook workbook = new Workbook();

            workbook.LoadFromFile(@"D:\marcelo.cona\Desktop\PruebaExcelToTxt\Detenciones ciclo 18-05-20212.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            sheet.SaveToFile(@"D:\marcelo.cona\Desktop\PruebaExcelToTxt\ExceltoTxt3.txt", "|", Encoding.UTF8);

            
            int numero = 10;

            List<int> guardando_i = new List<int>();

            //como flores reparaste mi triztesa con colores turururuun
            int[,] spiderverse;

            spiderverse = new int[,] { { 1, 2 }, { 3, 4 }, { 5, 6 }, { 7, 8 }, { 7, 8 }, { 7, 8 }, { 7, 8 }, { 7, 8 }, { 7, 8 }, { 7, 8 }, { 7, 8 } };
            // gravity tururu lara larara para bailar la bamba, para bailar la bamba se necesita una poca de gracia
            //
            List<bool> sino = new List<bool>();

            for (int i = 0; i < numero; i++)
            {
                Process currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                Process proc = Process.GetCurrentProcess();
                long totalBytesOfMemoryUsed = currentProcess.WorkingSet64;
                LinkedList<double> vs = new LinkedList<double>();
                numero++;
                Console.WriteLine("Almacenando Mb: {0}----------------------------------", totalBytesOfMemoryUsed/ 1000000);
                //Console.WriteLine(i);
                Console.WriteLine("----------------------------------------------------"+i);
                Console.WriteLine("Memoria : {0}----------------------------------------", proc.PrivateMemorySize64/ 1000000);
                guardando_i.Add(i);
            }


            Coleccion_Consolidacion_Validacion coleccion = new Coleccion_Consolidacion_Validacion();

            coleccion.GenerarListado(@"D:\marcelo.cona\Desktop\Cargos de Activacion pruebas\Consolidado.xlsx");

            Console.WriteLine("si temine ._. fire emblem");
            
             
            Console.WriteLine("TERMINE!");

            Console.ReadLine();





        }
    }
}
