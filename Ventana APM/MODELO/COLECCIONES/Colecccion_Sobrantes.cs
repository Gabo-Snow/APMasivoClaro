using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;

namespace Ventana_APM.MODELO.COLECCIONES
{
    public class Colecccion_Sobrantes
    {
        public Colecccion_Sobrantes()
        {

        }

        public List<PCS_SOBRANTES> GenerarListado(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<PCS_SOBRANTES> coleccion_C_Ars = new List<PCS_SOBRANTES>();
                fileName = @"C:\Users\Marcelo\Desktop\Ciclo 6\PCS_SOBRANTES.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        PCS_SOBRANTES c_Ar = new PCS_SOBRANTES();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            c_Ar.PCS = col[0];

                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ingresar archivos correctos");
                MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
    }
}
