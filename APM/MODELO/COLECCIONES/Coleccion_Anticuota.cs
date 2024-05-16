using APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APM.MODELO.COLECCIONES
{
   public class Coleccion_Anticuota
    {
        public Coleccion_Anticuota()
        {

        }

        public List<Anticuota> GenerarListado(string aux)
        {
            try
            {
                string fileName = string.Empty;
                List<Anticuota> coleccion_C_Ars = new List<Anticuota>();
                Console.WriteLine("Almacenando Anticuota");
                fileName = aux + @"\ENTRADA\CARGOS_ACTICUOTA.txt";
                Console.WriteLine(aux);

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Anticuota c_Ar = new Anticuota();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            c_Ar.montoCargos = double.Parse(col[8]);
                            c_Ar.montoDescuentos = double.Parse(col[9]);
                            c_Ar.montoTotal = double.Parse(col[10]);
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.prorrateo = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ingresar archivos correctos");
                MessageBox.Show("El error es {0}", ex.ToString());
                Application.Exit();
                return null;
            }



        }
    }
}
