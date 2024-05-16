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
    public class Coleccion_c_ar
    {
        public Coleccion_c_ar()
        {

        }

        public List<Cargos_Arrendamiento> GenerarListado(string aux)
        {
            try
            {
                string fileName = string.Empty;
                List<Cargos_Arrendamiento> coleccion_C_Ars = new List<Cargos_Arrendamiento>();
                Console.WriteLine("Almacenando Cargos");
                fileName = aux + @"\ENTRADA\CARGOS.txt";
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
                        Cargos_Arrendamiento c_Ar = new Cargos_Arrendamiento();
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
                            c_Ar.rut = col[15];
                            coleccion_C_Ars.Add(c_Ar);
                            Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("El archivo CARGOS_ARRENDAMIENTO esta erroneo comprobar");
                MessageBox.Show("El nobmre del archivo debe ser CARGOS.txt");
                MessageBox.Show("Ingresar archivos correctos");
                MessageBox.Show("El error es {0}",ex.ToString());
                Application.Exit();
                return null;
            }



        }
        

    }
}

