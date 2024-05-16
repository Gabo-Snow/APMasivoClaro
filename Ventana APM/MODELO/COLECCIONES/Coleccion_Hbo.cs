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
    public class Coleccion_Hbo
    {
        string fileName = string.Empty;
        public Coleccion_Hbo()
        {

        }

        public List<Hbo> GenerarListado(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Hbo> coleccion_C_hbo = new List<Hbo>();
                Console.WriteLine("Almacenando SimCard");
                fileName = @aux;
                Console.WriteLine(aux);

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Hbo c_Ar = new Hbo();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter >= 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');

                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.RECEIVER_CUSTOMER = col[1];
                            c_Ar.TIPO_CARGO = col[2];
                            c_Ar.Codigo_de_Cargo = col[3];
                            c_Ar.OFFER = col[4];
                            c_Ar.Nombre_OFFER = col[5];
                            c_Ar.PROMO = col[6];
                            c_Ar.DESCRIPTION = col[7];
                            c_Ar.Tipo_de_Cobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.Secuencia_de_Ciclo = col[13];
                            c_Ar.PRORRATEO = int.Parse(col[14]);
                            c_Ar.RUT = col[15];

                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.MONTO_CARGOS = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.MONTO_CARGOS = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.MONTO_DESCUENTOS = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.MONTO_DESCUENTOS = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.MONTO_TOTAL = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.MONTO_TOTAL = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------


                            coleccion_C_hbo.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_hbo;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
    }
}
