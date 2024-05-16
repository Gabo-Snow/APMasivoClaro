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
    public class Coleccion_Paramount
    {
        string fileName = string.Empty;
        public Coleccion_Paramount()
        {

        }

        public List<Paramount> GenerarListado(string aux)
        {
            List<Paramount> coleccion_Pra = new List<Paramount>();
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                
                fileName = aux;

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
                        Paramount pra = new Paramount();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter >= 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            pra.CuentaFinanciera = col[0];
                            pra.RECEIVER_CUSTOMER = col[1];
                            pra.TIPO_CARGO  = col[2];
                            pra.Codigo_de_Cargo  = col[3];
                            pra.OFFER  = col[4];
                            pra.Nombre_OFFER = col[5];
                            pra.PROMO  = col[6];
                            pra.DESCRIPTION = col[7];
                            pra.Tipo_de_Cobro = col[11];
                            pra.PCS  = col[12];
                            pra.Secuencia_de_Ciclo = col[13];
                            pra.PRORRATEO = col[14];
                            //todos los computadores tieneRUT = string.Empty;n distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                pra.MONTO_CARGOS = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                pra.MONTO_CARGOS = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                pra.MONTO_DESCUENTOS = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                pra.MONTO_DESCUENTOS = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                pra.MONTO_TOTAL = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                pra.MONTO_TOTAL = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            coleccion_Pra.Add(pra);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_Pra;
                }

            }
            
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return coleccion_Pra;//nose si eto funca pium pukm
            }



        }
    }
}
