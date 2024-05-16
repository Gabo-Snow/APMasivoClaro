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
    public class Coleccion_Fide_Cob_Pcs
    {
        public Coleccion_Fide_Cob_Pcs()
        {

        }

        public List<Fide_Cob_Pcs> GenerarListado(string aux)
        {
            int counter = 0;
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Fide_Cob_Pcs> coleccion_C_Ars = new List<Fide_Cob_Pcs>();
                fileName = aux;

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))//agregar encoding default a todo
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Fide_Cob_Pcs c_Ar = new Fide_Cob_Pcs();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter >= 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = string.Format(col[2], culture_coma);
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = string.Format(col[7], culture_coma);
                            if (tiene_punto_monto_cargos >= 0)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                bool success = double.TryParse(col[8], out double numero);
                                if (success)
                                {
                                    c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                                }
                                else
                                {
                                    c_Ar.montoCargos = double.Parse("0" + col[8], culture_coma);
                                    MessageBox.Show("0" + col[8]);
                                }                                
                            }
                            if (tiene_punto_montoDescuentos >= 0)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                bool success = double.TryParse(col[9], out double numero);
                                if (success)
                                {
                                    c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                                }
                                else
                                {
                                    c_Ar.montoDescuentos = double.Parse("0" + col[9], culture_coma);
                                }
                                
                            }
                            if (tiene_punto_montoTotal >=0 )
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                bool success = double.TryParse(col[10], out double numero);
                                if (success)
                                {
                                    c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                                }
                                else
                                {
                                    c_Ar.montoTotal = double.Parse("0" + col[10], culture_coma);
                                }
                               
                            }
                            //---------------------------------------------------------------------------------------------------------------------

                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            //c_Ar.prorrateo = col[14];
                            coleccion_C_Ars.Add(c_Ar);

                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                int algo = counter;
                MessageBox.Show("Ingresar archivos correctos");
                MessageBox.Show(ex.ToString());                
                MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
    }
}
