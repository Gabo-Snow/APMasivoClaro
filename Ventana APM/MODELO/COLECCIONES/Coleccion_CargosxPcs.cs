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
    public class Coleccion_CargosxPcs
    {
        public Coleccion_CargosxPcs()
        {

        }

        public List<CargosxPcs> GenerarListado(string aux)
        {
            int counter = 0;
            try
            {
                string fileName = string.Empty;
                List<CargosxPcs> coleccion_C_Ars = new List<CargosxPcs>();
                fileName = aux;
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        CargosxPcs c_Ar = new CargosxPcs();
                        string[] col = line.Split(new char[] { '|' });
                        int contador_de_col = col.Count();
                        if (counter == 1102616)
                        {
                            int algo = contador_de_col;
                        }
                        if (counter >= 1)
                        {
                            if (contador_de_col >= 1)
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
                                //todos los computadores tienen distintas configuraciones regionales----------------------------------------------------- hay que poner >= 1 ya que cuando es , es -1
                                if (tiene_punto_monto_cargos >= 0)
                                {
                                    c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                                }
                                else
                                {
                                    c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                                }
                                if (tiene_punto_montoDescuentos >= 0)
                                {
                                    c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                                }
                                else
                                {
                                    c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                                }
                                if (tiene_punto_montoTotal >= 0)
                                {
                                    c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                                }
                                else
                                {
                                    c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                                }
                                //---------------------------------------------------------------------------------------------------------------------
                                c_Ar.tipodeCobro = col[11];
                                c_Ar.PCS = col[12];
                                c_Ar.secuenciadeCiclo = col[13];
                                c_Ar.prorrateo = col[14];
                                if (c_Ar.tipocargo.Equals("Cargos por Tráfico Adicional"))
                                {
                                    coleccion_C_Ars.Add(c_Ar);
                                }
                                else if (c_Ar.tipocargo.Equals("Cargos por TrÃ¡fico Adicional"))
                                {
                                    coleccion_C_Ars.Add(c_Ar);
                                }
                                else if (c_Ar.tipocargo.Equals("Cargos por Tr?fico Adicional"))
                                {
                                    coleccion_C_Ars.Add(c_Ar);
                                }
                            }
                            
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                int linea = counter;
                //MessageBox.Show("Ingresar archivos correctos");
                MessageBox.Show("ERROR: ", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
    }
}
