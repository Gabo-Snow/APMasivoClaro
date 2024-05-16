using Ventana_APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ventana_APM.MODELO.COLECCIONES
{
   public class Coleccion_Portabilidad
    {
        public Coleccion_Portabilidad()
        {

        }

        public List<Portabilidad> GenerarListado(string aux)
        {
            try
            {
                string fileName = string.Empty;
                List<Portabilidad> portabilidads = new List<Portabilidad>();
                Console.WriteLine("Almacenando Portabilidad");
                fileName = aux;
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
                        Portabilidad portabilidad = new Portabilidad();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                           portabilidad.TIPO = col[0];
                           portabilidad.ENVIO_SOLICITUD = col[1];
                           portabilidad.NRO_DE_PCS = col[2];
                           portabilidad.TIPO_SERVICIO = col[3];
                            if (col[4].Equals("Prepago"))
                            {
                                portabilidad.MODALIDAD = "PRE A POST";
                            }
                            else if (col[4].Equals("Postpago"))
                            {
                                portabilidad.MODALIDAD = "PORTADO";
                            }
                            else
                            {
                                portabilidad.MODALIDAD = "N/A";
                            }
                           
                          // portabilidad.RECEPTOR = col[5];
    //                       portabilidad.DONANTE = col[6];
  //                         portabilidad.FECHA_PORTABILIDAD = DateTime.Parse(col[7]);
//                           portabilidad.SYSDATE = DateTime.Parse(col[8]);

                            portabilidads.Add(portabilidad);
                        }
                    }
                    return portabilidads;
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
