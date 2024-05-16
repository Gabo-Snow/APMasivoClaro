using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;

namespace Ventana_APM.MODELO.COLECCIONES
{
   public class Coleccion_Cartera_Empresarial
    {
        string fileName = string.Empty;
        public Coleccion_Cartera_Empresarial()
        {

        }

        public List<Cartera_Empresarial> GenerarListado(string aux)
        {
            List<numeros_raros> numeros_Raros = new List<numeros_raros>();
            List<Cartera_Empresarial> carteras = new List<Cartera_Empresarial>();
            List<Cartera_Empresarial> carteras_sin_guion = new List<Cartera_Empresarial>();
            int contador_bano = 0;
            int contador_rut = 0;
            int contador_segmento = 0;
            int contador_lineas = 0;
            try
            {
                fileName = aux;
                Console.WriteLine(fileName);
                Console.WriteLine("Cargando...");
                //fileName = fbd.SelectedPath; // primer archivo

                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                    {
                        WorkbookPart workbookPart = doc.WorkbookPart;
                        SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                        SharedStringTable sst = sstpart.SharedStringTable;

                        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                        Worksheet sheet = worksheetPart.Worksheet;

                        var cells = sheet.Descendants<Cell>();
                        var rows = sheet.Descendants<Row>();
                        

                        //Console.WriteLine("Row count = {0}", rows.LongCount());
                        //Console.WriteLine("Cell count = {0}", cells.LongCount());
                        // Or... via each row
                        //int contadorLista = 0;
                        //int contador_de_columnas = 0;
                        
                        foreach (Row row in rows)
                        {
                            Cartera_Empresarial cartera = new Cartera_Empresarial();
                            //contador_de_columnas++;
                            int contador = 0;

                            if (contador_lineas == 0)
                            {
                                int contador_encabezado = 0;
                                foreach (var item in row.Elements<Cell>())
                                {
                                    contador_encabezado++;
                                    int ssidd = int.Parse(item.CellValue.Text);
                                    string strr = sst.ChildElements[ssidd].InnerText;
                                    string igual = strr.ToUpper();
                                    if (igual.Equals("SEGMENTO"))
                                    {
                                        contador_segmento = contador_encabezado;
                                        //contador++;

                                    }
                                    else if (igual.Equals("RUT"))
                                    {
                                        contador_rut = contador_encabezado;
                                    }
                                }
                                contador_lineas++;
                                continue;//si falla aqui que detenga el proceso
                            }

                            
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                contador++;

                                
                                if(contador_segmento >= 1 && contador_rut >= 1)//este if no deberia existir
                                {
                                    try
                                    {
                                        if ((c.CellValue != null) && (contador == contador_rut) )//&& (c.DataType == CellValues.SharedString)
                                        {
                                            if (c.CellValue.Text.Length < 7)
                                            {
                                                numeros_raros numeros_Raros1 = new numeros_raros();
                                                int ssid = int.Parse(c.CellValue.Text);
                                                //numeros_Raros1.numero_rut = int.Parse(c.CellValue.Text);
                                                //numeros_Raros1.numero_columna = contador_de_columnas;
                                                //numeros_Raros.Add(numeros_Raros1);
                                                string str = sst.ChildElements[ssid].InnerText;
                                                string[] col = str.Split(new char[] { '-' });
                                                if (col.Count() == 1 )
                                                {
                                                    cartera.RUT = col[0];
                                                }
                                                else
                                                {
                                                    cartera.RUT = col[0] + col[1];
                                                }
                                                
                                            }
                                            else
                                            {
                                                cartera.RUT = c.CellValue.Text;
                                            }

                                        }
                                        else if ((c.CellValue != null) && (contador == contador_segmento) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            cartera.SEGMENTO = str.ToUpper();
                                        }
                                        else if ((c.CellValue != null) && (contador == 99))
                                        {
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        if ((c.CellValue != null) && (contador == contador_rut))
                                        {
                                            cartera.RUT = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == 99))
                                        {
                                            break;
                                        }
                                    }

                                }
                            }
                            if (cartera.RUT.Equals(""))
                            {

                            }
                            else
                            {
                                carteras.Add(cartera);
                            }

                        }
                        //  Console.ReadLine();
                    }



                }

                foreach (var item in carteras)
                {
                    item.RUT = Regex.Replace(item.RUT, "-", "");
                    carteras_sin_guion.Add(item);
                }

                int rarisimo = numeros_Raros.Count();
                return carteras_sin_guion;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falta la cartera empresarial");
                return carteras_sin_guion;
            }




        }
    }

    class numeros_raros
    {
        public int numero_rut = 0;
        public int numero_columna = 0;
    }
}
