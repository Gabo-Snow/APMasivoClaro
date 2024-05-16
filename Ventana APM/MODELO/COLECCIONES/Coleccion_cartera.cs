using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public class Coleccion_cartera
    {
        string fileName = string.Empty;
        public Coleccion_cartera()
        {

        }

        public List<Cartera> GenerarListado(string aux)
        {
            fileName = aux;
            List<Cartera> carteras = new List<Cartera>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando...");
            int contador_de_columnas = 0;
            int contador_bano = 0;
            int contador_rut = 0;
            int contador_segmento = 0;

            //fileName = fbd.SelectedPath; // primer archivo
            try
            {
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
                        
                        
                        foreach (Row row in rows)
                        {
                            contador_de_columnas++;
                            Cartera cartera = new Cartera();
                            int contador = 0;
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                contador++;
                                if (contador_segmento >= 1 && contador_rut >= 1 && contador_bano >= 1)
                                {
                                    try
                                    {
                                        if ((c.CellValue != null) && (contador == contador_rut))
                                        {
                                            // cartera.L7_RUT_ID_VALUE = c.CellValue.Text;
                                            if (c.CellValue.Text.Length < 7)
                                            {

                                                int ssid = int.Parse(c.CellValue.Text);
                                                string str = sst.ChildElements[ssid].InnerText;
                                                string[] col = str.Split(new char[] { '-' });


                                                if (col.Count() == 1)
                                                {
                                                    cartera.L7_RUT_ID_VALUE = col[0];
                                                }
                                                else
                                                {
                                                    cartera.L7_RUT_ID_VALUE = col[0] + col[1];
                                                }
                                            }
                                            else
                                            {
                                                cartera.L7_RUT_ID_VALUE = c.CellValue.Text;
                                            }

                                        }
                                        else if ((c.CellValue != null) && (contador == contador_bano))
                                        {
                                            cartera.BA_NO = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_segmento) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            cartera.SEGMENTO = str.ToUpper();
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        cartera.L7_RUT_ID_VALUE = c.CellValue.Text;
                                        Console.WriteLine(ex);
                                    }

                                }
                                else
                                {

                                    try
                                    {
                                        int ssidd = int.Parse(c.CellValue.Text);
                                        string strr = sst.ChildElements[ssidd].InnerText;
                                        string igual = strr.ToUpper();
                                        if (igual.Equals("SEGMENTO"))
                                        {
                                            contador_segmento = contador;
                                            contador++;

                                        }
                                        else if (igual.Equals("L7_RUT_ID_VALUE"))
                                        {
                                            contador_rut = contador;
                                        }
                                        else if (igual.Equals("BA_NO"))
                                        {
                                            contador_bano = contador;
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                        continue;
                                    }
                                }

                            }
                            if (cartera.BA_NO.Equals(""))
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
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex);
            }

           
            return carteras;



        }
    }
}
