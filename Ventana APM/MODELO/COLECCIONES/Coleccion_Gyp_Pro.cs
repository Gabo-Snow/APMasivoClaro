using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ventana_APM.MODELO.CLASES;

namespace Ventana_APM.MODELO.COLECCIONES
{
    public class Coleccion_Gyp_Pro
    {
        string fileName = string.Empty;
        public Coleccion_Gyp_Pro()
        {

        }

        public List<Gyp_Pro> GenerarListado(string aux)
        {
            fileName = aux;
            List<Gyp_Pro> gyp_s = new List<Gyp_Pro>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando...");
            int contador_de_columnas = 0;
            int contador_cuenta = 0;

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
                            Gyp_Pro gyp = new Gyp_Pro();
                            int contador = 0;
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                contador++;
                                if (contador_cuenta >= 1)
                                {

                                    try
                                    {
                                        if ((c.CellValue != null) && (contador == contador_cuenta))
                                        {
                                            gyp.cuenta = c.CellValue.Text;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        continue;
                                    }

                                }
                                else
                                {
                                    try
                                    {
                                        int ssidd = int.Parse(c.CellValue.Text);
                                        string strr = sst.ChildElements[ssidd].InnerText;
                                        string igual = strr.ToUpper();
                                        if (igual.Equals("CUENTA"))
                                        {
                                            contador_cuenta = contador;
                                            contador++;

                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                        continue;
                                    }
                                }
                                try
                                {


                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex);
                                }



                            }
                            if (gyp.cuenta.Equals(""))
                            {

                            }
                            else
                            {
                                gyp_s.Add(gyp);
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


            return gyp_s;



        }
    }
}
