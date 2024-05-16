using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ventana_APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace Ventana_APM.MODELO.COLECCIONES
{
    public class Coleccion_FVM
    {
        string fileName = string.Empty;
        public Coleccion_FVM()
        {

        }
        public List<FVM> GenerarListado(string aux)
        {
            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            fileName = aux;
            List<FVM> fVMs = new List<FVM>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando FVM...");
            //fileName = fbd.SelectedPath; // primer archivo
            int contador_primera_linea = 0;
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
                            contador_primera_linea++;
                            FVM vM = new FVM();
                            int contador = 0;
                            if (contador_primera_linea == 1)
                            {

                            }
                            else
                            {
                                foreach (Cell c in row.Elements<Cell>())
                                {
                                    contador++;

                                    try {

                                        if ((c.CellValue != null) && (contador == 1) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            vM.CCOCSUBEQPCOM = str;
                                        }
                                        else if ((c.CellValue != null) && (contador == 2))
                                        {
                                            int tiene_punto = c.CellValue.Text.IndexOf('.');
                                            if (tiene_punto > 1)
                                            {
                                                vM.MONTO_SIN_IMPUESTO = double.Parse(c.CellValue.Text, culture_punto);
                                            }
                                            else
                                            {
                                                vM.MONTO_SIN_IMPUESTO = double.Parse(c.CellValue.Text, culture_coma);
                                            }
                                        }
                                        else if ((c.CellValue != null) && (contador == 3))
                                        {
                                            vM.PCS = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == 4) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            vM.INGRESO_VTA = str;
                                        }
                                        else if ((c.CellValue != null) && (contador == 5))
                                        {
                                            vM.FA = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == 6))
                                        {
                                            vM.CBP = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == 7))
                                        {
                                            try
                                            {
                                                double d = double.Parse(c.CellValue.Text);
                                                DateTime conv = DateTime.FromOADate(d);
                                                vM.FECHA_ACTIVACION = conv.Day + "-" + conv.Month + "-" + conv.Year;
                                            }
                                            catch (Exception ex)
                                            {
                                                if ((c.CellValue != null) && (contador == 7))
                                                {
                                                    string[] col = c.CellValue.Text.Split(new char[] { '.' });
                                                    double d = double.Parse(col[0]);
                                                    DateTime conv = DateTime.FromOADate(d);
                                                    vM.FECHA_ACTIVACION = conv.Day + "-" + conv.Month + "-" + conv.Year;
                                                }
                                            }


                                        }
                                    } catch (Exception e)
                                    {
                                        continue;
                                    }
                                    }

                                fVMs.Add(vM);
                            }

                        }
                        //  Console.ReadLine();
                    }



                }
                return fVMs;
            }
            catch (Exception ex)
            {
                
                return null;
            }




        }
    }
}
