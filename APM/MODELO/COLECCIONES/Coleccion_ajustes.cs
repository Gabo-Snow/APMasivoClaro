using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
   public class Coleccion_ajustes
    {
        string fileName = string.Empty;
        public Coleccion_ajustes()
        {

        }

        public List<Ajustes> GenerarListado(string aux)
        {
            List<Ajustes> ajustes = new List<Ajustes>();

                Console.WriteLine("Cargando...");
                //fileName = fbd.SelectedPath; // primer archivo
                fileName = aux + @"\ENTRADA\C_ARRENDAMIENTO AJUSTES ESPECIALES.xlsx";
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
                            Ajustes ajuste = new Ajustes();
                            int contador = 0;
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                contador++;

                                if ((c.CellValue != null) && (contador == 1))
                                {
                                    ajuste.PCS = c.CellValue.Text;
                                }
                                else if ((c.CellValue != null) && (contador == 2))
                                {
                                    ajuste.CuentaFinanciera = c.CellValue.Text;
                                }
                                else if ((c.CellValue != null) && (contador == 3) && (c.DataType == CellValues.SharedString))
                                {
                                    int ssid = int.Parse(c.CellValue.Text);
                                    string str = sst.ChildElements[ssid].InnerText;
                                    ajuste.codigodecargo = str;
                                }
                                else if ((c.CellValue != null) && (contador == 4))
                                {
                                    ajuste.amount = double.Parse(c.CellValue.Text);
                                }
                                else if ((c.CellValue != null) && (contador == 5) && (c.DataType == CellValues.SharedString))
                                {
                                    int ssid = int.Parse(c.CellValue.Text);
                                    string str = sst.ChildElements[ssid].InnerText;
                                    ajuste.ciclo = str;
                                }

                            }

                            ajustes.Add(ajuste);
                        Console.WriteLine("Almacenando ajustes {0}",ajustes.Count());
                        }
                        //  Console.ReadLine();
                    }
                }


            
            return ajustes;

        }
    }
}
