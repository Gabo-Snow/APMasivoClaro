using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APM.MODELO.COLECCIONES
{
    public class Coleccion_FVM
    {
        string fileName = string.Empty;
        public Coleccion_FVM()
        {

        }
        public List<FVM> GenerarListado(string aux)
        {
            fileName = aux + @"\ENTRADA\CICLOS FVM.xlsx";
            List<FVM> fVMs = new List<FVM>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando FVM...");
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
                    foreach (Row row in rows)
                    {
                        FVM vM = new FVM();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;

                            if ((c.CellValue != null) && (contador == 1) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                vM.CCOCSUBEQPCOM = str;
                            }
                            else if ((c.CellValue != null) && (contador == 2))
                            {
                                vM.MONTO_SIN_IMPUESTO = c.CellValue.Text;
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
                                vM.FECHA_ACTIVACION = c.CellValue.Text;
                            }
                        }

                        fVMs.Add(vM);
                    }
                    //  Console.ReadLine();
                }



            }
            return fVMs;



        }
    }
}
