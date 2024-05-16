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
    public class Coleccion_cartera
    {
        string fileName = string.Empty;
        public Coleccion_cartera()
        {

        }

        public List<Cartera> GenerarListado(string aux)
        {
            fileName = aux + @"\ENTRADA\Cartera.xlsx";
            List<Cartera> carteras = new List<Cartera>();
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
                    foreach (Row row in rows)
                    {
                        Cartera cartera = new Cartera();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;

                            if ((c.CellValue != null) && (contador == 1))
                            {
                                cartera.L7_RUT_ID_VALUE = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 2))
                            {
                                cartera.BA_NO = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 3) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                cartera.SEGMENTO = str;
                            }
                        }

                        carteras.Add(cartera);
                    }
                    //  Console.ReadLine();
                }



            }
            return carteras;



        }
    }
}
