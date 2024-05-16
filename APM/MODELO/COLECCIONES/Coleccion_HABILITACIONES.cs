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
    public class Coleccion_HABILITACIONES
    {
        string fileName = string.Empty;
        public Coleccion_HABILITACIONES()
        {

        }
        public List<HABILITACIONES> GenerarListado(string aux)
        {
            fileName = aux + @"\ENTRADA\HABILITACIONES.xlsx";
            List<HABILITACIONES> hABILITACIONEs = new List<HABILITACIONES>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando HABILITACIONES...");
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
                    
                   // int contadorLista = 0;
                    foreach (Row row in rows)
                    {
                        HABILITACIONES hABILITACIONE = new HABILITACIONES();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;

                            if ((c.CellValue != null) && (contador == 1) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                hABILITACIONE.CCOCSUBEQPCOM = str;
                            }
                            else if ((c.CellValue != null) && (contador == 2))
                            {
                                hABILITACIONE.MONTO_SIN_IMPUESTO = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 3))
                            {
                                hABILITACIONE.PCS = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 4) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                hABILITACIONE.INGRESO_VTA = str;
                            }
                            else if ((c.CellValue != null) && (contador == 5))
                            {
                                hABILITACIONE.FA = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 6))
                            {
                                hABILITACIONE.CBP = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 7))
                            {
                                hABILITACIONE.FECHA_ACTIVACION = c.CellValue.Text;
                            }
                        }

                        hABILITACIONEs.Add(hABILITACIONE);
                    }
                    //  Console.ReadLine();
                }



            }
            return hABILITACIONEs;



        }
    }
}
