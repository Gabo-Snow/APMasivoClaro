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
    public class Coleccion_PlanSolidario
    {
        string fileName = string.Empty;
        public Coleccion_PlanSolidario()
        {

        }

        public List<PlanSolidario> GenerarListado(string aux)
        {
            fileName = aux + @"\ENTRADA\PLAN SOLIDARIO.xlsx";
            List<PlanSolidario> planSolidarios = new List<PlanSolidario>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando Plan Solidairo...");
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
                        PlanSolidario plan = new PlanSolidario();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;

                            if ((c.CellValue != null) && (contador == 2))
                            {
                                plan.PCS = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 3))
                            {
                                plan.CUENTA = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 4))
                            {
                                plan.CICLO = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 5))
                            {
                                plan.MONTO = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 6) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                plan.APLICA_CDP = str;
                            }
                            else if ((c.CellValue != null) && (contador == 7) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                plan.CAMPANA = str;
                            }
                        }

                        planSolidarios.Add(plan);
                    }
                    //  Console.ReadLine();
                }



            }
            return planSolidarios;



        }
    }
}
