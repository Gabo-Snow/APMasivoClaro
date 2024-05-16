using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;

namespace Ventana_APM.MODELO.COLECCIONES
{
   public class Coleccion_Cuentas_Asignadas
    {
        string fileName = string.Empty;
        public Coleccion_Cuentas_Asignadas()
        {

        }

        public List<Cuentas_Asignadas> GenerarListado(string aux)
        {
            fileName = aux;
            List<Cuentas_Asignadas> cuentas_Asignadas = new List<Cuentas_Asignadas>();
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
                    int contador_de_columnas = 0;
                    foreach (Row row in rows)
                    {
                        Cuentas_Asignadas cuentas = new Cuentas_Asignadas();
                        contador_de_columnas++;
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;
                            if (contador_de_columnas == 1)
                            {
                                break;
                            }
                            else
                            {
                                try
                                {
                                    if ((c.CellValue != null) && (contador == 1))
                                    {
                                        cuentas.PCS = c.CellValue.Text;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if ((c.CellValue != null) && (contador == 1))
                                    {
                                      //  cuentas.PCS = c.CellValue.Text;
                                    }

                                }

                            }


                        }
                        cuentas_Asignadas.Add(cuentas);

                    }
                    //  Console.ReadLine();
                }



            }

            return cuentas_Asignadas;



        }
    }
}
