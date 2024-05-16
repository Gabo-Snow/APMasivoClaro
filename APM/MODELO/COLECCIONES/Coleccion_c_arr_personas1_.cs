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
   public class Coleccion_c_arr_personas1_
    {
        string fileName = string.Empty;
        public Coleccion_c_arr_personas1_()
        {

        }

        public List<Cargos_Arrendamiento> GenerarListado(string aux)
        {
            List<Cargos_Arrendamiento> coleccion_C_Ars = new List<Cargos_Arrendamiento>();
            Console.WriteLine("Cargando...");

            string path = @aux+ @"\ARCHIVOS INTERMEDIOS\C_ARR_PERSONA_1.xlsx";
            string txt = String.Empty;
            
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart bkPart = doc.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = bkPart.Workbook;
                DocumentFormat.OpenXml.Spreadsheet.Sheet s = workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(sht => sht.Name == "ARR_PERSONA1").FirstOrDefault();
                WorksheetPart wsPart = (WorksheetPart)bkPart.GetPartById(s.Id);
                DocumentFormat.OpenXml.Spreadsheet.SheetData sheetdata = wsPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();

                var cells = sheetdata.Descendants<Cell>();
                var rows = sheetdata.Descendants<Row>();
                int contadorLista = 0;
                foreach (Row row in rows)
                {
                    
                    Cargos_Arrendamiento c_Ar = new Cargos_Arrendamiento();
                    int contador = 0;
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        if (contadorLista >= 1)
                        {
                            contador++;
                            if ((c.CellValue != null) && (contador == 1))
                            {
                                c_Ar.CuentaFinanciera = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 2))
                            {
                                c_Ar.recveivercustomer = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 3))
                            {
                                c_Ar.tipocargo = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 4))
                            {
                                c_Ar.codigodecargo = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 5))
                            {
                                c_Ar.offwer = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 6))
                            {
                                c_Ar.nombreOFFER = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 7))
                            {
                                c_Ar.promo = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 8))
                            {
                                c_Ar.description = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 9))
                            {
                                c_Ar.montoCargos = double.Parse(c.CellValue.Text);
                            }
                            else if ((c.CellValue != null) && (contador == 10))
                            {
                                c_Ar.montoDescuentos = double.Parse(c.CellValue.Text);
                            }
                            else if ((c.CellValue != null) && (contador == 11))
                            {
                                c_Ar.montoTotal = double.Parse(c.CellValue.Text);
                            }
                            else if ((c.CellValue != null) && (contador == 12))
                            {
                                c_Ar.tipodeCobro = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 13))
                            {
                                c_Ar.PCS = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 14))
                            {
                                c_Ar.secuenciadeCiclo = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 15))
                            {
                                c_Ar.segmento = c.CellValue.Text;
                            }
                        }

                    }
                    coleccion_C_Ars.Add(c_Ar);
                    contadorLista++;
                }
                
            }
            Console.WriteLine("Cargando...");
            return coleccion_C_Ars;
           

        }
    }
}