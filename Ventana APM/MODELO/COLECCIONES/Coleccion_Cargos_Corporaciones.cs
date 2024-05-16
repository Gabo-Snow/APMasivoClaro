using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ventana_APM.MODELO.CLASES;

namespace Ventana_APM.MODELO.COLECCIONES
{
    public class Coleccion_Cargos_Corporaciones
    {

        string fileName = string.Empty;
        public Coleccion_Cargos_Corporaciones()
        {

        }
        public List<Cargos_Cuenta_Corporaciones> GenerarListado(string aux)
        {

            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            fileName = aux;
            List<Cargos_Cuenta_Corporaciones> hABILITACIONEs = new List<Cargos_Cuenta_Corporaciones>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando HABILITACIONES...");
            int contador_1 = 0;
            int contador_2 = 0;
            int contador_3 = 0;
            int contador_4 = 0;
            int contador_5 = 0;
            int contador_6 = 0;
            int contador_7 = 0;
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
                        Cargos_Cuenta_Corporaciones hABILITACIONE = new Cargos_Cuenta_Corporaciones();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;

                            try
                            {
                                    if ((c.CellValue != null) && (contador == 8) && (c.DataType == CellValues.SharedString))
                                    {
                                        int ssid = int.Parse(c.CellValue.Text);
                                        string str = sst.ChildElements[ssid].InnerText;
                                        hABILITACIONE.DESCRIPTION = str;
                                    }
                                
                                //else
                                //{
                                //    try
                                //    {
                                //        int ssidd = int.Parse(c.CellValue.Text);
                                //        string strr = sst.ChildElements[ssidd].InnerText;
                                //        string igual = strr.ToUpper();


                                //        if (igual.Equals("CCOCSUBEQPCOM"))
                                //        {
                                //            contador_1 = contador;
                                //        }
                                //        else if (igual.Equals("MONTO_SIN_IMPUESTO"))
                                //        {
                                //            contador_2 = contador;
                                //        }
                                //        else if (igual.Equals("PCS"))
                                //        {
                                //            contador_3 = contador;
                                //        }
                                //        else if (igual.Equals("INGRESO_VTA"))
                                //        {
                                //            contador_4 = contador;
                                //        }
                                //        else if (igual.Equals("FA"))
                                //        {
                                //            contador_5 = contador;
                                //        }
                                //        else if (igual.Equals("CBP"))
                                //        {
                                //            contador_6 = contador;
                                //        }
                                //        else if (igual.Equals("FECHA_ACTIVACION"))
                                //        {
                                //            contador_7 = contador;
                                //            contador++;
                                //        }
                                //    }
                                //    catch (Exception ex)
                                //    {

                                //        continue;
                                //    }
                                //}

                            }
                            catch (Exception EX)
                            {

                                continue;
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
