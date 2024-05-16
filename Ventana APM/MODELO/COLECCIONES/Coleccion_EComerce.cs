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
   public class Coleccion_EComerce
    {
        string fileName = string.Empty;
        public Coleccion_EComerce()
        {

        }

        public List<eComerce> GenerarListado(string aux)
        {
            fileName = aux;
            List<eComerce> ecomerce = new List<eComerce>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando...");
            int contador_de_filas = 0;
            int contador_pcs = 0;

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

                        foreach (Row row in rows)
                        {
                            contador_de_filas++;
                            eComerce comerce = new eComerce();
                            int contador = 0;
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                contador++;
                                if (contador_pcs >= 1)
                                {

                                   try
                                {
                                    if ((c.CellValue != null) && (contador == contador_pcs))
                                    {
                                        comerce.PCS = c.CellValue.Text;
                                    }
                                }
                                catch (Exception ex)
                                {
                                        ////comerce.PCS = c.CellValue.Text;
                                        ////Console.WriteLine(ex);
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
                                        string sPattern = "PCS";
                                        if (igual.Equals("PCS"))
                                        {
                                            contador_pcs = contador;
                                            contador++;

                                        }
                                        else if (System.Text.RegularExpressions.Regex.IsMatch(igual, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                                        {
                                            contador_pcs = contador;
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                        continue;
                                    }
                                }
                            }
                            if (comerce.PCS.Equals(""))
                            {

                            }
                            else
                            {
                                ecomerce.Add(comerce);
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


            return ecomerce;



        }
    }
}
