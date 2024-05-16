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
   public class Coleccion_CFM
    {
        string fileName = string.Empty;
        public Coleccion_CFM()
        {

        }

        public List<CFM_AJUSTES> GenerarListado(string aux)
        {
            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            List<CFM_AJUSTES> cFM_AJUSTEs = new List<CFM_AJUSTES>();
            int contador_1 = 0;
            int contador_2 = 0;
            int contador_3 = 0;
            int contador_4 = 0;
            int contador_5 = 0;
            Console.WriteLine("Cargando...");
            //fileName = fbd.SelectedPath; // primer archivo
            fileName = aux;
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
                    int quitar_encabezado = 0;
                    foreach (Row row in rows)
                    {
                        CFM_AJUSTES cFM_ = new CFM_AJUSTES();
                        quitar_encabezado++;
                                int contador = 0;
                                foreach (Cell c in row.Elements<Cell>())
                                {
                                contador++;
                                if (contador_1 >= 1 && contador_2 >= 1 && contador_3 >= 1 && contador_4 >= 1 && contador_5 >= 1)
                                {
                                    try
                                    {
                                        if ((c.CellValue != null) && (contador == contador_1))
                                        {
                                            cFM_.PCS = c.CellValue.Text;
                                            // cFM_.PCS = str;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_2))
                                        {
                                            cFM_.CUENTA = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_4) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            cFM_.CAMPANA = str;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_3))
                                        {
                                            int tiene_punto = c.CellValue.Text.IndexOf('.');
                                            if (tiene_punto > 1)
                                            {
                                                cFM_.MONTO = double.Parse(c.CellValue.Text, culture_punto);
                                            }
                                            else
                                            {
                                                cFM_.MONTO = double.Parse(c.CellValue.Text, culture_coma);
                                            }
                                            //cFM_.MONTO = Convert.ToDouble(c.CellValue.Text);
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_5) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            cFM_.APLICA_CDP = str;
                                        }
                                        //else if ((c.CellValue != null) && (contador == 6) && (c.DataType == CellValues.SharedString))
                                        //{
                                        //    int ssid = int.Parse(c.CellValue.Text);
                                        //    string str = sst.ChildElements[ssid].InnerText;
                                        //    cFM_.CAMPANA = str;
                                        //}
                                    }

                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("El error es : {0}", ex);
                                        if ((c.CellValue != null) && (contador == 5))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            cFM_.APLICA_CDP = str;
                                        }
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        int ssidd = int.Parse(c.CellValue.Text);
                                        string strr = sst.ChildElements[ssidd].InnerText;
                                        string igual = strr.ToUpper();        

                                        if (igual.Equals("PCS"))
                                        {
                                            contador_1 = contador;
                                        }
                                        else if (igual.Equals("CUENTA"))
                                        {
                                            contador_2 = contador;
                                        }
                                        else if (igual.Equals("MONTO"))
                                        {
                                            contador_3 = contador;
                                        }
                                        else if (igual.Equals("CAMPANA"))
                                        {
                                            contador_4 = contador;
                                        }
                                        else if (igual.Equals("APLICA_CDP"))
                                        {
                                            contador_5 = contador;
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                        continue;
                                    }
                                }
                            
                        }
                        if (cFM_.PCS.Equals(""))
                        {

                        }
                        else
                        {
                            cFM_AJUSTEs.Add(cFM_);
                        }
                        
                    }
                    //  Console.ReadLine();
                }
            }



            return cFM_AJUSTEs;

        }
    }
}
