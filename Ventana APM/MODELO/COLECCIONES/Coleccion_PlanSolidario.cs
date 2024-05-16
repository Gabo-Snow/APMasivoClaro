using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ventana_APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace Ventana_APM.MODELO.COLECCIONES
{
    public class Coleccion_PlanSolidario
    {
        string fileName = string.Empty;
        public Coleccion_PlanSolidario()
        {

        }

        public List<PlanSolidario> GenerarListado(string aux)
        {
            
            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            fileName = aux;
            List<PlanSolidario> planSolidarios = new List<PlanSolidario>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando Plan Solidairo...");
            int contador_1 = 0;
            int contador_2 = 0;
            int contador_3 = 0;
            int contador_4 = 0;//love me love say 
            //.-. .-. ._. .-. ._. .-. ._. love me love say want u love me kiss me kiss me kiss rai nau
            int contador_5 = 0;
            int contador_6 = 0;
            if (aux.Equals(""))
            {
                PlanSolidario plan = new PlanSolidario();
                planSolidarios.Add(plan);
                return planSolidarios;
            }
            //fileName = fbd.SelectedPath; // primer archivo //falta un try catch por si llega vacio
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

                        //Console.WriteLine("Row count = {0}", rows.LongCount());
                        //Console.WriteLine("Cell count = {0}", cells.LongCount());
                        // Or... via each row yu feel mi fiil laik a shunaber de jeven wo u ou oooh 
                        //int contadorLista = 0;
                        //;-----------------;
                        foreach (Row row in rows)
                        {
                            PlanSolidario plan = new PlanSolidario();
                            int contador = 0;
                            var algo = row.Elements<Cell>();
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                contador++;
                                if (contador_1 >= 1 && contador_2 >= 1 && contador_3 >= 1 && contador_4 >= 1 && contador_5 >= 1 && contador_6 >= 1)
                                {
                                    try
                                    {
                                        if ((c.CellValue != null) && (contador == contador_1))//1 pc 2 cuenta 3 monto 4 campana 5 aplica cdp 6 ciclo 7 fecha
                                        {
                                            plan.PCS = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_2))
                                        {
                                            plan.CUENTA = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_6))
                                        {
                                            plan.CICLO = c.CellValue.Text;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_3))
                                        {
                                            int tiene_punto = c.CellValue.Text.IndexOf('.');
                                            if (tiene_punto > 1)
                                            {
                                                plan.MONTO = double.Parse(c.CellValue.Text, culture_punto);
                                            }
                                            else
                                            {
                                                plan.MONTO = double.Parse(c.CellValue.Text, culture_coma);
                                            }
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_5) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            plan.APLICA_CDP = str;
                                        }
                                        else if ((c.CellValue != null) && (contador == contador_4) && (c.DataType == CellValues.SharedString))
                                        {
                                            int ssid = int.Parse(c.CellValue.Text);
                                            string str = sst.ChildElements[ssid].InnerText;
                                            plan.CAMPANA = str;
                                        }
                                    }
                                    catch (Exception ex)
                                    {

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
                                        else if (igual.Equals("CICLO"))
                                        {
                                            contador_6 = contador;
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                        continue;
                                    }
                                }



                            }

                            planSolidarios.Add(plan);
                        }
                        //  Console.ReadLine();


                    }



                }
            }
            catch (Exception)
            {

                throw;
            }


            return planSolidarios;


        }
    }
}
