using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ventana_APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace Ventana_APM.MODELO.COLECCIONES
{

   public class Coleccion_ajustes
    {

        string fileName = string.Empty;
        CultureInfo culture_punto = new CultureInfo("en-US");
        CultureInfo culture_coma = new CultureInfo("es-ES");
        public Coleccion_ajustes()
        {

        }
        //y que la gente se enamore d emi voz, mejor que labures, mejor que trabajes, ya me canse de ser tu fuente de dinero
        //voy a ponerte esa guitarra de sombrero
        //cha cha cha
        //aqui empiezan los problemas

        public List<Ajustes> GenerarListado(string aux) //hacer un trycatch aqui
        {

            List<Ajustes> ajustes = new List<Ajustes>();
            //-----
            Console.WriteLine("Cargando...");//
            if (aux.Equals(""))
            {
                Ajustes ajuste = new Ajustes();
                ajustes.Add(ajuste);
                return ajustes;
            }


            //
            int contador_1 = 0;
            int contador_2 = 0;
            int contador_3 = 0; 
            int contador_4 = 0;
            int contador_5 = 0;
            int contador_lineas = 0;
            //fileName = fbd.SelectedPath; //
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
                        
                        //int contadorLista = 0;
                        foreach (Row row in rows)
                        {
                            Ajustes ajuste = new Ajustes();
                            int contador = 0;
                        if (contador_lineas == 0)
                        {
                            int contador_encabezado = 0;
                            foreach (var item in row.Elements<Cell>())
                            {
                                contador_encabezado++;
                                int ssidd = int.Parse(item.CellValue.Text);
                                string strr = sst.ChildElements[ssidd].InnerText;
                                string igual = strr.ToUpper();
                                string sPattern = "CARGO";
                                string sPattern2 = "FINANCIERA";

                                if (igual.Equals("PCS"))
                                {
                                    contador_1 = contador_encabezado;
                                }
                                else if (igual.Equals("CUENTA"))
                                {
                                    contador_2 = contador_encabezado;
                                }
                                else if (igual.Equals("CODIGO DE CARGO"))
                                {
                                    contador_3 = contador_encabezado;
                                }
                                else if (igual.Equals("AMOUNT"))
                                {
                                    contador_4 = contador_encabezado;
                                }
                                else if (igual.Equals("CICLO"))
                                {
                                    contador_5 = contador_encabezado;
                                }
                                if (System.Text.RegularExpressions.Regex.IsMatch(igual, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                                {
                                    contador_3 = contador_encabezado;
                                }
                                if (System.Text.RegularExpressions.Regex.IsMatch(igual, sPattern2, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                                {
                                    contador_2 = contador_encabezado;
                                }
                            }
                            contador_lineas++;
                            if (contador_1 == 0 || contador_2 == 0 || contador_3 == 0 || contador_4 == 0 || contador_5 == 0)
                            {
                                MessageBox.Show("ERROR EN EL ENCABEZADO DEL EXCEL AJUSTES");
                                MessageBox.Show("EL ENCABEZADO DEBE SER : PCS | CUENTA FINANCIERA | CODIGO DE CARGO | AMOUNT | CICLO");
                                break;
                                
                            }
                            else
                            {
                                continue;
                            }
                            
                        }


                        foreach (Cell c in row.Elements<Cell>())
                        {
                                contador++;
                                try
                                {

                                    if ((c.CellValue != null) && (contador == contador_1))
                                    {
                                        ajuste.PCS = c.CellValue.Text;
                                    }
                                    else if ((c.CellValue != null) && (contador == contador_2))
                                    {
                                        ajuste.CuentaFinanciera = c.CellValue.Text;
                                    }
                                    else if ((c.CellValue != null) && (contador == contador_3) && (c.DataType == CellValues.SharedString))
                                    {
                                        int ssid = int.Parse(c.CellValue.Text);
                                        string str = sst.ChildElements[ssid].InnerText;
                                        ajuste.codigodecargo = str;// str...
                                    }
                                    else if ((c.CellValue != null) && (contador == contador_4))
                                    {
                                        int tiene_punto = c.CellValue.Text.IndexOf('.');
                                        if (tiene_punto > 1)
                                        {
                                            ajuste.amount = double.Parse(c.CellValue.Text, culture_punto);
                                        }
                                        else
                                        {
                                            ajuste.amount = double.Parse(c.CellValue.Text, culture_coma);
                                        }
                                    }
                                    else if ((c.CellValue != null) && (contador == contador_5) && (c.DataType == CellValues.SharedString))
                                    {
                                        int ssid = int.Parse(c.CellValue.Text);
                                        string str = sst.ChildElements[ssid].InnerText;
                                        ajuste.ciclo = str;
                                    }
                                }
                                catch (Exception ex)
                                {

                                    continue;
                                }
                                //try
                                //{
                                //    int ssidd = int.Parse(c.CellValue.Text);
                                //    string strr = sst.ChildElements[ssidd].InnerText;
                                //    string igual = strr.ToUpper();
                                //    string sPattern = "CARGO";

                                //    if (igual.Equals("PCS"))
                                //    {
                                //        contador_1 = contador;
                                //    }
                                //    else if (igual.Equals("CUENTA"))
                                //    {
                                //        contador_2 = contador;
                                //    }
                                //    else if (igual.Equals("CODIGO DE CARGO"))
                                //    {
                                //        contador_3 = contador;
                                //    }
                                //    else if (igual.Equals("AMOUNT"))
                                //    {
                                //        contador_4 = contador;
                                //    }
                                //    else if (igual.Equals("CICLO"))
                                //    {
                                //        contador_5 = contador;
                                //    }
                                //    if (System.Text.RegularExpressions.Regex.IsMatch(igual, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                                //    {
                                //        contador_3 = contador;
                                //    }
                                //}
                                //catch (Exception ex)
                                //{

                                //    continue;
                                //}
                           //haciendo como que trabaj
                        }

                        if (ajuste.PCS.Equals(""))
                        {

                        }
                        else
                        {
                            ajustes.Add(ajuste);
                        }
                            
                        }
                        //  Console.ReadLine();
                    }
                }


            
            return ajustes;

        }
    }
}
