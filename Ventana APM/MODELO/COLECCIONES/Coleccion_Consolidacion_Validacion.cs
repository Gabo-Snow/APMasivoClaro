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
    public class Coleccion_Consolidacion_Validacion
    {
        string fileName = string.Empty;
        public Coleccion_Consolidacion_Validacion()
        {

        }

        public List<Consolidacion_Validacion> GenerarListado(string aux)//Cambiar todo lo que dice plansolidario por consolidado validacion
        {
            CultureInfo culture_punto = new CultureInfo("en-US");
            CultureInfo culture_coma = new CultureInfo("es-ES");
            fileName = aux;
            List<Consolidacion_Validacion> consolidacion_Validacions = new List<Consolidacion_Validacion>();
            Console.WriteLine(fileName);
            Console.WriteLine("Cargando Consolidacion...");
            int contador_1 = 0;
            int contador_2 = 0;
            int contador_3 = 0;
            int contador_4 = 0;
            int contador_5 = 0;
            int contador_6 = 0;
            bool hay_promo = false;
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

                    foreach (Row row in rows)
                    {
                        Consolidacion_Validacion validacion = new Consolidacion_Validacion();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;
                            if (contador_1 >= 1 && contador_2 >= 1 && contador_3 >= 1)
                            {// revisar y arreglar los excel cuando un campo viene nulo antes de la celda que se debe almacenar
                                try
                                {
                                    if ((c.CellValue != null) && (contador == contador_4))
                                    {
                                        int ssidd = int.Parse(c.CellValue.Text);
                                        string strr = sst.ChildElements[ssidd].InnerText;
                                        string igual = strr.ToUpper();
                                        validacion.codigo = igual;
                                        

                                    }
                                    if ((c.CellValue != null) && (contador == contador_1))//1 pc 2 cuenta 3 monto 4 campana 5 aplica cdp 6 ciclo 7 fecha
                                    {
                                        
                                        //int ssidd = int.Parse(c.CellValue.Text);
                                        //string strr = sst.ChildElements[ssidd].InnerText;
                                        //string igual = strr.ToUpper();

                                        
                                        validacion.pcs = c.CellValue.Text;
                                    }
                                    else if ((c.CellValue != null) && (contador == contador_2))
                                    {
                                        //int ssidd = int.Parse(c.CellValue.Text);
                                        //string strr = sst.ChildElements[ssidd].InnerText;
                                        //string igual = strr.ToUpper();

                                        validacion.monto_validador = c.CellValue.Text;//esto tendria que ir en int
                                    }
                                    else if ((c.CellValue != null) && (contador == contador_3))
                                    {
                                        validacion.monto_ajuste = c.CellValue.Text;//esto tendria que ir en int
                                    }
                                }
                                catch (Exception ex)
                                {

                                    continue;
                                }
                            }
                            else//
                            {
                                try//
                                {
                                    int ssidd = int.Parse(c.CellValue.Text);
                                    string strr = sst.ChildElements[ssidd].InnerText;
                                    string igual = strr.ToUpper();
                                    if (igual.Equals("PCS"))
                                    {
                                        contador_1 = contador;
                                        
                                        if (hay_promo)
                                        {
                                            contador_1 = contador_1 - 1;
                                        }

                                    }
                                    else if (igual.Equals("MONTO"))
                                    {
                                        contador_2 = contador;
                                        if (hay_promo)
                                        {
                                            contador_2 = contador_2 - 1;
                                        }
                                        
                                    }
                                    else if (igual.Equals("AJUSTE"))
                                    {
                                        contador_3 = contador;
                                        if (hay_promo)
                                        {
                                            contador_3 = contador_3 - 1;
                                        }
                                        
                                    }
                                    else if (igual.Equals("CODIGODECARGO"))
                                    {
                                        contador_4 = contador;
                                        if (hay_promo)
                                        {
                                            contador_4 = contador_4 - 1;
                                        }
                                    }
                                    else if (igual.Equals("CODIGO DE CARGO"))
                                    {
                                        contador_4 = contador;
                                        if (hay_promo)
                                        {
                                            contador_4 = contador_4 - 1;
                                        }
                                    }
                                    else if (igual.Equals("PROMO"))
                                    {
                                        hay_promo = true;
                                        //poner un mensaje de alerta, si no cruza nada favor de borrar la columna PROMO
                                    }
                                }
                                catch (Exception ex)
                                {

                                    continue;
                                }
                            }

                            //.

                        }

                        if (validacion.codigo.Equals("CCOCSUBEQPCOMM"))
                        {
                            consolidacion_Validacions.Add(validacion);
                        }
                        
                    }
                    //  Console.ReadLine();
                }



            }
            return consolidacion_Validacions;



        }
    }
}
