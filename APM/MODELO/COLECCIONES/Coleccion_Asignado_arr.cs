using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using APM.MODELO.CLASES;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APM.MODELO.COLECCIONES
{
    public class Coleccion_Asignado_arr
    {
        string fileName = string.Empty;
        public Coleccion_Asignado_arr()
        {

        }

        public List<Asignados_arr> GenerarListado(string aux)
        {
            List<Asignados_arr> asignados_Arrs = new List<Asignados_arr>();

            Console.WriteLine("Cargando...");
            aux = @"C:\Users\Marcelo\Desktop";//jim deberia renunciar o no rendirse
            //fileName = fbd.SelectedPath; // primer archivo
            fileName = aux + @"\ENTRADA\Asignacion Procesada.xlsx";
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
                        Asignados_arr asignados = new Asignados_arr();
                        int contador = 0;
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            contador++;
                            if ((c.CellValue != null) && (contador == 1) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.ASIGNACIÓN = str;
                            }
                            else if ((c.CellValue != null) && (contador == 2))
                            {
                                asignados.Cuenta_Financiera = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 3))
                            {
                                asignados.RECEIVER_CUSTOMER = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 4) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.TIPO_CARGO = str;
                            }
                            else if ((c.CellValue != null) && (contador == 5) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.Codigo_de_Cargo = str;
                            }
                            else if ((c.CellValue != null) && (contador == 6))
                            {
                                asignados.OFFER = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 7) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.Nombre_OFFER = str;
                            }
                            else if ((c.CellValue != null) && (contador == 8) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.DESCRIPTION = str;
                            }
                            else if ((c.CellValue != null) && (contador == 8))
                            {
                                asignados.PROMO = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 9))
                            {
                                asignados.MONTO_CARGOS = double.Parse(c.CellValue.Text);
                            }
                            else if ((c.CellValue != null) && (contador == 10))
                            {
                                asignados.MONTO_DESCUENTOS = double.Parse(c.CellValue.Text);
                            }
                            else if ((c.CellValue != null) && (contador == 11))
                            {
                                asignados.MONTO_TOTAL = double.Parse(c.CellValue.Text);
                            }
                            else if ((c.CellValue != null) && (contador == 12) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.Tipo_de_Cobro = str;
                            }
                            else if ((c.CellValue != null) && (contador == 13))
                            {
                                asignados.PCS = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 14))
                            {
                                asignados.Secuencia_de_Ciclo = c.CellValue.Text;
                            }
                            else if ((c.CellValue != null) && (contador == 15) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.COBRAR = str;
                            }
                            else if ((c.CellValue != null) && (contador == 16))
                            {
                                asignados.MONTO = Convert.ToDouble((c.CellValue.Text).ToString(CultureInfo.CreateSpecificCulture("en-US")));
                            }
                            else if ((c.CellValue != null) && (contador == 17) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.CUOTA = str;
                            }
                            else if ((c.CellValue != null) && (contador == 18) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                asignados.OBSERVACIÓN = str;
                            }

                        }

                        asignados_Arrs.Add(asignados);
                        Console.WriteLine("Almacenando ajustes {0}", asignados_Arrs.Count());
                    }
                    //  Console.ReadLine();
                }
            }


            //C_ARR_PERSONA_AJESP_ASIG_FINAL(asignados_Arrs);
            return asignados_Arrs;

        }

        //private bool C_ARR_PERSONA_AJESP_ASIG_FINAL(List<Asignados_arr> _Arrs)
        //{
        //    try
        //    {
        //        var miTable = new DataTable("OBSEVACION");
        //        miTable.Columns.Add("CUENTA FINANCIERA");
        //        miTable.Columns.Add("CODIGO DE CARGOS");
        //        miTable.Columns.Add("DESCRIPCION");
        //        miTable.Columns.Add("PCS");
        //        miTable.Columns.Add("TIPO");
        //        miTable.Columns.Add("AJUSTE");
        //        miTable.Columns.Add("OFFER");
        //        miTable.Columns.Add("OBSERVACION");




        //        foreach (var item in _Arrs)
        //        {
        //            double ajuste = item.MONTO - item.MONTO_TOTAL;
        //            if (item.COBRAR.ToUpper().Equals("NO")&& item.MONTO == 0 && ajuste <= 0 )
        //            {
        //                miTable.Rows.Add(new Object[] { item.Cuenta_Financiera, item.Codigo_de_Cargo, item.DESCRIPTION,item.PCS, item.TIPO_CARGO,
        //             item.MONTO_TOTAL - item.MONTO,item.OFFER, ""});
        //            }

        //        }

        //        SaveExcel.BuildExcel(miTable, @"C:\Users\Marcelo\Desktop" + @"\ARCHIVOS SALIDA" + @"\C_ARR_PERSONA_AJESP_FINAL.xlsx");

        //    }
        //    catch (Exception ex)
        //    {

        //        throw;
        //    }
        //    return true;
        //}

    }
}
