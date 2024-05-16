﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;

namespace Ventana_APM.Auxiliares
{
   public class Aux_Activacion
    {
        public List<Activacion> GenerarListado_fvm(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = @aux + @"\Programa\Auxiliares_Activacion\Auxiliar_FVM_Activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_habilitacion(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = @"Sdk\bin\Auxiliares_Activacion\Auxiliar_Habilitacion_activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {


                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_portados(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = @aux + @"\Programa\Auxiliares_Activacion\Auxiliar_Portados.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_solidarios(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = @aux + @"\Programa\Auxiliares_Activacion\Auxiliar_Solidario_Activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                ////Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_original(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = @aux+@"\Programa\Auxiliares_activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {


                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_neteados(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = aux+ @"\Programa\Auxiliares_Activacion\Neteado_Auxiliar_Activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                ////Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_asignados(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = aux + @"\Programa\Auxiliares_Activacion\Asignar_Activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.TIPO = col[14];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
        public List<Activacion> GenerarListado_EMPRESAS(string aux)
        {
            try
            {
                CultureInfo culture_punto = new CultureInfo("en-US");
                CultureInfo culture_coma = new CultureInfo("es-ES");
                string fileName = string.Empty;
                List<Activacion> coleccion_C_Ars = new List<Activacion>();
                fileName = @"Sdk\bin\Auxiliares_Activacion\Asignar_Activacion.txt";

                System.IO.StreamReader file = new System.IO.StreamReader(@fileName);
                BufferedStream bs = null;
                StreamReader sr = null;
                int counter = 0;

                using (FileStream fs = File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //
                using (bs = new BufferedStream(fs))
                using (sr = new StreamReader(bs, Encoding.Default))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        counter++;
                        Activacion c_Ar = new Activacion();
                        string[] col = line.Split(new char[] { '|' });
                        if (counter > 1)
                        {
                            int tiene_punto_monto_cargos = col[8].IndexOf('.');
                            int tiene_punto_montoDescuentos = col[9].IndexOf('.');
                            int tiene_punto_montoTotal = col[10].IndexOf('.');
                            c_Ar.CuentaFinanciera = col[0];
                            c_Ar.recveivercustomer = col[1];
                            c_Ar.tipocargo = col[2];
                            c_Ar.codigodecargo = col[3];
                            c_Ar.offwer = col[4];
                            c_Ar.nombreOFFER = col[5];
                            c_Ar.promo = col[6];
                            c_Ar.description = col[7];
                            //todos los computadores tienen distintas configuraciones regionales-----------------------------------------------------
                            if (tiene_punto_monto_cargos > 1)
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoCargos = double.Parse(col[8], culture_coma);
                            }
                            if (tiene_punto_montoDescuentos > 1)
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoDescuentos = double.Parse(col[9], culture_coma);
                            }
                            if (tiene_punto_montoTotal > 1)
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_punto);
                            }
                            else
                            {
                                c_Ar.montoTotal = double.Parse(col[10], culture_coma);
                            }
                            //---------------------------------------------------------------------------------------------------------------------
                            c_Ar.tipodeCobro = col[11];
                            c_Ar.PCS = col[12];
                            c_Ar.secuenciadeCiclo = col[13];
                            c_Ar.prorrateo = col[14];
                            c_Ar.rut = col[15];
                            coleccion_C_Ars.Add(c_Ar);
                            //Console.WriteLine("Cargos almacenados en memoria : {0}", counter);
                        }
                    }
                    return coleccion_C_Ars;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ingresar archivos correctos");
                //MessageBox.Show("El error es {0}", ex.ToString());
                //Application.Exit();
                return null;
            }



        }
    }
}
