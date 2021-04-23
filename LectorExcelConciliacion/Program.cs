﻿using System;
using System.IO;
using System.Configuration;
using System.Data;
using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LectorExcelConciliacion
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static void Main(string[] args)
        {
            const int SW_HIDE = 0;
            //const int SW_SHOW = 5;
            var handle = GetConsoleWindow();

            Process currentProcess = Process.GetCurrentProcess();
            Console.Title = "Lector Excel Conciliación (PID: " + currentProcess.Id + ")";
            Console.WriteLine("Lector Excel Conciliación (PID: " + currentProcess.Id + ")");

            //String de conexion
            ConnString connString = new ConnString();
            //customString
            string conexion;
            if (ConfigurationManager.AppSettings["customConnection"].ToString() != "")
            {
                conexion = ConfigurationManager.AppSettings["customConnection"].ToString();
            }
            else
            {
                conexion = connString.GetString(ConfigurationManager.AppSettings["ambiente"].ToString());
            }
            //

            //ruta INPUT
            string rutainput = ObtRuta(conexion, 14, "INPUT");
            //

            //ruta OUTPUT
            string rutaoutput = ObtRuta(conexion, 15, "OUTPUT");
            //

            //ruta WORK
            string rutawork = ObtRuta(conexion, 29, "WORK");
            //Sub carpeta dentro de WORK
            rutawork = rutawork + currentProcess.Id + Path.DirectorySeparatorChar;
            if (!Directory.Exists(rutawork))
            {
                try
                {
                    Directory.CreateDirectory(rutawork);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error creando carpeta WORK" + Path.DirectorySeparatorChar + currentProcess.Id + Path.DirectorySeparatorChar);
                    Console.WriteLine(ex.Message);
                    Environment.Exit(0);
                }
            }
            else
            {
                BorrarSubWork(rutawork, currentProcess.Id, false);
                Directory.CreateDirectory(rutawork);
            }
            //

            int offset = 0;
            bool pause = true;
            bool fileindividual = false;
            while (args.Length - offset > 0)
            {
                if (args[0 + offset] == "-help" || args[0 + offset] == "--help" || args[0 + offset] == "-nopause" || args[0 + offset] == "-hide" || args[0 + offset] == "-killall" || args[0 + offset] == "-file" || args[0 + offset] == "-path")
                {
                    if (args[0 + offset] == "-help" || args[0 + offset] == "--help")
                    {
                        BorrarSubWork(rutawork, currentProcess.Id, true);
                        ArgInvalid(false, false);
                    }
                    if (args[0 + offset] == "-nopause")
                    {
                        pause = false;
                        offset++;
                    }
                    else if (args[0 + offset] == "-hide")
                    {
                        ShowWindow(handle, SW_HIDE);
                        //ShowWindow(handle, SW_SHOW);
                        offset++;
                    }
                    else if (args[0 + offset] == "-killall")
                    {
                        foreach (var process in Process.GetProcessesByName("LectorExcelConciliacion"))
                        {
                            if (process.Id != currentProcess.Id)
                            {
                                process.Kill();
                                Console.WriteLine(" LectorExcelConciliacion con PID " + process.Id + " cerrado.");
                            }
                        }
                        BorrarSubWork(rutawork, currentProcess.Id, true);
                        Environment.Exit(0);
                    }
                    else if (args[0 + offset] == "-file")
                    {
                        fileindividual = true;
                        if (args[1 + offset] == null)
                        {
                            BorrarSubWork(rutawork, currentProcess.Id, true);
                            ArgInvalid(false, true);
                        }
                        if (File.Exists(args[1 + offset]))
                        {
                            Console.Write(" Moviendo archivo a ruta Work... ");
                            try
                            {
                                string nombrearchivo = Path.Combine(rutawork, Path.GetFileNameWithoutExtension(args[1 + offset]) + "_" + currentProcess.Id + Path.GetExtension(args[1 + offset]));
                                File.Move(args[1 + offset], nombrearchivo);
                                Console.WriteLine(" OK");
                                Console.WriteLine("    Procesando >> " + nombrearchivo);
                                ExecuteExcel(Path.GetFileName(nombrearchivo), rutawork, rutaoutput, conexion);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                        else
                        {
                            Console.WriteLine(" Archivo no encontrado\n Ruta: " + args[1 + offset]);
                        }
                        offset += 2;
                    }
                    else if (args[0 + offset] == "-path")
                    {
                        rutainput = args[1 + offset];
                        offset += 2;
                    }
                }
                else
                {
                    BorrarSubWork(rutawork, currentProcess.Id, true);
                    if (offset == 0)
                        ArgInvalid(true, true);
                    else
                        ArgInvalid(pause, true);
                }
            }

            //Caso no halla parametro -file, revisa carpeta
            if (!fileindividual)
            {
                Console.Write(" Buscardo archivos en la carpeta...");
                string[] dirs = Directory.GetFiles(rutainput);


                if (dirs.Length > 0)
                {
                    Console.WriteLine(" OK");
                    Console.WriteLine("  Se encontró " + dirs.Length + " archivos:");

                    foreach (string dir in dirs)
                    {
                        Console.WriteLine("   > " + dir);
                    }
                }
                else
                {
                    Console.WriteLine("  No se encontró archivos en la ruta\n  Ruta: " + rutainput);
                    if (pause)
                    {
                        Console.WriteLine("Presione cualquier tecla para salir...");
                        Console.ReadKey();
                    }

                    BorrarSubWork(rutawork, currentProcess.Id, true);
                    Environment.Exit(0);
                }
                Console.WriteLine(" Moviendo archivos a ruta Work...");
                foreach (string dir in dirs)
                {
                    try
                    {
                        string nombrearchivo = Path.Combine(rutawork, Path.GetFileNameWithoutExtension(dir) + "_" + currentProcess.Id + Path.GetExtension(dir));
                        File.Move(dir, nombrearchivo);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                string[] dirswork = Directory.GetFiles(rutawork);
                int count = 1;
                foreach (string dir in dirswork)
                {
                    Console.WriteLine(" Procesando " + count + "/" + dirswork.Length + " >> " + Path.GetFileName(dir));
                    ExecuteExcel(Path.GetFileName(dir), rutawork, rutaoutput, conexion);
                    count++;
                }
            }

            BorrarSubWork(rutawork, currentProcess.Id, true);
            Console.WriteLine("Fin");
            if (pause)
            {
                Console.WriteLine("Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
        }

        static void ExecuteExcel(string filename, string rutawork, string rutaoutput, string conexion)
        {
            string varchivovalido = SelectFromWhere(conexion, "SELECT SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1) " +
                                                    "FROM concargarchivos " +
                                                    "WHERE SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1) IS NOT NULL " +
                                                    "AND UPPER('" + filename + "') " +
                                                    "LIKE '%'||SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1)||'%' " +
                                                    "GROUP BY nombrearchivocarga " +
                                                    "ORDER BY nombrearchivocarga", true);
            if (!(String.IsNullOrEmpty(varchivovalido)))
            {
                Read_file(Path.Combine(rutawork, filename), varchivovalido, conexion);
                Readed_file(Path.Combine(rutawork, filename), rutaoutput, conexion);
            }
            else
            {
                Console.WriteLine(" >>> Rechazado: Nombre Archivo no se encuentra parametrizado en la tabla CONCARGARCHIVOS");
            }
        }
        public static void Read_file(string xlsFilePath, string ArchivoValido, string conexion)
        {
            if (!File.Exists(xlsFilePath))
                return;

            Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Lectura y Escritura del excel... ");
            FileInfo fi = new FileInfo(xlsFilePath);
            long filesize = fi.Length;

            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Range range;
            var misValue = Type.Missing;

            // abrir el documento 
            xlApp = new Application();
            xlWorkBook = xlApp.Workbooks.Open(xlsFilePath, misValue, misValue,
                misValue, misValue, misValue, misValue, misValue, misValue,
                misValue, misValue, misValue, misValue, misValue, misValue);

            // seleccion de la hoja de calculo
            // get_item() devuelve object y numera las hojas a partir de 1
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // seleccion rango activo
            range = xlWorkSheet.UsedRange;
            range.Columns.AutoFit();

            // Variables de Campos
            string vCampo_A, vCampo_B, vCampo_C, vCampo_D, vCampo_E, vCampo_F, vCampo_G, vCampo_H, vCampo_I, vCampo_J, vCampo_K, vCampo_L, vCampo_M, vCampo_N, vCampo_O, vCampo_P, vCampo_Q, vCampo_R, vCampo_S, vCampo_T;
            string queryinsert, queryvalues;

            // leer las celdas
            int nFilaNada = 0, nFilaAlgo = 0;
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;

            if (cols > 20) cols = 20;

            Process currentProcess = Process.GetCurrentProcess();
            int vId_Archivo = Convert.ToInt32(currentProcess.Id);

            DateTime hoy = DateTime.Now;

            InsUpdDel_Oracle(conexion, "DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo);
            for (int row = 1; row <= rows; row++)
            {
                vCampo_A = ""; vCampo_B = ""; vCampo_C = ""; vCampo_D = "";
                vCampo_E = ""; vCampo_F = ""; vCampo_G = ""; vCampo_H = "";
                vCampo_I = ""; vCampo_J = ""; vCampo_K = ""; vCampo_L = "";
                vCampo_M = ""; vCampo_N = ""; vCampo_O = ""; vCampo_P = "";
                vCampo_Q = ""; vCampo_R = ""; vCampo_S = ""; vCampo_T = "";

                for (int col = 1; col <= cols; col++)
                {
                    // lectura como cadena
                    var cellText = xlWorkSheet.Cells[row, col].Text;
                    cellText = Convert.ToString(cellText);
                    cellText = cellText.Replace("'", ""); // Comillas simples no pueden pasar en el Texto

                    switch (col)
                    {
                        case 1: vCampo_A = cellText; break;
                        case 2: vCampo_B = cellText; break;
                        case 3: vCampo_C = cellText; break;
                        case 4: vCampo_D = cellText; break;
                        case 5: vCampo_E = cellText; break;
                        case 6: vCampo_F = cellText; break;
                        case 7: vCampo_G = cellText; break;
                        case 8: vCampo_H = cellText; break;
                        case 9: vCampo_I = cellText; break;
                        case 10: vCampo_J = cellText; break;
                        case 11: vCampo_K = cellText; break;
                        case 12: vCampo_L = cellText; break;
                        case 13: vCampo_M = cellText; break;
                        case 14: vCampo_N = cellText; break;
                        case 15: vCampo_O = cellText; break;
                        case 16: vCampo_P = cellText; break;
                        case 17: vCampo_Q = cellText; break;
                        case 18: vCampo_R = cellText; break;
                        case 19: vCampo_S = cellText; break;
                        case 20: vCampo_T = cellText; break;
                    }
                }

                if (String.IsNullOrEmpty(vCampo_A.Trim()) && String.IsNullOrEmpty(vCampo_B.Trim()) && String.IsNullOrEmpty(vCampo_C.Trim()) && String.IsNullOrEmpty(vCampo_D.Trim()) && String.IsNullOrEmpty(vCampo_E.Trim()) && String.IsNullOrEmpty(vCampo_F.Trim()) && String.IsNullOrEmpty(vCampo_G.Trim()) && String.IsNullOrEmpty(vCampo_H.Trim()) && String.IsNullOrEmpty(vCampo_I.Trim()) && String.IsNullOrEmpty(vCampo_J.Trim()) && String.IsNullOrEmpty(vCampo_K.Trim()) && String.IsNullOrEmpty(vCampo_L.Trim()) && String.IsNullOrEmpty(vCampo_M.Trim()) && String.IsNullOrEmpty(vCampo_N.Trim()) && String.IsNullOrEmpty(vCampo_O.Trim()) && String.IsNullOrEmpty(vCampo_P.Trim()) && String.IsNullOrEmpty(vCampo_Q.Trim()) && String.IsNullOrEmpty(vCampo_R.Trim()) && String.IsNullOrEmpty(vCampo_S.Trim()) && String.IsNullOrEmpty(vCampo_T.Trim()))
                {
                    nFilaNada++;
                }
                else
                {
                    nFilaAlgo = row;
                    nFilaNada = 0;
                }

                if (nFilaNada > 10)
                    rows = row++;

                //Caracter  indica Ultima Fila Santander
                if (vCampo_A != "")
                {
                    queryinsert = "INSERT INTO ARCHIVOSCONCIBANCATMP (CAMPO_A, CAMPO_B, CAMPO_C, CAMPO_D, CAMPO_E, CAMPO_F, CAMPO_G, CAMPO_H, CAMPO_I, CAMPO_J, CAMPO_K, CAMPO_L, CAMPO_M, CAMPO_N, CAMPO_O, CAMPO_P, CAMPO_Q, CAMPO_R, CAMPO_S, CAMPO_T, ID_ARCHIVO, NOMBREARCHIVO, TAMANOARCHIVO, ID_FILAS, ESTADO, ARCHIVOVALIDO) ";
                    queryvalues = "VALUES ('" + vCampo_A + "', '" + vCampo_B + "', '" + vCampo_C + "', '" + vCampo_D + "', '" + vCampo_E + "', '" + vCampo_F + "', '" + vCampo_G + "', '" + vCampo_H + "', '" + vCampo_I + "', '" + vCampo_J + "', '" + vCampo_K + "', '" + vCampo_L + "', '" + vCampo_M + "', '" + vCampo_N + "', '" + vCampo_O + "', '" + vCampo_P + "', '" + vCampo_Q + "', '" + vCampo_R + "', '" + vCampo_S + "', '" + vCampo_T + "', " + vId_Archivo + ", '" + xlsFilePath + "', " + filesize + ", " + row + ", " + 0 /* 0 para Estado Carga Inicial */ + ", '" + ArchivoValido + "')";
                    InsUpdDel_Oracle(conexion, queryinsert + queryvalues);
                }

            }
            InsUpdDel_Oracle(conexion, "DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ID_FILAS > " + nFilaAlgo);

            // cerrar
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            // liberar
            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

            Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Validando Banco-Moneda-Cuenta... ");
            Function_Procedure_Oracle(conexion, 2, "PKG_CARGARARCHIVOSAUTO.P_UPD_BUSCA_BANCOMONEDACUENTA", "PIid_archivo", vId_Archivo, "", -1, "", -1);
            string codigobanco = SelectFromWhere(conexion, "SELECT DISTINCT CODIGOBANCO FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ROWNUM = 1", false);

            if (!(String.IsNullOrEmpty(codigobanco)))
            {
                int vCodigoBanco = Convert.ToInt32(codigobanco);
                
                //BBVA
                //eliminar caracter "Espacio Duro" del numero de movimiento (HTML)
                if (vCodigoBanco == 6)
                {
                    Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Removiendo Espacio Duro en el campo Num. Mvto (BBVA) ");
                    InsUpdDel_Oracle(conexion, "UPDATE ARCHIVOSCONCIBANCATMP SET CAMPO_E = REPLACE(CAMPO_E, ' ', '') WHERE ID_ARCHIVO = " + vId_Archivo);
                }


                Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Buscando Tipo Carga... ");
                int vTipoCarga = Convert.ToInt32(Function_Procedure_Oracle(conexion, 1, "PKG_CARGARARCHIVOSAUTO.F_OBT_BUSCA_TIPOCARGABANCO", "PIid_archivo", vId_Archivo, "PIcodigobanco", vCodigoBanco, "", -1));
                int vEstadoTipoCarga = 4;
                if (vTipoCarga > 0)
                    vEstadoTipoCarga = 3;
                InsUpdDel_Oracle(conexion, "UPDATE ARCHIVOSCONCIBANCATMP SET TIPOCARGA = " + vTipoCarga + ", ESTADO = " + vEstadoTipoCarga + " WHERE ID_ARCHIVO = " + vId_Archivo + " AND CODIGOBANCO = " + vCodigoBanco);


                Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Buscando Parametros... ");
                int vParametros = Convert.ToInt32(Function_Procedure_Oracle(conexion, 1, "PKG_CARGARARCHIVOSAUTO.F_UPD_BUSCA_EXISTEPARAMETRO", "PIid_archivo", vId_Archivo, "PIcodigobanco", vCodigoBanco, "PItipocarga", vTipoCarga));
                int vEstadoParametros = 6;
                if (vParametros > 0)
                    vEstadoParametros = 5;
                InsUpdDel_Oracle(conexion, "UPDATE ARCHIVOSCONCIBANCATMP SET ESTADO = " + vEstadoParametros + " WHERE ID_ARCHIVO = " + vId_Archivo + " AND CODIGOBANCO = " + vCodigoBanco + " AND TIPOCARGA = " + vTipoCarga);

                Console.Write("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Buscando Num. Cta... ");
                string vCodigoCuenta = SelectFromWhere(conexion, "SELECT DISTINCT NUMEROCUENTA FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ROWNUM = 1", false);
                Console.WriteLine(vCodigoCuenta);

                if (vEstadoParametros % 2 == 1)
                {
                    Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": APROBADO");
                }
                else
                {
                    Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": RECHAZADO");
                    return;
                }


                if (!(String.IsNullOrEmpty(vCodigoCuenta)))
                {
                    Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Insertando en CONCARGAPRIMERATMP... ");
                    Function_Procedure_Oracle(conexion, 2, "PKG_CARGARARCHIVOSAUTO.P_GEN_CONCARGAPRIMERATMP", "PIid_archivo", vId_Archivo, "", -1, "", -1);

                    //Sin parametros
                    //Function_Procedure_Oracle(conexion, 2, "PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA", "", -1, "", -1, "", -1);

                    using (OracleConnection connection =
                           new OracleConnection(conexion))
                    {
                        try
                        {
                            Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Generando Caja y Conciliando... ");

                            OracleCommand command = new OracleCommand("PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA", connection) { CommandType = CommandType.StoredProcedure };

                            var prm1 = new OracleParameter("PIvalarchivocarga", OracleDbType.Varchar2) { Direction = ParameterDirection.Input, Value = Path.GetFileName(xlsFilePath) };
                            command.Parameters.Add(prm1);

                            connection.Open();

                            command.ExecuteNonQuery();

                            connection.Close();
                            command.Dispose();

                            Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Fin.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }

                    //
                }
                else
                {
                    Console.WriteLine("                  >>> Error: Codigo Cuenta Errada");
                }
            }
            else
            {
                Console.WriteLine("                  >>> RECHAZADO. Banco no encontrado, revisar extracto descargado y parametros en SISGO");
            }
        }
        public static void Readed_file(string xlsFilePath, string rutaoutput, string conexion)
        {
            Process currentProcess = Process.GetCurrentProcess();
            int vId_Archivo = currentProcess.Id;

            string queryString = "SELECT DISTINCT acbtmp.ID_ARCHIVO, " +
                    "(CASE MOD(acbtmp.ESTADO, 2) WHEN 1 THEN 'APROBADO' ELSE 'RECHAZADO' END) APROBADO, " +
                    "(CASE WHEN LENGTH(pkg_syst900.F_OBT_TBLDESCRI(39, acbtmp.codigobanco)) > 0 THEN pkg_syst900.F_OBT_TBLDESCRI(39, acbtmp.codigobanco) WHEN LENGTH(pkg_syst900.F_OBT_TBLDESCRI(39, acbtmp.codigobanco)) = 0 THEN '' END) BANCO, " +
                    "(CASE WHEN LENGTH(pkg_syst900.F_OBT_TBLDESCRI(22, acbtmp.moneda)) > 0 THEN pkg_syst900.F_OBT_TBLDESCRI(22, acbtmp.moneda) WHEN LENGTH(pkg_syst900.F_OBT_TBLDESCRI(22, acbtmp.moneda)) = 0 THEN '' END) MONEDA, " +
                    "(CASE WHEN LENGTH(acbtmp.numerocuenta) > 0 THEN acbtmp.numerocuenta WHEN LENGTH(acbtmp.numerocuenta) = 0 THEN '' END) CUENTA, " +
                    "(CASE WHEN acbtmp.ESTADO = 1 OR acbtmp.ESTADO = 2 THEN 'BANCO-MONEDA-CUENTA' WHEN acbtmp.ESTADO = 3 OR acbtmp.ESTADO = 4 THEN 'TIPO-CARGA' WHEN acbtmp.ESTADO = 5 OR acbtmp.ESTADO = 6 THEN 'PARAMETROS' WHEN acbtmp.ESTADO = 7 THEN 'PROCESADO' ELSE 'SIN_INFORMACION' END) ESTADO, " +
                    "CASE WHEN LENGTH((Select Distinct Max(ct.secuencarga) From concargaprimeratmp ct Where ct.fechacarga = trunc(acbtmp.fechacarga) And ct.codigobanco = acbtmp.codigobanco)) > 0 THEN (Select Distinct Decode(Max(ct.secuencarga), Null, 0, Max(ct.secuencarga)) From concargaprimeratmp ct Where ct.fechacarga = trunc(acbtmp.fechacarga) And ct.codigobanco = acbtmp.codigobanco) ELSE (CASE MOD(acbtmp.ESTADO, 2) WHEN 1 THEN 0 ELSE NULL END) END MAXIDBANCO, " +
                    "TO_CHAR(SYSDATE, 'dd-mm-YYYY_HH24MISS') FECHA, " +
                    "SUBSTR(acbtmp.NOMBREARCHIVO, (INSTR(acbtmp.NOMBREARCHIVO, '.', -1, 1) + 1), length(acbtmp.NOMBREARCHIVO)) EXTENSION " +
                    "FROM archivosconcibancatmp acbtmp " +
            "WHERE ID_ARCHIVO = " + vId_Archivo;
            string vfileoutput;

            using (OracleConnection connection =
                   new OracleConnection(conexion))
            {
                OracleCommand command = connection.CreateCommand();
                command.CommandText = queryString;
                try
                {
                    string subPath = DateTime.Now.ToString("dd-MM-yyyy");

                    if (!Directory.Exists(rutaoutput + subPath))
                        Directory.CreateDirectory(rutaoutput + subPath);

                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        vfileoutput = rutaoutput + subPath + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(xlsFilePath); //Path.GetFileName
                        vfileoutput += "_" + reader["APROBADO"].ToString();
                        vfileoutput += "_" + reader["BANCO"].ToString().Replace(" ", "_");
                        vfileoutput += "_" + reader["MONEDA"].ToString();
                        vfileoutput += "_" + reader["CUENTA"].ToString();
                        vfileoutput += "_" + reader["ESTADO"].ToString();
                        vfileoutput += "_" + reader["MAXIDBANCO"].ToString();
                        vfileoutput += "_" + reader["FECHA"].ToString();
                        vfileoutput += "." + reader["EXTENSION"].ToString();
                        if (File.Exists(xlsFilePath))
                        {
                            File.Move(xlsFilePath, vfileoutput);
                        }
                        InsUpdDel_Oracle(conexion, "DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to release the object(object:{0})\n" + ex.Message, obj.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public static string SelectFromWhere(string conexion, string executequery, bool eslike)
        {
            string vDATO = "";
            string queryString = executequery.ToString();

            using (OracleConnection connection =
                   new OracleConnection(conexion))
            {
                OracleCommand command = connection.CreateCommand();
                command.CommandText = queryString;
                try
                {
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        vDATO = reader[0].ToString();
                        if (eslike) break;
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    vDATO = null;
                    Console.WriteLine(ex.Message);
                }
            }
            return vDATO;
        }
        public static void InsUpdDel_Oracle(string conexion, string executequery)
        {
            string queryString = executequery.ToString();
            string queryCommit = "COMMIT";

            using (OracleConnection connection =
                   new OracleConnection(conexion))
            {
                OracleCommand command = connection.CreateCommand();
                try
                {
                    connection.Open();
                    command.CommandType = CommandType.Text;
                    command.CommandText = queryString;
                    command.ExecuteNonQuery();
                    command.CommandText = queryCommit;
                    command.ExecuteNonQuery();

                    connection.Close();
                    command.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public static string Function_Procedure_Oracle(string conexion, int tipofunpro /* tipofunpro: 1 para function, 2 para procedure  */, string executequery, string nomprmt1, int prmt1, string nomprmt2, int prmt2, string nomprmt3, int prmt3)
        {
            string vDATO = "";
            using (OracleConnection connection =
                   new OracleConnection(conexion))
            {
                try
                {
                    OracleCommand command = new OracleCommand(executequery, connection) { CommandType = CommandType.StoredProcedure };

                    if (tipofunpro == 1)
                    {
                        var returnVal = new OracleParameter("Return_Value", OracleDbType.Int32) { Direction = ParameterDirection.ReturnValue };
                        command.Parameters.Add(returnVal);
                    }

                    if (prmt1 >= 0)
                    {
                        var prm1 = new OracleParameter(nomprmt1, OracleDbType.Int32) { Direction = ParameterDirection.Input , Value = prmt1 };
                        command.Parameters.Add(prm1);
                    }
                    if (prmt2 >= 0)
                    {
                        var prm2 = new OracleParameter(nomprmt2, OracleDbType.Int32) { Direction = ParameterDirection.Input, Value = prmt2 };
                        command.Parameters.Add(prm2);
                    }
                    if (prmt3 >= 0)
                    {
                        var prm3 = new OracleParameter(nomprmt3, OracleDbType.Int32) { Direction = ParameterDirection.Input , Value = prmt3 };
                        command.Parameters.Add(prm3);
                    }

                    connection.Open();
                    command.ExecuteNonQuery();

                    if (tipofunpro == 1)
                    {
                        vDATO = Convert.ToString(command.Parameters["Return_Value"].Value);
                    }

                    connection.Close();
                    command.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return vDATO;
        }
        public static void ArgInvalid(bool pause, bool invalid)
        {
            if (invalid)
            {
                Console.WriteLine("Argumentos Invalidos");
            }
            Console.WriteLine(" Ayuda:");
            Console.WriteLine("  -nopause: Finaliza la aplicacion al terminar las operaciones. Por defecto, pausa la aplicacion");
            Console.WriteLine("  -hide: Oculta la consola");
            Console.WriteLine("  -killall: Termina todos los procesos en ejecucion de nombre LectorExcelConciliacion.exe");
            Console.WriteLine("  -file <ruta>: Procesa el archivo <ruta>");
            Console.WriteLine("  Sin parametro -file: Procesa todos los archivos en la carpeta INPUT");
            if (pause)
            {
                Console.WriteLine("Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
            Environment.Exit(0);
        }
        public static string ObtRuta(string conexion, int tblcodarg, string name)
        {
            string obtRuta = SelectFromWhere(conexion, "SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLESTADO = 1 AND TBLCODARG IN (" + tblcodarg + ")", false);
            if (obtRuta == null)
            {
                Console.WriteLine("Error obteniendo la ruta " + name);
                Environment.Exit(0);
            }
            else if (!Directory.Exists(obtRuta))
            {
                Console.WriteLine("Error: Ruta " + name + " no encontrada\nRuta: " + obtRuta);
                Environment.Exit(0);
            }
            return obtRuta + Path.DirectorySeparatorChar;
        }
        public static bool BorrarSubWork(string rutawork, int id, bool msg)
        {
            try
            {
                Directory.Delete(rutawork, true);
                return true;
            }
            catch (Exception ex)
            {
                if (msg)
                {
                    Console.WriteLine("Error eliminando carpeta WORK" + Path.DirectorySeparatorChar + id + Path.DirectorySeparatorChar);
                    Console.WriteLine(ex.Message);
                }
                return false;
            }
        }
    }
}
