using System;
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

            OracleFunctions oracleFunctions = new OracleFunctions(conexion, null);
            //ruta INPUT
            string rutainput = oracleFunctions.ObtRuta(14, "INPUT");
            //

            //ruta OUTPUT
            string rutaoutput = oracleFunctions.ObtRuta(15, "OUTPUT");
            //

            //ruta WORK
            string rutawork = oracleFunctions.ObtRuta(29, "WORK");

            //inicio log
            LogWriter logWriter = new LogWriter(rutawork);

            try
            {
                logWriter.addLog("Inicio", false);
                logWriter.LogWrite();
                //Sub carpeta dentro de WORK
                string rutasubwork = rutawork + currentProcess.Id + Path.DirectorySeparatorChar;
                if (!Directory.Exists(rutasubwork))
                {
                    try
                    {
                        Directory.CreateDirectory(rutasubwork);
                    }
                    catch (Exception ex2)
                    {
                        Console.WriteLine("Error creando carpeta WORK" + Path.DirectorySeparatorChar + currentProcess.Id + Path.DirectorySeparatorChar);
                        Console.WriteLine(ex2.Message);
                        logWriter.addLog(ex2.Message, true);
                        logWriter.addLog("Fin", false);
                        logWriter.LogWrite();
                        Environment.Exit(0);
                    }
                }
                else
                {
                    BorrarSubWork(rutasubwork, currentProcess.Id, false, logWriter);
                    Directory.CreateDirectory(rutasubwork);
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
                            BorrarSubWork(rutasubwork, currentProcess.Id, true, logWriter);
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
                                    logWriter.addLog(" LectorExcelConciliacion con PID " + process.Id + " cerrado.", false);
                                }
                            }
                            BorrarSubWork(rutasubwork, currentProcess.Id, true, logWriter);
                            logWriter.addLog("Fin", false);
                            logWriter.LogWrite();
                            Environment.Exit(0);
                        }
                        else if (args[0 + offset] == "-file")
                        {
                            fileindividual = true;
                            if (args[1 + offset] == null)
                            {
                                BorrarSubWork(rutasubwork, currentProcess.Id, true, logWriter);
                                logWriter.LogWrite();
                                ArgInvalid(false, true);
                            }
                            if (File.Exists(args[1 + offset]))
                            {
                                Console.Write(" Moviendo archivo a ruta Work... ");
                                string nombrearchivo = Path.Combine(rutasubwork, Path.GetFileNameWithoutExtension(args[1 + offset]) + "_" + currentProcess.Id + Path.GetExtension(args[1 + offset]));
                                File.Move(args[1 + offset], nombrearchivo);
                                Console.WriteLine(" OK");
                                logWriter.addLog("Se movió " + Path.GetFileName(nombrearchivo) + " a SubWork", false);
                                logWriter.LogWrite();
                                Console.WriteLine("    Procesando >> " + nombrearchivo);
                                ExecuteExcel(Path.GetFileName(nombrearchivo), rutasubwork, rutaoutput, conexion, logWriter);
                            }
                            else
                            {
                                Console.WriteLine(" Archivo no encontrado\n Ruta: " + args[1 + offset]);
                                logWriter.addLog(" Archivo no encontrado. Ruta: " + args[1 + offset], true);
                                logWriter.LogWrite();
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
                        BorrarSubWork(rutasubwork, currentProcess.Id, true, logWriter);
                        if (offset == 0)
                        {
                            logWriter.LogWrite();
                            ArgInvalid(true, true);
                        }
                        else
                        {
                            logWriter.LogWrite();
                            ArgInvalid(pause, true);
                        }
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
                        logWriter.addLog("Se encontró " + dirs.Length + " archivos en la ruta: " + rutainput, false);
                        logWriter.LogWrite();

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

                        BorrarSubWork(rutasubwork, currentProcess.Id, true, logWriter);
                        logWriter.addLog("No se encontró archivos en la ruta: " + rutainput, true);
                        logWriter.addLog("Fin", false);
                        logWriter.LogWrite();
                        Environment.Exit(0);
                    }
                    Console.WriteLine(" Moviendo archivos a ruta Work...");
                    foreach (string dir in dirs)
                    {
                        string nombrearchivo = Path.Combine(rutasubwork, Path.GetFileNameWithoutExtension(dir) + "_" + currentProcess.Id + Path.GetExtension(dir));
                        File.Move(dir, nombrearchivo);
                        logWriter.addLog("Se movió " + Path.GetFileName(nombrearchivo) + " a SubWork", false);
                        //logWriter.LogWrite();
                    }

                    string[] dirswork = Directory.GetFiles(rutasubwork);
                    int count = 1;
                    foreach (string dir in dirswork)
                    {
                        Console.WriteLine(" Procesando " + count + "/" + dirswork.Length + " >> " + Path.GetFileName(dir));
                        ExecuteExcel(Path.GetFileName(dir), rutasubwork, rutaoutput, conexion, logWriter);
                        count++;
                    }
                }

                BorrarSubWork(rutasubwork, currentProcess.Id, true, logWriter);
                Console.WriteLine("Fin");
                logWriter.addLog("Fin", false);
                logWriter.LogWrite();
                if (pause)
                {
                    Console.WriteLine("Presione cualquier tecla para salir...");
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                logWriter.addLog(ex.Message, true);
                logWriter.LogWrite();
            }
        }

        static void ExecuteExcel(string filename, string rutasubwork, string rutaoutput, string conexion, LogWriter logWriter)
        {
            try
            {
                OracleFunctions oracleFunctions = new OracleFunctions(conexion, logWriter);
                string varchivovalido = oracleFunctions.SelectFromWhere("SELECT SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1) " +
                                                        "FROM concargarchivos " +
                                                        "WHERE SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1) IS NOT NULL " +
                                                        "AND UPPER('" + filename + "') " +
                                                        "LIKE '%'||SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1)||'%' " +
                                                        "GROUP BY nombrearchivocarga " +
                                                        "ORDER BY nombrearchivocarga", true);
                if (!(String.IsNullOrEmpty(varchivovalido)))
                {
                    logWriter.addLog("Procesando " + filename, false);
                    //logWriter.LogWrite();
                    Read_file(Path.Combine(rutasubwork, filename), varchivovalido, conexion, logWriter);
                    Readed_file(Path.Combine(rutasubwork, filename), rutaoutput, conexion, logWriter);
                }
                else
                {
                    Console.WriteLine(" >>> Rechazado: Nombre Archivo no se encuentra parametrizado en la tabla CONCARGARCHIVOS");
                    logWriter.addLog("Nombre Archivo no se encuentra parametrizado en la tabla CONCARGARCHIVOS", true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                logWriter.addLog(ex.Message, true);
                logWriter.LogWrite();
            }
        }
        public static void Read_file(string xlsFilePath, string ArchivoValido, string conexion, LogWriter logWriter)
        {
            try
            {
                if (!File.Exists(xlsFilePath))
                    return;

                Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Lectura y Escritura del excel... ");
                logWriter.addLog("Lectura y Escritura del excel", false);
                logWriter.LogWrite();
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

                OracleFunctions oracleFunctions = new OracleFunctions(conexion, logWriter);
                oracleFunctions.InsUpdDel_Oracle("DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo);

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
                        oracleFunctions.InsUpdDel_Oracle(queryinsert + queryvalues);
                    }

                }
                oracleFunctions.InsUpdDel_Oracle("DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ID_FILAS > " + nFilaAlgo);

                // cerrar
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                // liberar
                ReleaseObject(xlWorkSheet, logWriter);
                ReleaseObject(xlWorkBook, logWriter);
                ReleaseObject(xlApp, logWriter);

                Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Validando Banco-Moneda-Cuenta... ");
                logWriter.addLog("Validando Banco-Moneda-Cuenta", false);
                logWriter.LogWrite();
                oracleFunctions.Function_Procedure_Oracle(2, "PKG_CARGARARCHIVOSAUTO.P_UPD_BUSCA_BANCOMONEDACUENTA", "PIid_archivo", vId_Archivo, "", -1, "", -1);
                string codigobanco = oracleFunctions.SelectFromWhere("SELECT DISTINCT CODIGOBANCO FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ROWNUM = 1", false);

                if (!(String.IsNullOrEmpty(codigobanco)))
                {
                    int vCodigoBanco = Convert.ToInt32(codigobanco);

                    //BBVA
                    //eliminar caracter "Espacio Duro" del numero de movimiento (HTML)
                    if (vCodigoBanco == 6)
                    {
                        Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Removiendo Espacio Duro en el campo Num. Mvto (BBVA) ");
                        logWriter.addLog("Removiendo Espacio Duro en el campo Num. Mvto (BBVA)", false);
                        //logWriter.LogWrite();
                        oracleFunctions.InsUpdDel_Oracle("UPDATE ARCHIVOSCONCIBANCATMP SET CAMPO_E = REPLACE(CAMPO_E, ' ', '') WHERE ID_ARCHIVO = " + vId_Archivo);
                    }


                    Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Buscando Tipo Carga... ");
                    logWriter.addLog("Buscando Tipo Carga", false);
                    logWriter.LogWrite();
                    int vTipoCarga = Convert.ToInt32(oracleFunctions.Function_Procedure_Oracle(1, "PKG_CARGARARCHIVOSAUTO.F_OBT_BUSCA_TIPOCARGABANCO", "PIid_archivo", vId_Archivo, "PIcodigobanco", vCodigoBanco, "", -1));
                    int vEstadoTipoCarga = 4;
                    if (vTipoCarga > 0)
                        vEstadoTipoCarga = 3;
                    oracleFunctions.InsUpdDel_Oracle("UPDATE ARCHIVOSCONCIBANCATMP SET TIPOCARGA = " + vTipoCarga + ", ESTADO = " + vEstadoTipoCarga + " WHERE ID_ARCHIVO = " + vId_Archivo + " AND CODIGOBANCO = " + vCodigoBanco);


                    Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Buscando Parametros... ");
                    logWriter.addLog("Buscando Parametros", false);
                    logWriter.LogWrite();
                    int vParametros = Convert.ToInt32(oracleFunctions.Function_Procedure_Oracle(1, "PKG_CARGARARCHIVOSAUTO.F_UPD_BUSCA_EXISTEPARAMETRO", "PIid_archivo", vId_Archivo, "PIcodigobanco", vCodigoBanco, "PItipocarga", vTipoCarga));
                    int vEstadoParametros = 6;
                    if (vParametros > 0)
                        vEstadoParametros = 5;
                    oracleFunctions.InsUpdDel_Oracle("UPDATE ARCHIVOSCONCIBANCATMP SET ESTADO = " + vEstadoParametros + " WHERE ID_ARCHIVO = " + vId_Archivo + " AND CODIGOBANCO = " + vCodigoBanco + " AND TIPOCARGA = " + vTipoCarga);

                    Console.Write("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Buscando Num. Cta... ");
                    logWriter.addLog("Buscando Num. Cta", false);
                    //logWriter.LogWrite();
                    string vCodigoCuenta = oracleFunctions.SelectFromWhere("SELECT DISTINCT NUMEROCUENTA FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ROWNUM = 1", false);
                    Console.WriteLine(vCodigoCuenta);
                    logWriter.addLog("Numero de Cuenta " + vCodigoCuenta, false);
                    logWriter.LogWrite();

                    if (vEstadoParametros % 2 == 1)
                    {
                        Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": APROBADO");
                        logWriter.addLog("APROBADO", false);
                        logWriter.LogWrite();
                    }
                    else
                    {
                        Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": RECHAZADO");
                        logWriter.addLog("RECHAZADO", true);
                        logWriter.LogWrite();
                        return;
                    }


                    if (!(String.IsNullOrEmpty(vCodigoCuenta)))
                    {
                        Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Insertando en CONCARGAPRIMERATMP... ");
                        logWriter.addLog("Insertando en CONCARGAPRIMERATMP", false);
                        logWriter.LogWrite();
                        oracleFunctions.Function_Procedure_Oracle(2, "PKG_CARGARARCHIVOSAUTO.P_GEN_CONCARGAPRIMERATMP", "PIid_archivo", vId_Archivo, "", -1, "", -1);

                        //Sin parametros
                        //Function_Procedure_Oracle(conexion, 2, "PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA", "", -1, "", -1, "", -1);

                        using (OracleConnection connection =
                               new OracleConnection(conexion))
                        {
                            Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Generando Caja y Conciliando... ");
                            logWriter.addLog("Generando Caja y Conciliando", false);
                            logWriter.LogWrite();

                            OracleCommand command = new OracleCommand("PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA", connection) { CommandType = CommandType.StoredProcedure };

                            var prm1 = new OracleParameter("PIvalarchivocarga", OracleDbType.Varchar2) { Direction = ParameterDirection.Input, Value = Path.GetFileName(xlsFilePath) };
                            command.Parameters.Add(prm1);

                            bool reintentar = true;
                            int cantidadintentos = 0;
                            connection.Open();
                            while (reintentar && cantidadintentos < 5)
                            {
                                cantidadintentos++;
                                try
                                {
                                    command.ExecuteNonQuery();
                                    reintentar = false;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Intento "+ cantidadintentos + ": " + ex.Message);
                                    logWriter.addLog("Intento " + cantidadintentos + ": " + ex.Message, true);
                                    logWriter.LogWrite();
                                }
                            }

                            connection.Close();
                            command.Dispose();

                            Console.WriteLine("                  >>> " + DateTime.Now.ToString("HH:mm:ss") + ": Fin.");
                            logWriter.addLog("Archivo Procesado", false);
                            //logWriter.LogWrite();
                        }
                        //
                    }
                    else
                    {
                        Console.WriteLine("                  >>> Error: Codigo Cuenta Errada");
                        Console.WriteLine("Codigo Cuenta Errada");
                    }
                }
                else
                {
                    Console.WriteLine("                  >>> RECHAZADO. Banco no encontrado, revisar extracto descargado y parametros en SISGO");
                    logWriter.addLog("Banco no encontrado, revisar extracto descargado y parametros en SISGO", true);
                    logWriter.LogWrite();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                logWriter.addLog(ex.Message, true);
                logWriter.LogWrite();
            }
            
        }
        public static void Readed_file(string xlsFilePath, string rutaoutput, string conexion, LogWriter logWriter)
        {
            int vId_Archivo = 0;
            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                vId_Archivo = currentProcess.Id;

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
                        OracleFunctions oracleFunctions = new OracleFunctions(conexion, logWriter);
                        oracleFunctions.InsUpdDel_Oracle("DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo);
                    }
                    reader.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                logWriter.addLog(ex.Message, true);
                try
                {
                    OracleFunctions oracleFunctions = new OracleFunctions(conexion, logWriter);
                    oracleFunctions.InsUpdDel_Oracle("DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo);
                }
                catch (Exception ex2)
                {
                    Console.WriteLine(ex2.Message);
                    logWriter.addLog(ex2.Message, true);
                }
                logWriter.LogWrite();
            }
        }
        public static void ReleaseObject(object obj, LogWriter logWriter)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to release the object(object:{0})\n" + ex.Message, obj.ToString());
                Console.WriteLine(ex.Message);
                logWriter.addLog(ex.Message, true);
            }
            finally
            {
                GC.Collect();
            }
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
        public static bool BorrarSubWork(string rutasubwork, int id, bool msg, LogWriter logWriter)
        {
            try
            {
                Directory.Delete(rutasubwork, true);
                return true;
            }
            catch (Exception ex)
            {
                if (msg)
                {
                    Console.WriteLine("Error eliminando carpeta WORK" + Path.DirectorySeparatorChar + id + Path.DirectorySeparatorChar);
                    Console.WriteLine(ex.Message);
                    logWriter.addLog("Error eliminando carpeta SubWork" + Path.DirectorySeparatorChar + id + Path.DirectorySeparatorChar, true);
                }
                return false;
            }
        }
    }
}
