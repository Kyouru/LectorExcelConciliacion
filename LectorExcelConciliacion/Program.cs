﻿using System;
using System.IO;
using System.Configuration;
using System.Data;
using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;

namespace LectorExcelConciliacion
{
    class Program
    {
        static void Main(string[] args)
        {

            Process currentProcess = Process.GetCurrentProcess();
            Console.Title = "Lector Excel Conciliacion (PID: "+ currentProcess.Id + ")";

            //Importante:
            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@     @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@@@@@@,                @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@     &@@@@@@@@@@@  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@    @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@   (@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@   /@@@@@@@@@@@@@@          ###@@@  @@*@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@   @@@@@@@@@@@@       @@@@@@@@@@@@@@@@@@@@@. ,@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@   @@@@@@@@@@@      @@@@@@@@@@@@@@@@@@@@@@@@@@@@@  @@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@  #@@@@@@@@@      @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ (@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@#  @@@@@@@@@      @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@*  @@@@@@@@@      @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@, @@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@&  @@@@@@@@       @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@");
            Console.WriteLine("@@@  @@@@@@@@.      @@@@@@@@@@@@@@@@@@@  *@@@@@@@@@@@@@@@@@@@@@@@  @@@@@@@@@@@@@");
            Console.WriteLine("@@,  @@@@@@@@      @@@@@@@@@@@@@@@@@@@@   @@@@@@@@@@@@@@@@@@@@@@@#  @@@@@@@@@@@@");
            Console.WriteLine("@@  (@@@@@@@       @@@@@@@@@@@@@@@@@@@     @@@@@@@@@@@@@@@@@@@@@@@  @@@@@@@@@@@@");
            Console.WriteLine("@@  @@@@@@@@       @@@@@@@@@@@@@@              @@@@@@@@@@@@@@@@@@@  (@@@@@@@@@@@");
            Console.WriteLine("@@  @@@@@@@@       .@@@@@@@@@@@@@@            ,@@@@@@@@@@@@@@@@@@   /@@@@@@@   @");
            Console.WriteLine("@@  #@@@@@@@        @@@@@@@@@@@@@@@@@@     @@@@@@@@@@@@@@@@@@@@@@   @@@@@@@@  ,@");
            Console.WriteLine("@@,  @@@@@@@@        @@@@@@@@@@@@@@@@@@  *@@@@@@@@@@@@@@@@@@@@@@    @@@@@@@@  @@");
            Console.WriteLine("@@@  @@@@@@@@.        @@@@@@@@@@@@@@@@@  *@@@@@@@@@@@@@@@@@@@@&    @@@@@@@@   @@");
            Console.WriteLine("@@@&  @@@@@@@@(         @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@     @@@@@@@@/  @@@");
            Console.WriteLine("@@@@(  @@@@@@@@@          @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@      @@@@@@@@%  @@@@");
            Console.WriteLine("@@@@@@  @@@@@@@@@            @@@@@@@@@@@@@@@@@@@@@@@@@@#       @@@@@@@@@/  @@@@@");
            Console.WriteLine("@@@@@@@  *@@@@@@@@@               &@@@@@@@@@@@@@@@@          @@@@@@@@@@   @@@@@@");
            Console.WriteLine("@@@@@@@@   @@@@@@@@@@@                                     @@@@@@@@@@   @@@@@@@@");
            Console.WriteLine("@@@@@@@@@@   @@@@@@@@@@@@                              %@@@@@@@@@@@   @@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@.   @@@@@@@@@@@@@@                    *@@@@@@@@@@@@@@   @@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@    @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@    @@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@    &@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@    *@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@      @@@@@@@@@@@@@@@@@@@@@@@@@     .@@@@@@@@@@@@@@@@@@@@@");
            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@@@@@@@                         @@@@@@@@@@@@@@@@@@@@@@@@@@@");

            //ruta INPUT
            string rutainput = obtRuta(14, "INPUT");
            //

            //ruta OUTPUT
            string rutaoutput = obtRuta(15, "OUTPUT");
            //

            //ruta WORK
            string rutawork = obtRuta(17, "WORK");
            //Sub carpeta dentro de WORK
            rutawork = rutawork + "\\" + currentProcess.Id + "\\";
            if (!Directory.Exists(rutawork))
            {
                try
                {
                    Directory.CreateDirectory(rutawork);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error creando carpeta WORK\\" + currentProcess.Id + "\\");
                    Console.WriteLine(ex.Message);
                }
            }
            //

            int offset = 0;
            bool pause = false;
            while (args.Length - offset > 0)
            {
                if (args[0 + offset] == "-killall" || args[0 + offset] == "-file" || args[0 + offset] == "-pauseend")
                {
                    if (args[0 + offset] == "-pauseend")
                    {
                        pause = true;
                        offset++;
                    }
                    if (args[0 + offset] == "-killall")
                    {
                        foreach (var process in Process.GetProcessesByName("LectorExcelConciliacion"))
                        {
                            if (process.Id != currentProcess.Id)
                            {
                                process.Kill();
                                Console.WriteLine("LectorExcelConciliacion con PID " + process.Id + " cerrado.");
                            }
                        }
                        offset++;
                    }
                    else if (args[0 + offset] == "-file")
                    {
                        if (args[1 + offset] != null)
                        {
                            argInvalid(pause);
                        }
                        if (File.Exists(args[1 + offset]))
                        {
                            Console.Write("Procesando >> " + args[1 + offset]);
                            ExecuteExcel(Path.GetFileName(args[1 + offset]), Path.GetDirectoryName(args[1 + offset]), rutaoutput);
                        }
                        else
                        {
                            Console.WriteLine("Archivo no encontrado\nRuta: " + args[1 + offset]);
                        }
                        offset = offset + 2;
                    }
                }
                else
                {
                    argInvalid(pause);
                }
            }
            
            //Caso sin parametros, revisa INPUT
            if (args.Length == 0)
            {
                string[] dirs = Directory.GetDirectories(rutainput);

                if (dirs.Length > 0)
                {
                    Console.WriteLine("Se encontró " + dirs.Length + " archivos");
                }
                Console.WriteLine(" Copiando achivos a ruta Work...");
                foreach (string dir in dirs)
                {
                    File.Copy(Path.Combine(rutainput, dir), rutawork, true);
                }

                string[] dirswork = Directory.GetDirectories(rutawork);
                foreach (string dir in dirswork)
                {
                    Console.WriteLine(" Procesando 1/" + dirswork.Length + " >> " + dir);
                    ExecuteExcel(dir, rutawork, rutaoutput);
                }
                try
                {
                    Directory.Delete(rutawork, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error eliminando carpeta WORK\\" + currentProcess.Id + "\\");
                    Console.WriteLine(ex.Message);
                }
            }

            Console.WriteLine("Fin");
            if (pause)
            {
                Console.ReadKey();
            }
        }

        static void ExecuteExcel(string filename, string rutawork, string rutaoutput)
        {
            string varchivovalido = SelectFromWhere("SELECT SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1) " +
                                                    "FROM concargarchivos " +
                                                    "WHERE SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1) IS NOT NULL " +
                                                    "AND UPPER('" + filename + "') " +
                                                    "LIKE '%'||SUBSTR(nombrearchivocarga, 1, INSTR(nombrearchivocarga, '.', 1, 1) - 1)||'%' " +
                                                    "GROUP BY nombrearchivocarga " +
                                                    "ORDER BY nombrearchivocarga", true);

            if (!(String.IsNullOrEmpty(varchivovalido)))
            {
                read_file(Path.Combine(rutawork, filename), varchivovalido);
                readed_file(filename, rutaoutput);
            }
        }

        public static void read_file(string xlsFilePath, string ArchivoValido)
        {
            if (!File.Exists(xlsFilePath))
                return;

            FileInfo fi = new FileInfo(xlsFilePath);
            long filesize = fi.Length;

            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Range range;
            var misValue = Type.Missing;//System.Reflection.Missing.Value;

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
            int vId_Archivo = currentProcess.Id;

            DateTime hoy = DateTime.Now;

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

                queryinsert = "INSERT INTO ARCHIVOSCONCIBANCATMP (CAMPO_A, CAMPO_B, CAMPO_C, CAMPO_D, CAMPO_E, CAMPO_F, CAMPO_G, CAMPO_H, CAMPO_I, CAMPO_J, CAMPO_K, CAMPO_L, CAMPO_M, CAMPO_N, CAMPO_O, CAMPO_P, CAMPO_Q, CAMPO_R, CAMPO_S, CAMPO_T, ID_ARCHIVO, NOMBREARCHIVO, TAMANOARCHIVO, ID_FILAS, ESTADO, ARCHIVOVALIDO) ";
                queryvalues = "VALUES ('" + vCampo_A + "', '" + vCampo_B + "', '" + vCampo_C + "', '" + vCampo_D + "', '" + vCampo_E + "', '" + vCampo_F + "', '" + vCampo_G + "', '" + vCampo_H + "', '" + vCampo_I + "', '" + vCampo_J + "', '" + vCampo_K + "', '" + vCampo_L + "', '" + vCampo_M + "', '" + vCampo_N + "', '" + vCampo_O + "', '" + vCampo_P + "', '" + vCampo_Q + "', '" + vCampo_R + "', '" + vCampo_S + "', '" + vCampo_T + "', " + vId_Archivo + ", '" + xlsFilePath + "', " + filesize + ", " + row + ", " + 0 /* 0 para Estado Carga Inicial */ + ", '" + ArchivoValido + "')";
                InsUpdDel_Oracle(queryinsert + queryvalues);
            }

            InsUpdDel_Oracle("DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ID_FILAS > " + nFilaAlgo);

            Function_Procedure_Oracle(2, "PKG_CARGARARCHIVOSAUTO.P_UPD_BUSCA_BANCOMONEDACUENTA", "PIid_archivo", vId_Archivo, "", -1, "", -1);
            string datobuscado = SelectFromWhere("SELECT DISTINCT CODIGOBANCO FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ROWNUM = 1", false);
            if (!(String.IsNullOrEmpty(datobuscado)))
            {
                int vCodigoBanco = Convert.ToInt32(datobuscado);
                int vTipoCarga = Convert.ToInt32(Function_Procedure_Oracle(1, "PKG_CARGARARCHIVOSAUTO.F_OBT_BUSCA_TIPOCARGABANCO", "PIid_archivo", vId_Archivo, "PIcodigobanco", vCodigoBanco, "", -1));
                int vEstadoTipoCarga = 4;
                if (vTipoCarga > 0)
                    vEstadoTipoCarga = 3;
                InsUpdDel_Oracle("UPDATE ARCHIVOSCONCIBANCATMP SET TIPOCARGA = " + vTipoCarga + ", ESTADO = " + vEstadoTipoCarga + " WHERE ID_ARCHIVO = " + vId_Archivo + " AND CODIGOBANCO = " + vCodigoBanco);
                int vParametros = Convert.ToInt32(Function_Procedure_Oracle(1, "PKG_CARGARARCHIVOSAUTO.F_UPD_BUSCA_EXISTEPARAMETRO", "PIid_archivo", vId_Archivo, "PIcodigobanco", vCodigoBanco, "PItipocarga", vTipoCarga));
                int vEstadoParametros = 6;
                if (vParametros > 0)
                    vEstadoParametros = 5;
                InsUpdDel_Oracle("UPDATE ARCHIVOSCONCIBANCATMP SET ESTADO = " + vEstadoParametros + " WHERE ID_ARCHIVO = " + vId_Archivo + " AND CODIGOBANCO = " + vCodigoBanco + " AND TIPOCARGA = " + vTipoCarga);
                Function_Procedure_Oracle(2, "PKG_CARGARARCHIVOSAUTO.P_GEN_CONCARGAPRIMERATMP", "PIid_archivo", vId_Archivo, "", -1, "", -1);
                Function_Procedure_Oracle(2, "PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA", "", -1, "", -1, "", -1);
            }

            // cerrar
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            // liberar
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        public static void readed_file(string xlsFileName, string rutaoutput)
        {
            string connectionString = ConfigurationManager.ConnectionStrings[Parameters.ambiente].ToString();
            Process currentProcess = Process.GetCurrentProcess();
            int vId_Archivo = currentProcess.Id;

            string queryString = "SELECT DISTINCT acbtmp.ID_ARCHIVO, " +
                    "SUBSTR(acbtmp.NOMBREARCHIVO,(INSTR(acbtmp.NOMBREARCHIVO,'\\',-1)+1)) FILEIN, " +
                    "SUBSTR(SUBSTR(acbtmp.NOMBREARCHIVO, (INSTR(acbtmp.NOMBREARCHIVO, '\\',-1)+1)),1,INSTR(SUBSTR(acbtmp.NOMBREARCHIVO,(INSTR(acbtmp.NOMBREARCHIVO,'\\',-1)+1)),'.',1,1)-1), " +
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
            string vfilework;
            string vfileoutput;

            using (OracleConnection connection =
                   new OracleConnection(connectionString))
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
                        vfilework = reader["FILEIN"].ToString();
                        vfileoutput = rutaoutput + subPath + "\\" + xlsFileName;
                        vfileoutput += "_" + reader["APROBADO"].ToString();
                        vfileoutput += "_" + reader["BANCO"].ToString();
                        vfileoutput += "_" + reader["MONEDA"].ToString();
                        vfileoutput += "_" + reader["CUENTA"].ToString();
                        vfileoutput += "_" + reader["ESTADO"].ToString();
                        vfileoutput += "_" + reader["MAXIDBANCO"].ToString();
                        vfileoutput += "_" + reader["FECHA"].ToString();
                        vfileoutput += "." + reader["EXTENSION"].ToString();

                        if (File.Exists(vfilework))
                        {
                            File.Move(vfilework, vfileoutput);
                        }
                        InsUpdDel_Oracle("DELETE FROM ARCHIVOSCONCIBANCATMP WHERE ID_ARCHIVO = " + vId_Archivo);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public static void releaseObject(object obj)
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
                obj = null;
                GC.Collect();
            }
        }
        public static string SelectFromWhere(string executequery, bool eslike)
        {
            string vDATO = "";
            string queryString = executequery.ToString();

            using (OracleConnection connection =
                   new OracleConnection(ConfigurationManager.ConnectionStrings[Parameters.ambiente].ConnectionString))
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
        public static void InsUpdDel_Oracle(string executequery)
        {
            string queryString = executequery.ToString();
            string queryCommit = "COMMIT";

            using (OracleConnection connection =
                   new OracleConnection(ConfigurationManager.ConnectionStrings[Parameters.ambiente].ConnectionString))
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
        public static string Function_Procedure_Oracle(int tipofunpro /* tipofunpro: 1 para function, 2 para procedure  */, string executequery, string nomprmt1, int prmt1, string nomprmt2, int prmt2, string nomprmt3, int prmt3)
        {

            string vDATO = "";
            string queryString = executequery.ToString();

            using (OracleConnection connection =
                   new OracleConnection(ConfigurationManager.ConnectionStrings[Parameters.ambiente].ConnectionString))
            {
                try
                {
                    connection.Open();

                    OracleCommand command = connection.CreateCommand();
                    command.CommandText = queryString;
                    command.CommandType = CommandType.StoredProcedure;
                    if (prmt1 >= 0)
                    {
                        command.Parameters.Add(new OracleParameter(nomprmt1, System.Data.OracleClient.OracleType.Int32)).Value = prmt1;
                        command.Parameters[nomprmt1].Direction = ParameterDirection.Input;
                    }
                    if (prmt2 >= 0)
                    {
                        command.Parameters.Add(new OracleParameter(nomprmt2, System.Data.OracleClient.OracleType.Int32)).Value = prmt2;
                        command.Parameters[nomprmt2].Direction = ParameterDirection.Input;
                    }
                    if (prmt3 >= 0)
                    {
                        command.Parameters.Add(new OracleParameter(nomprmt3, System.Data.OracleClient.OracleType.Int32)).Value = prmt3;
                        command.Parameters[nomprmt3].Direction = ParameterDirection.Input;
                    }

                    if (tipofunpro == 1)
                    {
                        command.Parameters.Add("retorno", System.Data.OracleClient.OracleType.Int32);
                        command.Parameters["retorno"].Direction = ParameterDirection.ReturnValue;
                    }

                    command.ExecuteNonQuery();

                    if (tipofunpro == 1)
                        vDATO = Convert.ToString(command.Parameters["retorno"].Value);

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
        public static void argInvalid(bool pause)
        {
            Console.WriteLine("Argumentos Invalidos");
            Console.WriteLine(" Ayuda:");
            Console.WriteLine("  -pauseend: Pausa el proceso al llegar al final. Por defecto no pausa");
            Console.WriteLine("  -killall: Termina todos los procesos en ejecucion de nombre LectorExcelConciliacion.exe");
            Console.WriteLine("  -file <ruta>: Procesa el archivo <ruta>");
            Console.WriteLine("  Sin parametros: Procesa todos los archivos en la carpeta INPUT");
            if (pause)
            {
                Console.WriteLine("Presione cualquier tecla para continuar...");
                Console.ReadKey();
            }
            Environment.Exit(0);
        }
        public static string obtRuta(int tblcodarg, string name)
        {
            string obtRuta = SelectFromWhere("SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = 1 AND tblcodarg IN (" + tblcodarg + ")", false);
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
            return obtRuta + "//";
        }
    }
}
