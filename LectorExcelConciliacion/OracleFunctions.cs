using System;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace LectorExcelConciliacion
{
    class OracleFunctions
    {
        string Conexion;
        LogWriter LogWriter;
        public OracleFunctions(string conexion, LogWriter logWriter)
        {
            Conexion = conexion;
            LogWriter = logWriter;
        }

        public string SelectFromWhere(string executequery, bool eslike)
        {
            string vDATO = "";
            string queryString = executequery.ToString();

            try
            {
                using (OracleConnection connection =
                   new OracleConnection(Conexion))
                {
                    OracleCommand command = connection.CreateCommand();
                    command.CommandText = queryString;
                        connection.Open();
                        OracleDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            vDATO = reader[0].ToString();
                            if (eslike) break;
                        }
                        reader.Close();
                }
            }
            catch (Exception ex)
            {
                vDATO = null;
                Console.WriteLine(ex.Message);
                using (StreamWriter w = File.AppendText("\\\\wlspacifico\\e\\ConciliacionBancaria$\\WORK\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                {
                    Process currentProcess = Process.GetCurrentProcess();
                    string log = "";
                    log += currentProcess.Id.ToString() + "|";
                    log += DateTime.Now.ToString("HH:mm:ss") + "|";
                    log += "SelectFromWhere|";
                    log += "ERROR|";
                    log += "\r\n";
                    w.Write(log);
                    log = "";
                    w.Close();
                }
                if (LogWriter != null)
                    LogWriter.addLog(ex.Message, true);
            }
            return vDATO;
        }
        public void InsUpdDel_Oracle(string executequery)
        {
            string queryString = executequery.ToString();
            string queryCommit = "COMMIT";

            try
            {
                using (OracleConnection connection =
                   new OracleConnection(Conexion))
                {
                    OracleCommand command = connection.CreateCommand();
                        connection.Open();
                        command.CommandType = CommandType.Text;
                        command.CommandText = queryString;
                        command.ExecuteNonQuery();
                        command.CommandText = queryCommit;
                        command.ExecuteNonQuery();

                        connection.Close();
                        command.Dispose();
                    }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                using (StreamWriter w = File.AppendText("\\\\wlspacifico\\e\\ConciliacionBancaria$\\WORK\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                {
                    Process currentProcess = Process.GetCurrentProcess();
                    string log = "";
                    log += currentProcess.Id.ToString() + "|";
                    log += DateTime.Now.ToString("HH:mm:ss") + "|";
                    log += "InsUpdDel_Oracle|" + executequery + "";
                    log += "ERROR|";
                    log += "\r\n";
                    w.Write(log);
                    log = "";
                    w.Close();
                }
                if (LogWriter != null)
                    LogWriter.addLog(ex.Message, true);
            }
        }
        public string Function_Procedure_Oracle(int tipofunpro /* tipofunpro: 1 para function, 2 para procedure  */, string executequery, string nomprmt1, int prmt1, string nomprmt2, int prmt2, string nomprmt3, int prmt3)
        {
            string vDATO = "";
            try
            {
                using (OracleConnection connection =
                   new OracleConnection(Conexion))
                {
                        OracleCommand command = new OracleCommand(executequery, connection) { CommandType = CommandType.StoredProcedure };

                        if (tipofunpro == 1)
                        {
                            var returnVal = new OracleParameter("Return_Value", OracleDbType.Int32) { Direction = ParameterDirection.ReturnValue };
                            command.Parameters.Add(returnVal);
                        }

                        if (prmt1 >= 0)
                        {
                            var prm1 = new OracleParameter(nomprmt1, OracleDbType.Int32) { Direction = ParameterDirection.Input, Value = prmt1 };
                            command.Parameters.Add(prm1);
                        }
                        if (prmt2 >= 0)
                        {
                            var prm2 = new OracleParameter(nomprmt2, OracleDbType.Int32) { Direction = ParameterDirection.Input, Value = prmt2 };
                            command.Parameters.Add(prm2);
                        }
                        if (prmt3 >= 0)
                        {
                            var prm3 = new OracleParameter(nomprmt3, OracleDbType.Int32) { Direction = ParameterDirection.Input, Value = prmt3 };
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
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                using (StreamWriter w = File.AppendText("\\\\wlspacifico\\e\\ConciliacionBancaria$\\WORK\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                {
                    Process currentProcess = Process.GetCurrentProcess();
                    string log = "";
                    log += currentProcess.Id.ToString() + "|";
                    log += DateTime.Now.ToString("HH:mm:ss") + "|";
                    log += "Function_Procedure_Oracle:" + executequery + "|";
                    log += "ERROR|";
                    log += "\r\n";
                    w.Write(log);
                    log = "";
                    w.Close();
                }
                if (LogWriter != null)
                    LogWriter.addLog(ex.Message, true);
            }
            return vDATO;
        }
        public string ObtRuta(int tblcodarg, string name)
        {
            string obtRuta = SelectFromWhere("SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLESTADO = 1 AND TBLCODARG IN (" + tblcodarg + ")", false);
            if (obtRuta == null)
            {
                Console.WriteLine("Error obteniendo la ruta " + name);
                if (LogWriter != null)
                {
                    LogWriter.addLog("Error obteniendo la ruta " + name, true);
                    LogWriter.addLog("Fin", false);
                    LogWriter.LogWrite();
                }
                Environment.Exit(0);
            }
            else if (!Directory.Exists(obtRuta))
            {
                Console.WriteLine("Error: Ruta " + name + " no encontrada\nRuta: " + obtRuta);
                if (LogWriter != null)
                {
                    LogWriter.addLog("Ruta " + name + " no encontrada. " + obtRuta, true);
                    LogWriter.addLog("Fin", false);
                    LogWriter.LogWrite();
                }
                Environment.Exit(0);
            }
            return obtRuta + Path.DirectorySeparatorChar;
        }

    }
}
