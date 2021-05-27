using System;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace LectorExcelConciliacion
{
    public class LogWriter
    {
        string Rutawork;
        string log;
        string id;
        int linea = 0;
        int cont = 0;
        int limite = 30;
        public LogWriter(string rutawork)
        {
            Rutawork = rutawork;
            Process currentProcess = Process.GetCurrentProcess();
            id = currentProcess.Id.ToString();
            log = "";
        }

        public void LogWrite()
        {
            string exmsg = "";
            while (log != "" && cont < limite)
            {
                cont++;
                try
                {
                    using (StreamWriter w = File.AppendText(Rutawork + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                    {
                        w.Write(log);
                        log = "";
                        w.Close();
                    }
                }
                catch (Exception ex)
                {
                    System.Threading.Thread.Sleep(254);
                    exmsg = ex.Message;
                    //Console.WriteLine(ex.Message);
                }
            }
            if (log != "")
            {
                Console.WriteLine(exmsg);
            }
            cont = 0;
        }

        public void addLog(string logMessage, Boolean error)
        {
            try
            {
                log += id + "|";
                log += linea + "|";
                log += DateTime.Now.ToString("HH:mm:ss") + "|";
                log += Regex.Replace(logMessage, @"\r\n?|\n", " ") + "|";
                if (error)
                {
                    log += "ERROR|";
                }
                log += "\r\n";
                linea++;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
