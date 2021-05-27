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

        public LogWriter(string rutawork)
        {
            Rutawork = rutawork;
            Process currentProcess = Process.GetCurrentProcess();
            id = currentProcess.Id.ToString();
            log = "";
        }

        public void LogWrite()
        {
            try
            {
                using (StreamWriter w = File.AppendText(Rutawork + "\\" + DateTime.Now.ToString("yyyy-MM-dd") +".txt"))
                {
                    w.Write(log);
                    log = "";
                    w.Close();
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
            }
        }

        public void addLog(string logMessage, Boolean error)
        {
            try
            {
                log += id + "|";
                log += DateTime.Now.ToString("HH:mm:ss") + "|";
                log += Regex.Replace(logMessage, @"\r\n?|\n", " ") + "|";
                if (error)
                {
                    log += "ERROR|";
                }
                log += "\r\n";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
