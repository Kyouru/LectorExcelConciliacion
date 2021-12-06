using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LectorExcelConciliacion
{
    public class ConnString
    {
        public string desarrollo = "Data Source=";
        public string qa = "Data Source=";
        public string produccion = "Data Source=";
        
        public string desarrollo2 = "Data Source=;User Id=;Password=;";
        public string desarrollo3 = "Data Source=;User Id=;Password=;";
        public string qa2 = "Data Source=;User Id=;Password=;";
        public string produccion2 = "Data Source=;User Id=;Password=;";
        
        public string GetString(string _name)
        {
            return (string)typeof(ConnString).GetField(_name).GetValue(this);
        }
    }

}
