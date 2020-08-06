using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class Bracket:Models.SYS_Bracket
    {
        public string ChineseName { get; set; }
        public string WhetherStarta { get; set; }
        public bool WhetherStartb { get; set; }
        private string time;
        public string timer
        {
            set
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(value);
                    time = dt.ToString("yyyy-MM-dd");
                }
                catch (Exception)
                {
                    time = value;
                }
            }
            get
            {
                return time;
            }
        }
    }
}