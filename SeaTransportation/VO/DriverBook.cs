using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class DriverBook:Models.SYS_DriverBook
    {
        public string ChauffeurName { get; set; }
        
        private string time;
        private string timeerer;
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
        public string ExpirationDatetimer
        {
            set
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(value);
                    timeerer = dt.ToString("yyyy-MM-dd");
                }
                catch (Exception)
                {
                    timeerer = value;
                }
            }
            get
            {
                return timeerer;
            }
        }
    }
}