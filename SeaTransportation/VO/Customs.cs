using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class Customs:Models.SYS_Customs
    {
        public string GatedotName { get; set; }
        public string WhetherStarta { get; set; }
        public bool WhetherStartb { get; set; }
        private string time;
        private string timee;
        public string timer
        {
            set
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(value);
                    time = dt.ToString("HH:mm");
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
        public string timerer
        {
            set
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(value);
                    timee = dt.ToString("HH:mm");
                }
                catch (Exception)
                {
                    timee = value;
                }
            }
            get
            {
                return timee;
            }
        }
    }
}