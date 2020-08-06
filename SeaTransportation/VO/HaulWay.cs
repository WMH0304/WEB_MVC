using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class HaulWay:Models.SYS_HaulWay
    {
        public string MentionArea { get; set; }
        public string AlsoTankArea { get; set; }
        public string GatedotName { get; set; }
        public string CustomsName { get; set; }
        public string WhetherStarta { get; set; }
        public bool WhetherStartb { get; set; }
        private string time;
        private string timeer;
        private string timee;
        public string OutwardVoyageTimetimer
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
        public string ReturnTripTimetimer
        {
            set
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(value);
                    timeer = dt.ToString("yyyy-MM-dd");
                }
                catch (Exception)
                {
                    timeer = value;
                }
            }
            get
            {
                return timeer;
            }
        }
        public string TotalTimetimer
        {
            set
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(value);
                    timee = dt.ToString("yyyy-MM-dd");
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