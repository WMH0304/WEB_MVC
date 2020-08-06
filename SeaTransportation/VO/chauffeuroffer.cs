using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class chauffeuroffer:Models.SYS_Offer
    {
        public string ChineseName { get; set; }
        private string time;
        public string OfferDatetime
        {
            set
            {
                try
                {
                    DateTime tt = Convert.ToDateTime(value);
                    time = tt.ToString("yyyy-MM-dd");
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
        private string timer;
        public string TakeEffectDatetime
        {
            set
            {
                try
                {
                    DateTime tt = Convert.ToDateTime(value);
                    timer = tt.ToString("yyyy-MM-dd");
                }
                catch (Exception)
                {
                    timer = value;
                }
            }
            get
            {
                return timer;
            }
        }
        private string timee;
        public string LoseEfficacyDatetime
        {
            set
            {
                try
                {
                    DateTime tt = Convert.ToDateTime(value);
                    timee = tt.ToString("yyyy-MM-dd");
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