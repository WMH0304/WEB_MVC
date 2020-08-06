using System;

namespace SeaTransportation.Areas.PPAP.Controllers
{
    internal class Staffdttcc
    {
        public string ChineseName { get;  set; }
        public string CultureLevel { get;  set; }
        public string Duty { get; set; }

        private string ss;
  

        public string JoinJob
        {
          //  get;set;

            get { return ss; }
            set
            {
                DateTime ww = Convert.ToDateTime(value);
                ss = ww.ToString("yyyy-MM-hh");
            }
        }
                
            
        public string Sex { get;  set; }
        public string StaffName { get; set; }
        public string StaffNumber { get; set; }
        public string State { get; set; }
        public bool WhetherStart { get;  set; }
        public short MessageID { get;  set; }
        public int DepartmentID { get;  set; }
        public int StaffID { get;  set; }
    }
}