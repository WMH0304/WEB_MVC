using System;

namespace SeaTransportation.Areas.PPAP.Controllers
{
    internal class UserManagement
    {
        public string AccountName { get; set; }
        public string ChineseName { get; set; }
        public string DepartmentName { get; set; }
        public string StaffName { get;  set; }
        public string Describe { get; set; }
        private string ss;
       

        public string Time
        {set;get;
            //get { return ss; }
            //set
            //{
            //    DateTime ww = Convert.ToDateTime(value);
            //    ss = ww.ToString("yyyy-MM-hh");
            //}
        }
        public string password { get; set; }
        public int UserID { get; set; }
        public bool WhetherStart { get; set; }
        public short MessageID { get; set; }
        public int StaffID { get;  set; }
        public int DepartmentID { get;  set; }
 
        public string FinallyRegisterTime { get;  set; }
    }
}