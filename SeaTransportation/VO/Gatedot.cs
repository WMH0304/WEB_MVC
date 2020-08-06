using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class Gatedot:Models.SYS_Gatedot
    {
        public string Province { get; set; }
        public string CityName { get; set; }
        public string WhetherSuperiorGatedota { get; set; }
        public bool WhetherSuperiorGatedotb { get; set; }
        public string WhetherValida { get; set; }
        public bool WhetherValidb { get; set; }
    }
}