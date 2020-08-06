using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class Ship:Models.SYS_Ship
    {
        public string ClientType { get; set; }
        public int ClientID { get; set; }
        public string ChineseNamel { get; set; }
        public bool WhetherStartl { get; set; }
        public string WhetherStarta { get; set; }
    }
}