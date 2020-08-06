using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class Chauffeur:Models.SYS_Chauffeur
    {
        public string ChineseName { get; set; }
        public string ClientType { get; set; }
        public string WhetherStarta { get; set; }
        public bool WhetherStartb { get; set; }
    }
}