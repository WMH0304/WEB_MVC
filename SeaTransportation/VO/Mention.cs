using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class Mention:Models.SYS_Mention
    {
        public string GatedotName { get; set; }
        public string WhetherStarta { get; set; }
        public bool WhetherStartb { get; set; }
    }
}