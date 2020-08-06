using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class VehicleInformation:Models.SYS_VehicleInformation
    {
        public string ChauffeurName { get; set; }
        public string ChineseName { get; set; }
        public string BracketCode { get; set; }
    }
}