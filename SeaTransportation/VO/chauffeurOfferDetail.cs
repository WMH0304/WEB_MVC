using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class chauffeurOfferDetail:Models.SYS_OfferDetail
    {
        public string OfferType { get; set; }
        public string ExpenseName { get; set; }
        public string HaulWayDescription { get; set; }
        public string StaffName { get; set; }
    }
}