//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成。
//
//     手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace SeaTransportation.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class SYS_Port
    {
        public int PortID { get; set; }
        public string PortCode { get; set; }
        public string ChineseName { get; set; }
        public string EnglishName { get; set; }
        public string CountryCode { get; set; }
        public string Area { get; set; }
        public Nullable<int> Berthage { get; set; }
        public string Phone { get; set; }
        public string Remark { get; set; }
        public bool WhetherStart { get; set; }
    }
}
