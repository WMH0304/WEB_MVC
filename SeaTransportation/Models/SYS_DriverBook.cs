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
    
    public partial class SYS_DriverBook
    {
        public int DriverBookID { get; set; }
        public int ChauffeurID { get; set; }
        public string CustomsNumber { get; set; }
        public string BusinessLicenseNumber { get; set; }
        public string IssuingAuthority { get; set; }
        public string Remark { get; set; }
        public string AccurateLoadNumber { get; set; }
        public Nullable<System.DateTime> TermValidity { get; set; }
        public Nullable<System.DateTime> ExpirationDate { get; set; }
        public string DriverBook { get; set; }
    }
}