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
    
    public partial class SYS_Commercel
    {
        public int CommerceID { get; set; }
        public int EtrustID { get; set; }
        public Nullable<int> CommerceExaminePID { get; set; }
        public Nullable<int> CommerceQiExaminePID { get; set; }
        public Nullable<int> FinancingExaninePID { get; set; }
        public Nullable<int> FinancingQiExaninePID { get; set; }
        public Nullable<System.DateTime> CommerceQiExamineTime { get; set; }
        public Nullable<System.DateTime> FinancingQiExanineTime { get; set; }
        public Nullable<System.DateTime> CommerceExamineTime { get; set; }
        public Nullable<System.DateTime> FinancingExanineTime { get; set; }
    }
}
