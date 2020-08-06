using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.VO
{
    public class ClientVo
    {
        public string ClientCode { get; set; }//客户代码
        public string ClientAbbreviation { get; set; }//客户简称
        public string ClientType { get; set; }//客户类型
        public string ClientTypee { get; set; }
        public string ChineseName { get; set; }//中文名称
        public string ClientRank { get; set; }//客户级别
        public string CustomsCode { get; set; }//海关编码
        public string GatedotName { get; set; }//所属区域
        public string ClientSource { get; set; }//客户来源
        public string ClientPhone { get; set; }//客户电话
        public bool? WhetherStart { get; set; }//是否启用
        public string WhetherStartt { get; set;  }
        public int? GatedotID { get; set; }//门点ID
        public int? MessageID { get; set; }//组织结构（本公司ID）
        public string StaffName { get; set; }//员工名称
        public int? ClientID { get; set; }//客户信息ID
        public string ClientFax { get; set; }//客户传真
        public string Email { get; set; }
        public string PostCode { get; set; }//邮编
        public string  OfficeHours { get; set; }//上班时间
        public string ClosingTime { get; set; }//下班时间
        public string Site { get; set; }//地址
        public string Website { get; set; }//网站
        public string OpenAccount { get; set; }//开户行
        public string OpenAccountCode { get; set; }//开户行账户
        public string Describe { get; set; }//描述
        public int? ClientScontactsID { get; set; }//客户联系人ID
        public string contactsPhone { get; set; }//联系人电话
        public string contactsName { get; set; }//联系人名称
        public int? ClientSiteID { get; set; }//客户地址ID
        public string ContactsName { get; set; }//联系人名称
        public string ContactsPhone { get; set; }//联系人电话
        public string ClientSite { get; set; }//客户地址
        public string FactoryName { get; set; }//工厂名称
        public string FactoryCode { get; set; }//工厂代码
        public string Remark { get; set; }//备注
        public int? OfferID { get; set; }//报价ID
        public string TakeEffectDate { get; set; }//生效日期
        public string LoseEfficacyDate { get; set; }//失效日期
        public string WhetherShuii { get; set; }//是否含税
        public bool? WhetherShui { get; set; }
        public int? OfferDetailID { get; set; }//报价明细ID
        public string HaulWayDescription { get; set; }//运输路线
        public string ExpenseName { get; set; }//费用项目名称
        public string CabinetType { get; set; }//箱型
        public string EtryClasses { get; set; }//报关方式
        public decimal? Money { get; set; }//金额
        public string Currency { get; set; }//币种
        public int? HaulWayID { get; set; }//运输路线ID
        public string Abbreviation { get; set; }//地点
        public string Abbreviationn { get; set; }
        public string BoxQuantity { get; set; }//箱量\
        public DateTime? OfferDate { get; set; }















        //转换上班时间
        public string OfficeHours1
        {
            get
            {
                try
                {
                    OfficeHours = Convert.ToDateTime(OfficeHours).ToString("HH:mm");
                    return OfficeHours;
                }
                catch (Exception)
                {
                    return OfficeHours;
                }
            }
            set
            {
                OfficeHours = value;
            }
        }

        //转换下班时间
        public string ClosingTime1
        {
            get
            {
                try
                {
                    ClosingTime = Convert.ToDateTime(ClosingTime).ToString("HH:mm");
                    return ClosingTime;
                }
                catch (Exception)
                {
                    return ClosingTime;
                }
            }
            set
            {
                ClosingTime = value;
            }
        }

        private string offerDate { get; set; }
        //转换报价时间
        public string OfferDate1
        {
            get
            {
                try
                {
                    offerDate = Convert.ToDateTime(offerDate).ToString("yyyy-MM-hh");
                    return offerDate;
                }
                catch (Exception)
                {
                    return offerDate;
                }
            }
            set
            {
                offerDate = value;
            }
        }
        //转换生效时间
        public string TakeEffectDate1
        {
            get
            {
                try
                {
                    TakeEffectDate = Convert.ToDateTime(TakeEffectDate).ToString("yyyy-MM-hh");
                    return TakeEffectDate;
                }
                catch (Exception)
                {
                    return TakeEffectDate;
                }
            }
            set
            {
                TakeEffectDate = value;
            }
        }
        //转换失效时间
        public string LoseEfficacyDate1
        {
            get
            {
                try
                {
                    LoseEfficacyDate = Convert.ToDateTime(LoseEfficacyDate).ToString("yyyy-MM-hh");
                    return LoseEfficacyDate;
                }
                catch (Exception)
                {
                    return LoseEfficacyDate;
                }
            }
            set
            {
                LoseEfficacyDate = value;
            }
        }
    }
}