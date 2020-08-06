using SeaTransportation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.VO
{
    public class ExamineVo 
    {
      
        public string ClientAbbreviation { get; set; }//客户简称
        public string ClientCodee { get; set; }//用于多条件查询的客户代码
        public string Finance { get; set; }//财务审核否的状态
        public int? ClientID { get; set; }//客户信息ID
        public string User { get; set; }//商务弃审人
        public string User1 { get; set; }//商务审核人
        public string ExpenseName { get; set; }//费用项目名称
        public int ExpenseID { get; set; }//费用项目ID
        public string ChargeType { get; set; }//费用类型
        public string SettleAccounts { get; set; }//收付款名称
        public decimal? UnitPrice { get; set; }//单价
        public string BoxQuantity { get; set; }//箱量
        public decimal? Money { get; set; }//金额
        public string Currency { get; set; }//币种
        public int ChargeID { get; set; }//收费ID
        public string ReckoningUnit { get; set; }//结算单位
        public string ArriveFactoryTime { get; set; }//到达工厂时间
        public  string LeftFactoryTime { get; set; }//离开工厂时间
        public Nullable<decimal> YingShouZ { get; set; }//应收总额
        public Nullable<decimal> YingFuZ { get; set; }//应付总额
        public Nullable<decimal> LiRun { get; set; }//利润
        public int? SignCloseAccountID { get; set; }//对账结算ID
        public string SettleWay { get; set; }//结算方式
        public string Parities { get; set; }// 汇率
        public string ChineseName { get; set; }//船中文名
        public int ChauffeurID { get; set; }//司机资料ID
        public string PlateNumbers { get; set; }// 车牌号
        public int BracketID { get; set; }//托架资料ID
        public string YiJieMoney { get; set; }//已结金额
        public string HuaiWeightTime { get; set; }//还重柜时间
        public int? SignID { get; set; }//对账ID
        public string Bgs { get; set; }//本公司的结算单位
        public string ReckoningUnits { get; set; }//客户公司的结算单位
        public string Kh { get; set; }//客户公司的结算单位
        public string TissueCode { get; set; }//本公司的组织代码（结算单位）
        public string SettleNumber { get; set; }//结算单号
        public string PayWay { get; set; }//付款方式
        public string YuSettleDate { get; set; }//预结款日期
        public string FinishType { get; set; }//结算状态(已核销，未核销)
        public string YiHeXiao { get; set; }//已核销金额
        public string YiChHeX { get; set; }//异常核销金额
        public int EtrustID { get; set; } //委托单ID
        public string EtrustNmber { get; set; }//委托单号
        public string CarriageNumber { get; set; }//运输单号
        public string EtrustType { get; set; }//委托状态
        public string CheXing { get; set; }//车型
        public string AuditType { get; set; }//审核状态
        public string Cupboard { get; set; }//柜号
        public string Seal { get; set; } //封条号
        public string IndentNumber { get; set; }//订单号
        public int? TiGuiDiDianID { get; set; }//提柜地点ID
        public int? HuaiGuiDiDianID { get; set; }//还柜地点ID
        public string ClientCode { get; set; }//客户代码
        public string WorkCategory { get; set; }//作业类别
        public string HangCi { get; set; }//航次
        public int? ShipID { get; set; }//船舶资料ID
        public int? PortID { get; set; }//港口资料ID
        public int? GoalHarborID { get; set; }//目的港
        public int? UndertakeID { get; set; }//承运公司
        public int? VehicleInformationID { get; set; }//车辆信息ID
        public string BookingSpace { get; set; }//订舱号
        public string CabinetType { get; set; }//箱型
        public string ClientType { get; set; }//客户类型

        /// <summary>
        /// 商务审核表
        /// </summary>
        public int CommerceID { get; set; }
        public Nullable<int> CommerceExaminePID { get; set; }
        public Nullable<int> CommerceQiExaminePID { get; set; }
        public Nullable<int> FinancingExaninePID { get; set; }
        public Nullable<int> FinancingQiExaninePID { get; set; }
        public string  CommerceQiExamineTime { get; set; }
        public Nullable<System.DateTime> FinancingQiExanineTime { get; set; }
        private string CommerceExamineTime { get; set; }
        public Nullable<System.DateTime> FinancingExanineTime { get; set; }


        public string SignType { get; set; }//对账状态ID
        public string SignExplain { get; set; }//对账说明
        public string SignDate { get; set; }//对账日期
        public int? PayWayID { get; set; }//付款方式ID
        public int? SettleWayID { get; set; }//结算方式ID
        public decimal? SettleMoney { get; set; }//结算金额
        public decimal? SettleMoney1 { get; set; }//结算金额
        public decimal? BCYShou { get; set; }//本次应收
        public decimal? BCYFu { get; set; }//本次应付
        public int? CloseAccountID { get; set; }//结算ID
        public int? CalculateID { get; set; }//计费ID
        public string CalculateState { get; set; }//计费状态
        public string CalculateNumber { get; set; }//收付款单号
        public decimal? SFMoney { get; set; }//收费金额
        public decimal? AlreadyCancelMoney { get; set; }//已核销金额
        public decimal? StayCancelMoney { get; set; }//待核销金额
        public string SFDate { get; set; }//收费日期
        public string BankName { get; set; }//银行名称
        public string BankAccount { get; set; }//银行账号
        public string ChequeNumber { get; set; }//支票号
        public int? MessageID { get; set; }//本公司信息ID

















        //转换商务审核时间
        public string ShangWuShenHeTime
        {
            get
            {
                try
                {
                    CommerceExamineTime = Convert.ToDateTime(CommerceExamineTime).ToString("yyyy-MM-dd HH:mm");
                    return CommerceExamineTime;
                }
                catch (Exception)
                {
                    return CommerceExamineTime;
                }
            }
            set
            {
                CommerceExamineTime = value;
            }
        }

        //转换商务弃审时间
        public string ShangWuQiShenTime
        {
            get
            {
                try
                {
                    CommerceQiExamineTime = Convert.ToDateTime(CommerceQiExamineTime).ToString("yyyy-MM-dd HH:mm");
                    return CommerceQiExamineTime;
                }
                catch (Exception)
                {
                    return CommerceQiExamineTime;
                }
            }
            set
            {
                CommerceQiExamineTime = value;
            }
        }

        public DateTime? ArriveFactoryTimes { get; set; }//到达工厂时间
        //转换到达工厂时间
        public string ArriveFactoryTime1
        {
            get
            {
                try
                {
                    ArriveFactoryTime = Convert.ToDateTime(ArriveFactoryTime).ToString("yyyy-MM-dd HH:mm:ss");
                    return ArriveFactoryTime;
                }
                catch (Exception)
                {
                    return ArriveFactoryTime;
                }
            }
            set
            {
                ArriveFactoryTime = value;
            }
        }
        public DateTime? LeftFactoryTimes { get; set; }
        //转换离开工厂时间
        public string LeftFactoryTime1
        {
            get
            {
                try
                {
                    LeftFactoryTime = Convert.ToDateTime(LeftFactoryTime).ToString("yyyy-MM-dd HH:mm");
                    return LeftFactoryTime;
                }
                catch (Exception)
                {
                    return LeftFactoryTime;
                }
            }
            set
            {
                LeftFactoryTime = value;
            }
        }

        public DateTime? HuaiWeightTimee { get; set; }
        //转换还重柜时间
        public string HuaiWeightTime1
        {
            get
            {
                try
                {
                    HuaiWeightTime = Convert.ToDateTime(HuaiWeightTime).ToString("yyyy-MM-dd HH:mm");
                    return HuaiWeightTime;
                }
                catch (Exception)
                {
                    return HuaiWeightTime;
                }
            }
            set
            {
                HuaiWeightTime = value;
            }
        }

        //转换预结款日期
        public string YuSettleDate1
        {
            get
            {
                try
                {
                    YuSettleDate = Convert.ToDateTime(YuSettleDate).ToString("yyyy-MM-dd");
                    return YuSettleDate;
                }
                catch (Exception)
                {
                    return YuSettleDate;
                }
            }
            set
            {
                YuSettleDate = value;
            }
        }

        //转换对账日期
        public string SignDate1
        {
            get
            {
                try
                {
                    SignDate = Convert.ToDateTime(SignDate).ToString("yyyy-MM-dd");
                    return SignDate;
                }
                catch (Exception)
                {
                    return SignDate;
                }
            }
            set
            {
                SignDate = value;
            }
        }

        /// <summary>
        /// 转换收付日期
        /// </summary>
        public string SFDate1
        {
            get
            {
                try
                {
                    SFDate = Convert.ToDateTime(SFDate).ToString("yyyy-MM-dd");
                    return SFDate;
                }
                catch (Exception)
                {
                    return SFDate;
                }
            }
            set
            {
                SFDate = value;
            }
        }


    }
}