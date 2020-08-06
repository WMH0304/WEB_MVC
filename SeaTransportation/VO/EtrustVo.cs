using SeaTransportation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Vo
{
    public class EtrustVo: SYS_Etrust
    {
        public string Name { get; set; }
        public decimal? C { get; set; }
        public decimal? J { get; set; }
        public decimal? S { get; set; }
        public decimal? M { get; set; }
        public  string   PlateNumbers   { get; set; }
         public  string   BracketTag   { get; set; }
         public  int?    ChauffeurID { get; set; }
         public  string    ClientAbbreviation    { get; set; }
         public  string    contactsName         { get; set; }
         public  string    contactsPhone        { get; set; }
         public  string    ContactsName       { get; set; }
         public  string    ContactsPhone       { get; set; }
         public  string    chauffeurName       { get; set; }
         public  string    ChauffeurNumber     { get; set; }
         public  string    BracketCode          { get; set; }
         public  string    VehicleCode          { get; set; }
         public  int    ChargeID             { get; set; }
         public  string    ExpenseName         { get; set; }
         public  string    SettleAccounts        { get; set; }
         public  string    CloseAccountUnit     { get; set; }
         public  string    BoxQuantity          { get; set; }
         public  Decimal    Money                { get; set; }
         public string Currency    { get; set; }
        private new string ArriveTime { get; set; }
        public string arriveTime 
        {
            get
            {
                try
                {
                    ArriveTime = Convert.ToDateTime(ArriveTime).ToString("yyyy-MM-dd HH:mm:ss");
                    return ArriveTime;
                }
                catch (Exception)
                {
                    return ArriveTime;
                }
            }
            set
            {
                ArriveTime = value;
            }
        }
        public string arriveTimes
        {
            get
            {
                try
                {
                    ArriveTime = Convert.ToDateTime(ArriveTime).ToString("yyyy-MM-dd HH:mm");
                    return ArriveTime;
                }
                catch (Exception)
                {
                    return ArriveTime;
                }
            }
            set
            {
                ArriveTime = value;
            }
        }
        public string Undertake { get; set; }
        public string ChauffeurName { get; set; }
        public Decimal? UnitPrice { get; set; }
        public string ReckoningUnit { get; set; }
        public string ReckoningUnits { get; set; }
        public string Reckoning { get; set; }
        public int ClientID { get; set; }
        public string ChineseName { get; set; }
        private new string PlanHandTime { get; set; }
        public  string planHandTime
        {
            get
            {
                try
                {
                    PlanHandTime = Convert.ToDateTime(PlanHandTime).ToString("yyyy-MM-dd HH:mm:ss");
                    return PlanHandTime;
                }
                catch (Exception)
                {
                    return PlanHandTime;
                }
            }
            set
            {
                PlanHandTime = value;
            }
        }
        public string T { get; set; }
        public string H { get; set; }
        public string FactoryName { get; set; }
        public string FactoryCode { get; set; }
        public string ClientSite { get; set; }
        public string Chuangs { get; set; }
        private new string HuiChangTime { get; set; }
        public string huiChangTime
        {
            get
            {
                try
                {
                    HuiChangTime = Convert.ToDateTime(HuiChangTime).ToString("yyyy-MM-dd HH:mm:ss");
                    return HuiChangTime;
                }
                catch (Exception)
                {
                    return HuiChangTime;
                }
            }
            set
            {
                HuiChangTime = value;
            }
        }
        public string GatedotName { get; set; }
        public string StaffName { get; set; }
        public string Staff{ get; set; }
        private new string KaiCangTime { get; set; }
        public string kaiCangTime
        {
            get
            {
                try
                {
                    KaiCangTime = Convert.ToDateTime(KaiCangTime).ToString("yyyy-MM-dd HH:mm:ss");
                    return KaiCangTime;
                }
                catch (Exception)
                {
                    return KaiCangTime;
                }
            }
            set
            {
                KaiCangTime = value;
            }
        }
        public string TG { get; set; }
        public string HG { get; set; }
        public decimal? money { get; set; }
        public int? clientID { get; set; }
        public int? offerDetailID { get; set; }
        public string OfferType { get; set; }
        public string YSLX { get; set; }
        public string Shij { get; set; }
        public string ICNumber { get; set; }
        public string PhoneOne { get; set; }
        public string BracketType { get; set; }
        private new string PaiCheDate { get; set; }
        public string paiCheDate
        {
            get
            {
                try
                {
                    PaiCheDate = Convert.ToDateTime(PaiCheDate).ToString("yyyy-MM-dd");
                    return PaiCheDate;
                }
                catch (Exception)
                {
                    return PaiCheDate;
                }
            }
            set
            {
                PaiCheDate = value;
            }
        }
        public string MessageName { get; set; }
    }
}               