using Hotel.VO;
using SeaTransportation.Models;
using SeaTransportation.VO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Areas.Commerce.Controllers
{
    public class ExamineController : Controller
    {
        Models.SeaTransportationEntities myModels = new Models.SeaTransportationEntities();
        // GET: Commerce/Examine


        #region 视图


        /// <summary>
        ///  商务审核
        /// </summary>
        /// <returns></returns>
        public ActionResult ExamineIndex()
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }

        /// <summary>
        /// 财务结算
        /// </summary>
        /// <returns></returns>
        public ActionResult FinanceIndex()
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }

        /// <summary>
        /// 财务结算左边表格新增结算单
        /// </summary>
        /// <returns></returns>
        public ActionResult InsetIndex(int? SignCloseAccountID)
        {
            Session["SignCloseAccountID"] = SignCloseAccountID;//对账结算ID
            return View();
        }

        /// <summary>
        /// 财务结算下边表格新增结算明细单
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertDetail(int? SignID)
        {
            Session["SignID"] = SignID;//对账ID
            return View();
        }

        /// <summary>
        /// 实收实付页面(用于核销)
        /// </summary>
        /// <returns></returns>
        public ActionResult ShiShouShiFu()
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }





        #endregion

        #region 查询表格数据

        /// <summary>
        /// 查询审核上方表格的数据
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectExamine(BsgridPage bsgridPage,int? CommerceID,string CarriageNumber,string Cupboard,string Seal,string IndentNumber,string ArriveFactoryTime,string LeftFactoryTime,string ClientAbbreviation,string EtrustNmber,int? TiGuiDiDianID,int? HuaiGuiDiDianID)
        {
            var listExamine = (from tbExamine in myModels.SYS_Commercel
                              join tbEtrust in myModels.SYS_Etrust on tbExamine.EtrustID equals tbEtrust.EtrustID
                              join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              join tbUser in myModels.SYS_User on tbExamine.CommerceQiExaminePID equals tbUser.UserID into From
                              from Add in From.DefaultIfEmpty()//左链接(同样一个表有两个不同的外键ID，而这两个外键ID是代表另一个(相同)表的主键ID，这时候就用左链接进行相连)
                              join tbUser in myModels.SYS_User on tbExamine.CommerceExaminePID equals tbUser.UserID into From1
                              from Add1 in From1.DefaultIfEmpty()
                              where tbEtrust.AuditType.Trim() != "未审核"
                              orderby tbExamine.CommerceID descending
                              select new ExamineVo
                              {
                                  EtrustID = tbExamine.EtrustID,//委托单号
                                  CommerceID = tbExamine.CommerceID,//审核ID
                                  EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                  CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                  EtrustType = tbEtrust.EtrustType,//状态
                                  ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                  CheXing = tbEtrust.CheXing,//车型
                                  CommerceExaminePID = tbExamine.CommerceExaminePID,//商务审核人 
                                  CommerceQiExaminePID = tbExamine.CommerceQiExaminePID,//商务弃审人
                                  FinancingExaninePID = tbExamine.FinancingExaninePID,//财务审核人
                                  FinancingQiExaninePID = tbExamine.FinancingQiExaninePID,//财务弃审人
                                  ShangWuQiShenTime = tbExamine.CommerceQiExamineTime.ToString(),//商务弃审时间
                                  FinancingQiExanineTime = tbExamine.FinancingQiExanineTime,//财务弃审时间
                                  ShangWuShenHeTime = tbExamine.CommerceExamineTime.ToString(),//商务审核时间
                                  FinancingExanineTime = tbExamine.FinancingExanineTime,//财务审核时间
                                  AuditType = tbEtrust.AuditType.Trim(),//审核状态
                                  User = Add.AccountName,//商务弃审人 
                                  User1 = Add1.AccountName,//商务审核人
                                  Cupboard = tbEtrust.Cupboard,//柜号
                                  Seal = tbEtrust.Seal,//封条号
                                  IndentNumber = tbEtrust.IndentNumber,//订单号
                                  ArriveFactoryTime1 = tbEtrust.ArriveFactoryTime.ToString(),//到达工厂时间
                                  ArriveFactoryTimes = tbEtrust.ArriveFactoryTime,
                                  LeftFactoryTime1 = tbEtrust.LeftFactoryTime.ToString(),//离开工厂时间
                                  LeftFactoryTimes = tbEtrust.LeftFactoryTime,
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                              }).ToList();
            if (!string.IsNullOrEmpty(CarriageNumber))
            {
                CarriageNumber = CarriageNumber.Trim();
                listExamine = listExamine.Where(p => p.CarriageNumber.Contains(CarriageNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))
            {
                Cupboard = Cupboard.Trim();
                listExamine = listExamine.Where(p => p.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))
            {
                Seal = Seal.Trim();
                listExamine = listExamine.Where(p => p.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(IndentNumber))
            {
                IndentNumber = IndentNumber.Trim();
                listExamine = listExamine.Where(p => p.IndentNumber.Contains(IndentNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listExamine = listExamine.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustNmber))
            {
                EtrustNmber = EtrustNmber.Trim();
                listExamine = listExamine.Where(p => p.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (TiGuiDiDianID > 0)
            {
                listExamine = listExamine.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)
            {
                listExamine = listExamine.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (!string.IsNullOrEmpty(ArriveFactoryTime) && string.IsNullOrEmpty(LeftFactoryTime))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveFactoryTime);
                    listExamine = listExamine.Where(m => m.ArriveFactoryTimes == Time).ToList();
                }
                catch (Exception)
                {
                    listExamine = listExamine.Where(m => m.CommerceID == CommerceID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(ArriveFactoryTime) && !string.IsNullOrEmpty(LeftFactoryTime))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(LeftFactoryTime);
                    listExamine = listExamine.Where(m => m.ArriveFactoryTimes <= Times).ToList();
                }
                catch (Exception)
                {
                    listExamine = listExamine.Where(m => m.CommerceID == CommerceID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(ArriveFactoryTime) && !string.IsNullOrEmpty(LeftFactoryTime))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveFactoryTime);
                    DateTime Times = Convert.ToDateTime(LeftFactoryTime);
                    listExamine = listExamine.Where(m => m.ArriveFactoryTimes >= Time && m.ArriveFactoryTimes <= Times).ToList();
                }
                catch (Exception)
                {
                    listExamine = listExamine.Where(m => m.CommerceID == CommerceID).ToList();
                }
            }

            //获取当前查询出来的条数

            var totalCount = listExamine.Count();

            List<ExamineVo> listItem = listExamine
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };


            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询审核下方表格的数据
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <param name="EtrustID"></param>
        /// <returns></returns>
        public ActionResult SelectCharge(BsgridPage bsgridPage,int? EtrustID,int? A)
        {
            var listCharge = (from tbCharge in myModels.SYS_Charge
                             join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                             join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                             join tbOfferDetail in myModels.SYS_OfferDetail on tbEtrust.OfferDetailID equals tbOfferDetail.OfferDetailID 
                             into From from Add in From.DefaultIfEmpty() //左链接
                              orderby tbCharge.ChargeID descending
                             select new ExamineVo
                             {
                                 EtrustID = tbEtrust.EtrustID,
                                 ChargeID = tbCharge.ChargeID,
                                 ReckoningUnit = "",
                                 ExpenseName = tbExpense.ExpenseName,
                                 SettleAccounts = tbExpense.SettleAccounts.Trim(),
                                 UnitPrice = tbCharge.UnitPrice,
                                 ChargeType = "正常",  //直接在查出来的表格给默认值，数据库不需要存在的字段
                                 BoxQuantity = Add.BoxQuantity,
                                 Money = Add.Money,
                                 Currency = Add.Currency,
                                 YingShouZ =0,  //这是直接给用来计算的文本框的字段，数据库是没有的
                                 YingFuZ=0 ,
                                 LiRun=0,
                                 Bgs= tbCharge.ReckoningUnit, //本公司的结算代码
                                 Kh= tbCharge.ReckoningUnits  //客户公司的结算代码
                             }).ToList();

            if (EtrustID > 0)
            {
                listCharge = listCharge.Where(m => m.EtrustID == EtrustID).ToList();
            }
            if (listCharge.Count()>0)  // Count 条数
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    if (!string.IsNullOrEmpty(listCharge[i].Bgs))
                    {
                        var B = listCharge[i].Bgs.Trim();  //TissueCode 本公司的组织代码  ChineseAbbreviation  客户简称
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == B).Select(m => m.ChineseAbbreviation).Single();
                        listCharge[i].ReckoningUnit = C;
                    }
                    else if (!string.IsNullOrEmpty(listCharge[i].Kh))
                    {
                        var B = listCharge[i].Kh.Trim();
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Select(m => m.ClientAbbreviation).Single();
                        listCharge[i].ReckoningUnit = C;
                    }
                }
            }
            if (A>0)  //点击上方数据触发下方表格数据并同时把应收，应付和利润计算出来  Sum sel 总数
            {
                listCharge[0].YingShouZ = listCharge.Where(m => m.SettleAccounts == "应收").Sum(sel => sel.UnitPrice);
                listCharge[0].YingFuZ = listCharge.Where(m => m.SettleAccounts == "应付").Sum(sel => sel.UnitPrice);
                listCharge[0].LiRun = listCharge[0].YingShouZ - listCharge[0].YingFuZ;
                return Json(listCharge[0], JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数

            var totalCount = listCharge.Count();

            List<ExamineVo> listItem = listCharge
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);

        }

        /// <summary>
        /// 查询财务结算左边的表
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectFinance(BsgridPage bsgridPage,string ClientCode,string ClientAbbreviation)
        {
            var listFinance = (from tbFinance in myModels.SYS_SignCloseAccount
                               join tbClient in myModels.SYS_Client on tbFinance.ClientCode equals tbClient.ClientCode into From
                               from Add in From.DefaultIfEmpty()//通过客户代码把两个表连接起来
                               join tbMessage in myModels.SYS_Message on tbFinance.ClientCode equals tbMessage.TissueCode into From1
                               from Add1 in From1.DefaultIfEmpty()//通过客户代码把两个表连接起来
                               orderby tbFinance.SignCloseAccountID descending
                               select new ExamineVo
                               {
                                   SignCloseAccountID = tbFinance.SignCloseAccountID, //对账结算ID
                                   ClientAbbreviation = "",  //如果数据库有两个相同的字段(就是两个列放在一个表格下)要放在同一个表格下，就用这种方法，查出来为空，下面用一个if (listFinance.Count()>0) 里面用for循环就可以了
                                   ClientCodee = tbFinance.ClientCode, //用于多条件查询的客户代码
                                   SettleWay = "票结",
                                   Currency = "RMB",
                                   YingShouZ = 0,
                                   YingFuZ = 0,
                                   ClientCode= Add.ClientCode.Trim(), //客户代码
                                   TissueCode = Add1.TissueCode.Trim(), //组织代码 
                               }).ToList();

            if (listFinance.Count()>0)  //条数
            {
                for (int i = 0; i < listFinance.Count(); i++) //两列放在同一列表格显示出来写的判断
                {
                    var A = listFinance[i].ClientCode; //客户代码
                    var B= listFinance[i].TissueCode; //组织代码
                    if (!string.IsNullOrEmpty(A))
                    {
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == A).Select(m => m.ClientAbbreviation).Single();
                        listFinance[i].ClientAbbreviation = C;
                    }
                    else if (!string.IsNullOrEmpty(B))
                    {
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == B).Select(m => m.ChineseAbbreviation).Single();
                        listFinance[i].ClientAbbreviation = C;
                    }
                }
                for (int i = 0; i < listFinance.Count(); i++) //计算方法
                {
                    var CC = listFinance[i].SignCloseAccountID; //对账结算表ID
                    var listJieSuan = from tbJieSuan in myModels.SYS_SignCloseAccount
                                      join tbJieBiao in myModels.SYS_CloseAccount on tbJieSuan.SignCloseAccountID equals tbJieBiao.SignCloseAccountID
                                      join tbCharge in myModels.SYS_Charge on tbJieBiao.ChargeID equals tbCharge.ChargeID
                                      join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                                      where tbJieSuan.SignCloseAccountID == CC && tbJieBiao.SignID == null
                                      select new
                                      {
                                          UnitPrice = tbCharge.UnitPrice, //单价
                                          SettleAccounts = tbExpense.SettleAccounts.Trim(), //收付款类型
                                      };

                    listFinance[i].YingShouZ = listJieSuan.Where(m => m.SettleAccounts == "应收").Sum(sel => sel.UnitPrice); //结算应收总额
                    listFinance[i].YingFuZ = listJieSuan.Where(m => m.SettleAccounts == "应付").Sum(sel => sel.UnitPrice); //计算应付总额
                }
                for (int i = 0; i < listFinance.Count(); i++)
                {
                    if (listFinance[i].YingShouZ==null) //假如应收总额为空，就让它显示为0
                    {
                        listFinance[i].YingShouZ = 0;
                    }
                    if (listFinance[i].YingFuZ == null) //假如应付总额为空，就让它显示为0
                    {
                        listFinance[i].YingFuZ = 0;
                    }
                }
            }
            //多条件查询
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listFinance = listFinance.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientCode))//客户代码
            {
                ClientCode = ClientCode.Trim();
                listFinance = listFinance.Where(p => p.ClientCodee.Contains(ClientCode)).ToList();
            }

            //获取当前查询出来的条数
            var totalCount = listFinance.Count();

            List<ExamineVo> listItem = listFinance
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询财务结算右边的表
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectSign(BsgridPage bsgridPage,int? SignCloseAccountID)
        {
            var listSign = (from tbSign in myModels.SYS_Sign
                           join tbClose in myModels.SYS_SignCloseAccount on tbSign.SignCloseAccountID equals tbClose.SignCloseAccountID
                           join tbPay in myModels.SYS_PayWay on tbSign.PayWayID equals tbPay.PayWayID
                           join tbSettle in myModels.SYS_SettleWay on tbSign.SettleWayID equals tbSettle.SettleWayID
                           where tbSign.FinishType.Trim() == "未核销"
                           orderby tbSign.SignID descending
                           select new ExamineVo
                           {
                               SignCloseAccountID = tbClose.SignCloseAccountID,//对账结算ID
                               SignID = tbSign.SignID,//标记对账ID
                               SignType = tbSign.SignType.Trim(),//对账状态   需要用来筛选的字段在查询时最好要去空格，要不然后面有可能出错
                               SettleNumber = tbSign.SettleNumber,//结算单号
                               PayWay = tbPay.PayWay,//付款方式
                               SettleWay = tbSettle.SettleWay,//结算方式
                               YuSettleDate1 = tbSign.YuSettleDate.ToString(),//预结款日期
                               FinishType = tbSign.FinishType, //结算状态(已核销，未核销)
                               SignExplain = tbSign.SignExplain,//对账说明
                               SignDate = tbSign.SignDate.ToString(),//对账日期
                               SignDate1 = tbSign.SignDate.ToString(),//转换对账日期
                               PayWayID = tbSign.PayWayID,//付款方式ID
                               SettleWayID = tbSign.SettleWayID,//结算方式ID
                               SettleMoney = tbSign.SettleMoney,//结算金额
                               BCYShou = tbSign.BCYShou,//本次应收
                               BCYFu = tbSign.BCYFu,//本次应付
                           }).ToList();

            for (int i = 0; i < listSign.Count(); i++)  //让查出来的金钱(SettleMoney)为正数
            {
                if (listSign[i].SettleMoney < 0)
                {
                    listSign[i].SettleMoney = listSign[i].SettleMoney * -1;  //把为负数的值乘于负一
                }
            }

            if (SignCloseAccountID > 0)
            {
                listSign = listSign.Where(m => m.SignCloseAccountID == SignCloseAccountID).ToList();
            }

            //获取当前查询出来的条数
            var totalCount = listSign.Count();

            List<ExamineVo> listItem = listSign
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 查询财务结算下边的表格数据
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectCloseA(BsgridPage bsgridPage,int? SignID,int? A)
        {
            var listClose = (from tbClose in myModels.SYS_CloseAccount
                            join tbCharge in myModels.SYS_Charge on tbClose.ChargeID equals tbCharge.ChargeID
                            join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                            join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                            orderby tbClose.CloseAccountID descending
                            select new ExamineVo
                            {
                                CloseAccountID = tbClose.CloseAccountID,//结算ID
                                SignID = tbClose.SignID,//对账ID
                                ExpenseName = tbExpense.ExpenseName,//费用项目名称
                                SettleAccounts = tbExpense.SettleAccounts.Trim(),//收付款类型
                                UnitPrice = tbCharge.UnitPrice,//单价
                                YiHeXiao = "0",//已核销金额
                                YiChHeX = "0", //异常核销金额
                                Parities = "1.00",//当前汇率
                                Currency = "RMB",//币种
                                EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                YingShouZ = 0,
                                YingFuZ = 0,
                            }).ToList();

            if (SignID >= 0)
            {
                listClose = listClose.Where(m => m.SignID == SignID).ToList();
            }

            if (A > 0)  //点击上方数据触发下方表格数据并同时把应收，应付和利润计算出来  Sum sel 总数
            {
                listClose[0].YingShouZ = listClose.Where(m => m.SettleAccounts == "应收").Sum(sel => sel.UnitPrice);
                listClose[0].YingFuZ = listClose.Where(m => m.SettleAccounts == "应付").Sum(sel => sel.UnitPrice);
                return Json(listClose[0], JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listClose.Count();

            List<ExamineVo> listItem = listClose
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }



        /// <summary>
        /// 查询新增结算单下方表格的数据
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectInset(BsgridPage bsgridPage,int? ChargeID,string EtrustNmber,string Cupboard,string Seal,string HuaiWeightTime,string HuaiWeightTimes, string ClientAbbreviation,int? TiGuiDiDianID,int? HuaiGuiDiDianID,int? ClientID, int? ShipID,string HangCi,int? PortID,int? GoalHarborID,int? UndertakeID, string EtrustType, int? ChauffeurID,int? VehicleInformationID,int? BracketID,string BookingSpace,string IndentNumber,string WorkCategory,string CheXing,string ReckoningUnit,int? ExpenseID,string SettleAccounts)
        {
            var listCharge = (from tbCharge in myModels.SYS_Charge
                             join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                             join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                             join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                             join tbVehicle in myModels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicle.VehicleInformationID
                             join tbSign in myModels.SYS_CloseAccount on tbCharge.ChargeID equals tbSign.ChargeID
                             where tbSign.SignID == null 
                              orderby tbCharge.ChargeID descending
                             select new ExamineVo
                             {
                                 ChargeID = tbCharge.ChargeID, //收费ID
                                 EtrustID = tbEtrust.EtrustID,//委托单ID
                                 ExpenseName = tbExpense.ExpenseName, //费用名称
                                 ExpenseID = tbExpense.ExpenseID,//费用项目ID
                                 SettleAccounts = tbExpense.SettleAccounts,//收付款类型
                                 UnitPrice = tbCharge.UnitPrice,//单价
                                 Parities = "1",  //汇率
                                 Currency = "RMB",//币种
                                 YiJieMoney = "0",//已结金额
                                 CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                 EtrustType = tbEtrust.EtrustType,//委托单状态
                                 ClientAbbreviation = "",//客户简称
                                 CheXing = tbEtrust.CheXing,//车型
                                 WorkCategory = tbEtrust.WorkCategory,//作业类别
                                 EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                 Cupboard = tbEtrust.Cupboard,//柜号
                                 Seal = tbEtrust.Seal,//封条号
                                 TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                 HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                                 IndentNumber = tbEtrust.IndentNumber,//订单号
                                 HuaiWeightTime1 = tbEtrust.HuaiWeightTime.ToString(),//还柜时间
                                 HuaiWeightTimee = tbEtrust.HuaiWeightTime,
                                 HangCi = tbEtrust.HangCi,//航次
                                 ClientID = tbClient.ClientID,//客户信息(船公司)
                                 ShipID = tbEtrust.ShipID,//船名称
                                 PortID = tbEtrust.PortID,//进出口岸
                                 GoalHarborID = tbEtrust.GoalHarborID,//目的港
                                 UndertakeID = tbEtrust.UndertakeID,//承运公司
                                 ChauffeurID = tbVehicle.ChauffeurID,//司机资料ID
                                 VehicleInformationID = tbEtrust.VehicleInformationID,//车辆信息ID
                                 PlateNumbers = tbVehicle.PlateNumbers,//车牌号
                                 BracketID = tbVehicle.BracketID,//托架资料ID
                                 BookingSpace = tbEtrust.BookingSpace,//订舱号
                                 ReckoningUnit = "", //结算单位
                                 Bgs = tbCharge.ReckoningUnit.Trim(),//结算单位(本公司)
                                 Kh = tbCharge.ReckoningUnits.Trim(), //结算单位(客户)
                                 SignID = tbSign.SignID,//对账ID
                                 SignCloseAccountID = tbSign.SignCloseAccountID,//对账结算ID
                                 CabinetType = tbEtrust.CabinetType,//箱型
                             }).ToList();

            if (listCharge.Count() > 0)  // Count 条数  跳转页面根据对账结算ID去筛选
            {
                var A = Convert.ToInt32(Session["SignCloseAccountID"].ToString()); //接收跳转页面时传过来的对账结算ID 
                var B = myModels.SYS_CloseAccount.Where(m => m.SignCloseAccountID == A).Select(m => m.ChargeID).ToList(); //让结算表里面的对账结算ID等于传过来的对账结算ID，然后查询出收费ID 
                if (B.Count() > 0)
                {
                    var C = B[0]; //声明一个变量接收上面查询出来的收费ID  B[0]获取第一条数据进行筛选
                    var D = myModels.SYS_Charge.Where(m => m.ChargeID == C).ToList(); //声明一个变量接收收费表里的收费ID等于前面用对账结算ID查询出来的收费ID 进行筛选出来相同的数据
                    var E = D[0].ReckoningUnit;  //声明一个变量接收收费ID对应的结算单位(本公司)
                    var F = D[0].ReckoningUnits; //声明一个变量接收收费ID对应的结算单位(客户)
                    if (!string.IsNullOrEmpty(E))
                    {
                        listCharge = listCharge.Where(m => m.Bgs == E.Trim()).ToList();  //本公司
                    }
                    else if (!string.IsNullOrEmpty(F))
                    {
                        listCharge = listCharge.Where(m => m.Kh == F.Trim()).ToList();  //客户
                    }
                }
            }
            if (listCharge.Count() > 0)  //两列数据放在表格的同一列显示出来写的判断，前面查询数据时，显示这一列的字段要让它为空，然后进行下面判断进行显示
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    var A = listCharge[i].Bgs; //本公司   Bgs前面表格查询需要显示出来的字段
                    var B = listCharge[i].Kh; //客户
                    if (!string.IsNullOrEmpty(A)) //字符串不为空
                    {
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == A).Single(); //查出本公司表里面的组织代码等于前面定义的变量查出那一条数据
                        listCharge[i].ClientAbbreviation = C.ChineseAbbreviation; //客户简称
                        listCharge[i].ReckoningUnit = C.ChineseName; //结算单位
                    }
                    else if (!string.IsNullOrEmpty(B))
                    {
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Single();
                        listCharge[i].ClientAbbreviation = C.ClientAbbreviation;
                        listCharge[i].ReckoningUnit = C.ChineseName;
                    }
                }
             }
            //多条件查询
            if (!string.IsNullOrEmpty(EtrustNmber))//委托单号
            {
                EtrustNmber = EtrustNmber.Trim();
                listCharge = listCharge.Where(p => p.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))//柜号
            {
                Cupboard = Cupboard.Trim();
                listCharge = listCharge.Where(p => p.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))//封条号
            {
                Seal = Seal.Trim();
                listCharge = listCharge.Where(p => p.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(HuaiWeightTime) && string.IsNullOrEmpty(HuaiWeightTimes))//还柜时间
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(HuaiWeightTime);
                    listCharge = listCharge.Where(m => m.HuaiWeightTimee == Time).ToList();
                }
                catch (Exception)
                {
                    listCharge = listCharge.Where(m => m.ChargeID == ChargeID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(HuaiWeightTime) && !string.IsNullOrEmpty(HuaiWeightTimes))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(HuaiWeightTimes);
                    listCharge = listCharge.Where(m => m.HuaiWeightTimee <= Times).ToList();
                }
                catch (Exception)
                {
                    listCharge = listCharge.Where(m => m.ChargeID == ChargeID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(HuaiWeightTime) && !string.IsNullOrEmpty(HuaiWeightTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(HuaiWeightTime);
                    DateTime Times = Convert.ToDateTime(HuaiWeightTimes);
                    listCharge = listCharge.Where(m => m.HuaiWeightTimee >= Time && m.HuaiWeightTimee <= Times).ToList();
                }
                catch (Exception)
                {
                    listCharge = listCharge.Where(m => m.ChargeID == ChargeID).ToList();
                }
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listCharge = listCharge.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (TiGuiDiDianID > 0)//提柜地点
            {
                listCharge = listCharge.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)//还柜地点
            {
                listCharge = listCharge.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (ClientID > 0)//船公司
            {
                listCharge = listCharge.Where(m => m.ClientID == ClientID).ToList();
            }
            if (ShipID > 0)//船名
            {
                listCharge = listCharge.Where(m => m.ShipID == ShipID).ToList();
            }
            if (!string.IsNullOrEmpty(HangCi))//航次
            {
                HangCi = HangCi.Trim();
                listCharge = listCharge.Where(p => p.HangCi.Contains(HangCi)).ToList();
            }
            if (PortID > 0)//进出口岸
            {
                listCharge = listCharge.Where(m => m.PortID == PortID).ToList();
            }
            if (GoalHarborID > 0)//目的港
            {
                listCharge = listCharge.Where(m => m.GoalHarborID == GoalHarborID).ToList();
            }
            if (UndertakeID > 0)//承运公司
            {
                listCharge = listCharge.Where(m => m.UndertakeID == UndertakeID).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustType))//状态
            {
                EtrustType = EtrustType.Trim();
                listCharge = listCharge.Where(p => p.EtrustType.Contains(EtrustType)).ToList();
            }
            if (ChauffeurID > 0)//司机
            {
                listCharge = listCharge.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (VehicleInformationID > 0)//车牌号
            {
                listCharge = listCharge.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (BracketID > 0)//托架编号
            {
                listCharge = listCharge.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(BookingSpace))//订舱号
            {
                BookingSpace = BookingSpace.Trim();
                listCharge = listCharge.Where(p => p.BookingSpace.Contains(BookingSpace)).ToList();
            }
            if (!string.IsNullOrEmpty(IndentNumber))//订单号
            {
                IndentNumber = IndentNumber.Trim();
                listCharge = listCharge.Where(p => p.IndentNumber.Contains(IndentNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(WorkCategory))//作业类别
            {
                WorkCategory = WorkCategory.Trim();
                listCharge = listCharge.Where(p => p.WorkCategory.Contains(WorkCategory)).ToList();
            }
            if (!string.IsNullOrEmpty(CheXing))//车型
            {
                CheXing = CheXing.Trim();
                listCharge = listCharge.Where(p => p.CheXing.Contains(CheXing)).ToList();
            }
            if (!string.IsNullOrEmpty(ReckoningUnit))//结算单位
            {
                ReckoningUnit = ReckoningUnit.Trim();
                listCharge = listCharge.Where(p => p.ReckoningUnit.Contains(ReckoningUnit)).ToList();
            }
            if (ExpenseID > 0)//费用名称
            {
                listCharge = listCharge.Where(m => m.ExpenseID == ExpenseID).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))//收付款类型
            {
                SettleAccounts = SettleAccounts.Trim();
                listCharge = listCharge.Where(p => p.SettleAccounts.Contains(SettleAccounts)).ToList();
            }

            //获取当前查询出来的条数
            var totalCount = listCharge.Count();

            List<ExamineVo> listItem = listCharge
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 查询新增结算单明细下方表格的数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectInsertDetail(BsgridPage bsgridPage,int? ChargeID, string EtrustNmber, string Cupboard, string Seal, string HuaiWeightTime, string HuaiWeightTimes, string ClientAbbreviation, int? TiGuiDiDianID, int? HuaiGuiDiDianID, int? ClientID, int? ShipID, string HangCi, int? PortID, int? GoalHarborID, int? UndertakeID, string EtrustType, int? ChauffeurID, int? VehicleInformationID, int? BracketID, string BookingSpace, string IndentNumber, string WorkCategory, string CheXing, string ReckoningUnit, int? ExpenseID, string SettleAccounts)
        {
            var listCharge = (from tbCharge in myModels.SYS_Charge
                              join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                              join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                              join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              join tbVehicle in myModels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicle.VehicleInformationID
                              join tbSign in myModels.SYS_CloseAccount on tbCharge.ChargeID equals tbSign.ChargeID
                              where tbSign.SignID == null
                              orderby tbCharge.ChargeID descending
                              select new ExamineVo
                              {
                                  ChargeID = tbCharge.ChargeID, //收费ID
                                  EtrustID = tbEtrust.EtrustID,//委托单ID
                                  ExpenseName = tbExpense.ExpenseName, //费用名称
                                  ExpenseID = tbExpense.ExpenseID,//费用项目ID
                                  SettleAccounts = tbExpense.SettleAccounts,//收付款类型
                                  UnitPrice = tbCharge.UnitPrice,//单价
                                  Parities = "1",  //汇率
                                  Currency = "RMB",//币种
                                  YiJieMoney = "0",//已结金额
                                  CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                  EtrustType = tbEtrust.EtrustType,//委托单状态
                                  ClientAbbreviation = "",//客户简称
                                  CheXing = tbEtrust.CheXing,//车型
                                  WorkCategory = tbEtrust.WorkCategory,//作业类别
                                  EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                  Cupboard = tbEtrust.Cupboard,//柜号
                                  Seal = tbEtrust.Seal,//封条号
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                                  IndentNumber = tbEtrust.IndentNumber,//订单号
                                  HuaiWeightTime1 = tbEtrust.HuaiWeightTime.ToString(),//还柜时间
                                  HuaiWeightTimee = tbEtrust.HuaiWeightTime,
                                  HangCi = tbEtrust.HangCi,//航次
                                  ClientID = tbClient.ClientID,//客户信息(船公司)
                                  ShipID = tbEtrust.ShipID,//船名称
                                  PortID = tbEtrust.PortID,//进出口岸
                                  GoalHarborID = tbEtrust.GoalHarborID,//目的港
                                  UndertakeID = tbEtrust.UndertakeID,//承运公司
                                  ChauffeurID = tbVehicle.ChauffeurID,//司机资料ID
                                  VehicleInformationID = tbEtrust.VehicleInformationID,//车辆信息ID
                                  PlateNumbers = tbVehicle.PlateNumbers,//车牌号
                                  BracketID = tbVehicle.BracketID,//托架资料ID
                                  BookingSpace = tbEtrust.BookingSpace,//订舱号
                                  ReckoningUnit = "", //结算单位
                                  Bgs = tbCharge.ReckoningUnit.Trim(),//结算单位(本公司)
                                  Kh = tbCharge.ReckoningUnits.Trim(), //结算单位(客户)
                                  SignID = tbSign.SignID,//对账ID
                                  SignCloseAccountID = tbSign.SignCloseAccountID,//对账结算ID
                                  CabinetType = tbEtrust.CabinetType,//箱型
                              }).ToList();

            if (listCharge.Count() > 0)  // Count 条数  跳转页面根据对账结算ID去筛选
            {
                var A = Convert.ToInt32(Session["SignID"].ToString()); //接收跳转页面时传过来的对账结算ID 
                var B = myModels.SYS_CloseAccount.Where(m => m.SignID == A).Select(m => m.ChargeID).ToList(); //让结算表里面的对账结算ID等于传过来的对账结算ID，然后查询出收费ID 
                if (B.Count() > 0)
                {
                    var C = B[0]; //声明一个变量接收上面查询出来的收费ID  B[0]获取第一条数据进行筛选
                    var D = myModels.SYS_Charge.Where(m => m.ChargeID == C).ToList(); //声明一个变量接收收费表里的收费ID等于前面用对账结算ID查询出来的收费ID 进行筛选出来相同的数据
                    var E = D[0].ReckoningUnit;  //声明一个变量接收收费ID对应的结算单位(本公司)
                    var F = D[0].ReckoningUnits; //声明一个变量接收收费ID对应的结算单位(客户)
                    if (!string.IsNullOrEmpty(E))
                    {
                        listCharge = listCharge.Where(m => m.Bgs == E.Trim()).ToList();  //本公司
                    }
                    else if (!string.IsNullOrEmpty(F))
                    {
                        listCharge = listCharge.Where(m => m.Kh == F.Trim()).ToList();  //客户
                    }
                }
            }
            if (listCharge.Count() > 0)  //两列数据放在表格的同一列显示出来写的判断，前面查询数据时，显示这一列的字段要让它为空，然后进行下面判断进行显示
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    var A = listCharge[i].Bgs; //本公司   Bgs前面表格查询需要显示出来的字段
                    var B = listCharge[i].Kh; //客户
                    if (!string.IsNullOrEmpty(A)) //字符串不为空
                    {
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == A).Single(); //查出本公司表里面的组织代码等于前面定义的变量查出那一条数据
                        listCharge[i].ClientAbbreviation = C.ChineseAbbreviation; //客户简称
                        listCharge[i].ReckoningUnit = C.ChineseName; //结算单位
                    }
                    else if (!string.IsNullOrEmpty(B))
                    {
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Single();
                        listCharge[i].ClientAbbreviation = C.ClientAbbreviation;
                        listCharge[i].ReckoningUnit = C.ChineseName;
                    }
                }
            }
            //多条件查询
            if (!string.IsNullOrEmpty(EtrustNmber))//委托单号
            {
                EtrustNmber = EtrustNmber.Trim();
                listCharge = listCharge.Where(p => p.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))//柜号
            {
                Cupboard = Cupboard.Trim();
                listCharge = listCharge.Where(p => p.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))//封条号
            {
                Seal = Seal.Trim();
                listCharge = listCharge.Where(p => p.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(HuaiWeightTime) && string.IsNullOrEmpty(HuaiWeightTimes))//还柜时间
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(HuaiWeightTime);
                    listCharge = listCharge.Where(m => m.HuaiWeightTimee == Time).ToList();
                }
                catch (Exception)
                {
                    listCharge = listCharge.Where(m => m.ChargeID == ChargeID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(HuaiWeightTime) && !string.IsNullOrEmpty(HuaiWeightTimes))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(HuaiWeightTimes);
                    listCharge = listCharge.Where(m => m.HuaiWeightTimee <= Times).ToList();
                }
                catch (Exception)
                {
                    listCharge = listCharge.Where(m => m.ChargeID == ChargeID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(HuaiWeightTime) && !string.IsNullOrEmpty(HuaiWeightTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(HuaiWeightTime);
                    DateTime Times = Convert.ToDateTime(HuaiWeightTimes);
                    listCharge = listCharge.Where(m => m.HuaiWeightTimee >= Time && m.HuaiWeightTimee <= Times).ToList();
                }
                catch (Exception)
                {
                    listCharge = listCharge.Where(m => m.ChargeID == ChargeID).ToList();
                }
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listCharge = listCharge.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (TiGuiDiDianID > 0)//提柜地点
            {
                listCharge = listCharge.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)//还柜地点
            {
                listCharge = listCharge.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (ClientID > 0)//船公司
            {
                listCharge = listCharge.Where(m => m.ClientID == ClientID).ToList();
            }
            if (ShipID > 0)//船名
            {
                listCharge = listCharge.Where(m => m.ShipID == ShipID).ToList();
            }
            if (!string.IsNullOrEmpty(HangCi))//航次
            {
                HangCi = HangCi.Trim();
                listCharge = listCharge.Where(p => p.HangCi.Contains(HangCi)).ToList();
            }
            if (PortID > 0)//进出口岸
            {
                listCharge = listCharge.Where(m => m.PortID == PortID).ToList();
            }
            if (GoalHarborID > 0)//目的港
            {
                listCharge = listCharge.Where(m => m.GoalHarborID == GoalHarborID).ToList();
            }
            if (UndertakeID > 0)//承运公司
            {
                listCharge = listCharge.Where(m => m.UndertakeID == UndertakeID).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustType))//状态
            {
                EtrustType = EtrustType.Trim();
                listCharge = listCharge.Where(p => p.EtrustType.Contains(EtrustType)).ToList();
            }
            if (ChauffeurID > 0)//司机
            {
                listCharge = listCharge.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (VehicleInformationID > 0)//车牌号
            {
                listCharge = listCharge.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (BracketID > 0)//托架编号
            {
                listCharge = listCharge.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(BookingSpace))//订舱号
            {
                BookingSpace = BookingSpace.Trim();
                listCharge = listCharge.Where(p => p.BookingSpace.Contains(BookingSpace)).ToList();
            }
            if (!string.IsNullOrEmpty(IndentNumber))//订单号
            {
                IndentNumber = IndentNumber.Trim();
                listCharge = listCharge.Where(p => p.IndentNumber.Contains(IndentNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(WorkCategory))//作业类别
            {
                WorkCategory = WorkCategory.Trim();
                listCharge = listCharge.Where(p => p.WorkCategory.Contains(WorkCategory)).ToList();
            }
            if (!string.IsNullOrEmpty(CheXing))//车型
            {
                CheXing = CheXing.Trim();
                listCharge = listCharge.Where(p => p.CheXing.Contains(CheXing)).ToList();
            }
            if (!string.IsNullOrEmpty(ReckoningUnit))//结算单位
            {
                ReckoningUnit = ReckoningUnit.Trim();
                listCharge = listCharge.Where(p => p.ReckoningUnit.Contains(ReckoningUnit)).ToList();
            }
            if (ExpenseID > 0)//费用名称
            {
                listCharge = listCharge.Where(m => m.ExpenseID == ExpenseID).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))//收付款类型
            {
                SettleAccounts = SettleAccounts.Trim();
                listCharge = listCharge.Where(p => p.SettleAccounts.Contains(SettleAccounts)).ToList();
            }

            //获取当前查询出来的条数
            var totalCount = listCharge.Count();

            List<ExamineVo> listItem = listCharge
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }



        /// <summary>
        /// 查询实收实付的上方表格数据(核销)
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectCalculate(BsgridPage bsgridPage,string CalculateNumber,int? ClientID,int? MessageID, string BankName,string SFDate,int? PayWayID,string Currency,string SFMoney, string AlreadyCancelMoney, string StayCancelMoney)
        {
            var listCalculate = (from tbCalculate in myModels.SYS_Calculate
                                join tbPayWay in myModels.SYS_PayWay on tbCalculate.PayWayID equals tbPayWay.PayWayID
                                orderby tbCalculate.CalculateID descending
                                select new ExamineVo
                                {
                                    CalculateID = tbCalculate.CalculateID,//计费ID
                                    SignID = tbCalculate.SignID,//对账ID
                                    CalculateState = "已审核",//计费状态
                                    CalculateNumber = tbCalculate.CalculateNumber.Trim(),//收付款单号
                                    SettleAccounts = tbCalculate.SettleAccounts.Trim(),//收付款类型
                                    SFMoney = tbCalculate.SFMoney,//收付金额
                                    AlreadyCancelMoney = tbCalculate.AlreadyCancelMoney,//已核销金额
                                    StayCancelMoney = tbCalculate.StayCancelMoney,//待核销金额
                                    Currency = tbCalculate.Currency.Trim(),//币种
                                    SFDate1 = tbCalculate.SFDate.ToString(),//收付日期
                                    BankName = tbCalculate.BankName.Trim(),//银行名称
                                    BankAccount = tbCalculate.BankAccount.Trim(),//银行账号
                                    PayWay = tbPayWay.PayWay.Trim(),//付款方式
                                    ChequeNumber = tbCalculate.ChequeNumber.Trim(),//支票号
                                    ClientID = tbCalculate.ClientID,//客户ID  结算单位多条件查询
                                    PayWayID = tbCalculate.PayWayID,//付款方式ID
                                    ReckoningUnit = "",
                                    MessageID = tbCalculate.MessageID,
                                }).ToList();

            if (listCalculate.Count() > 0) // Count 条数 两列放在同一列显示，在上面声明一个空的字段在表格显示，然后判断查询该显示哪张表的名称
            {
                for (int i = 0; i < listCalculate.Count(); i++)
                {
                    if (listCalculate[i].MessageID > 0)
                    {
                        var B = listCalculate[i].MessageID;  //listCalculate[i] 前面查询要用ToList()
                        var C = myModels.SYS_Message.Where(m => m.MessageID == B).Select(m => m.ChineseAbbreviation).Single();
                        listCalculate[i].ReckoningUnit = C;
                    }
                    else if (listCalculate[i].ClientID > 0)
                    {
                        var B = listCalculate[i].ClientID;
                        var C = myModels.SYS_Client.Where(m => m.ClientID == B).Select(m => m.ClientAbbreviation).Single();
                        listCalculate[i].ReckoningUnit = C;
                    }
                }
            }


            if (!string.IsNullOrEmpty(CalculateNumber))//收付款单号
            {
                CalculateNumber = CalculateNumber.Trim();
                listCalculate = listCalculate.Where(p => p.CalculateNumber.Contains(CalculateNumber)).ToList();
            }
            if (ClientID > 0)//结算单位
            {
                listCalculate = listCalculate.Where(m => m.ClientID == ClientID).ToList();
            }
            if (MessageID > 0)//结算单位
            {
                listCalculate = listCalculate.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(BankName))//银行名称
            {
                BankName = BankName.Trim();
                listCalculate = listCalculate.Where(p => p.BankName.Contains(BankName)).ToList();
            }
            if (!string.IsNullOrEmpty(SFDate))//收付日期
            {
                DateTime SFDate1 = Convert.ToDateTime(SFDate);
                listCalculate = listCalculate.Where(p => p.SFDate1.Contains(SFDate)).ToList();
            }
            if (PayWayID > 0)//付款方式ID
            {
                listCalculate = listCalculate.Where(m => m.PayWayID == PayWayID).ToList();
            }
            if (!string.IsNullOrEmpty(Currency))//币种
            {
                Currency = Currency.Trim();
                listCalculate = listCalculate.Where(p => p.Currency.Contains(Currency)).ToList();
            }
            if (!string.IsNullOrEmpty(SFMoney))//收付金额
            {
                SFMoney = SFMoney.Trim();
                listCalculate = listCalculate.Where(p => p.SFMoney.ToString().Contains(SFMoney)).ToList();
            }
            if (!string.IsNullOrEmpty(AlreadyCancelMoney))//已核销金额
            {
                AlreadyCancelMoney = AlreadyCancelMoney.Trim();
                listCalculate = listCalculate.Where(p => p.AlreadyCancelMoney.ToString().Contains(AlreadyCancelMoney)).ToList();
            }
            if (!string.IsNullOrEmpty(StayCancelMoney))//待核销金额
            {
                StayCancelMoney = StayCancelMoney.Trim();
                listCalculate = listCalculate.Where(p => p.StayCancelMoney.ToString().Contains(StayCancelMoney)).ToList();
            }

            //获取当前查询出来的条数
            var totalCount = listCalculate.Count();

            List<ExamineVo> listItem = listCalculate
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);

        }


        /// <summary>
        /// 应收应付下方左边的表数据(未核销)
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectSignY(BsgridPage bsgridPage)
        {
            var listSign = (from tbSign in myModels.SYS_Sign
                           join tbClose in myModels.SYS_SignCloseAccount on tbSign.SignCloseAccountID equals tbClose.SignCloseAccountID
                           join tbPay in myModels.SYS_PayWay on tbSign.PayWayID equals tbPay.PayWayID
                           join tbSettle in myModels.SYS_SettleWay on tbSign.SettleWayID equals tbSettle.SettleWayID
                           where tbSign.FinishType.Trim() == "未核销" && tbSign.SignType.Trim()=="审核"
                           orderby tbSign.SignID descending
                           select new ExamineVo
                           {
                               SignCloseAccountID = tbClose.SignCloseAccountID,//对账结算ID
                               SignID = tbSign.SignID,//标记对账ID
                               SignType = tbSign.SignType.Trim(),//对账状态   需要用来筛选的字段在查询时最好要去空格，要不然后面有可能出错
                               SettleNumber = tbSign.SettleNumber,//结算单号
                               PayWay = tbPay.PayWay,//付款方式
                               SettleWay = tbSettle.SettleWay,//结算方式
                               YuSettleDate1 = tbSign.YuSettleDate.ToString(),//预结款日期
                               FinishType = tbSign.FinishType, //结算状态(已核销，未核销)
                               SignExplain = tbSign.SignExplain,//对账说明
                               SignDate = tbSign.SignDate.ToString(),//对账日期
                               SignDate1 = tbSign.SignDate.ToString(),//转换对账日期
                               PayWayID = tbSign.PayWayID,//付款方式ID
                               SettleWayID = tbSign.SettleWayID,//结算方式ID
                               SettleMoney = tbSign.SettleMoney,//结算金额(让查出来的数据为正数)
                               BCYShou = tbSign.BCYShou,//本次应收
                               BCYFu = tbSign.BCYFu,//本次应付
                               Currency = "RMB",
                           }).ToList();
            for (int i = 0; i < listSign.Count(); i++)  //让查出来的金钱为正数
            {
                if (listSign[i].SettleMoney < 0)
                {
                    listSign[i].SettleMoney=listSign[i].SettleMoney * -1;
                }
            }            
               

            //获取当前查询出来的条数
            var totalCount = listSign.Count();

            List<ExamineVo> listItem = listSign
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }



        /// <summary>
        /// 查询应收应付下方左边已核销的数据
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectSignn(BsgridPage bsgridPage)
        {
            var listSign = (from tbSign in myModels.SYS_Sign
                           join tbClose in myModels.SYS_SignCloseAccount on tbSign.SignCloseAccountID equals tbClose.SignCloseAccountID
                           join tbPay in myModels.SYS_PayWay on tbSign.PayWayID equals tbPay.PayWayID
                           join tbSettle in myModels.SYS_SettleWay on tbSign.SettleWayID equals tbSettle.SettleWayID
                           where tbSign.FinishType.Trim() == "已核销"
                           orderby tbSign.SignID descending
                           select new ExamineVo
                           {
                               SignCloseAccountID = tbClose.SignCloseAccountID,//对账结算ID
                               SignID = tbSign.SignID,//标记对账ID
                               SignType = tbSign.SignType.Trim(),//对账状态   需要用来筛选的字段在查询时最好要去空格，要不然后面有可能出错
                               SettleNumber = tbSign.SettleNumber,//结算单号
                               PayWay = tbPay.PayWay,//付款方式
                               SettleWay = tbSettle.SettleWay,//结算方式
                               YuSettleDate1 = tbSign.YuSettleDate.ToString(),//预结款日期
                               FinishType = tbSign.FinishType, //结算状态(已核销，未核销)
                               SignExplain = tbSign.SignExplain,//对账说明
                               SignDate = tbSign.SignDate.ToString(),//对账日期
                               SignDate1 = tbSign.SignDate.ToString(),//转换对账日期
                               PayWayID = tbSign.PayWayID,//付款方式ID
                               SettleWayID = tbSign.SettleWayID,//结算方式ID
                               SettleMoney1 = tbSign.SettleMoney,//结算金额
                               BCYShou = tbSign.BCYShou,//本次应收
                               BCYFu = tbSign.BCYFu,//本次应付
                               Currency = "RMB",
                           }).ToList();

            for (int i = 0; i < listSign.Count(); i++)  //让查出来的金钱为正数
            {
                if (listSign[i].SettleMoney1 < 0)
                {
                    listSign[i].SettleMoney1 = listSign[i].SettleMoney1 * -1;
                }
            }

            //获取当前查询出来的条数
            var totalCount = listSign.Count();

            List<ExamineVo> listItem = listSign
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ExamineVo> bsgrid = new Bsgrid<ExamineVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }




        #endregion

        #region 保存代码

        /// <summary>
        /// 保存商务审核
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveExamine(int EtrustID)
        {
            try
            {
                SYS_Etrust mySubscibe = (from tbspecial in myModels.SYS_Etrust
                                         where tbspecial.EtrustID == EtrustID
                                         select tbspecial).Single();

                SYS_Commercel mySubscibe1 = (from tbspecial in myModels.SYS_Commercel
                                             where tbspecial.EtrustID == EtrustID  //通过审核表里的委托表ID等于委托表的委托单ID查询出来
                                             select tbspecial).Single();

                mySubscibe.AuditType = "已审核";

                mySubscibe1.CommerceExaminePID = Convert.ToInt32(Session["UserID"].ToString()); //用Session获取UserID里面的用户名
                mySubscibe1.CommerceExamineTime = DateTime.Now;

                myModels.Entry(mySubscibe).State = System.Data.Entity.EntityState.Modified;
                myModels.Entry(mySubscibe1).State = System.Data.Entity.EntityState.Modified;

                //实例化对象（新增）
                SYS_CloseAccount myInserts = new SYS_CloseAccount(); //结算表
                SYS_SignCloseAccount MyAccout = new SYS_SignCloseAccount(); //对账结算表

                //List<string> ClientCode = myModels.SYS_Etrust.Where(m => m.EtrustID == EtrustID).Select(m => m.ClientCode).ToList();//客户代码
                var Charge = (from tbCharge in myModels.SYS_Charge
                               where tbCharge.EtrustID == EtrustID
                               select new
                               {
                                   ChargeID= tbCharge.ChargeID,//收费ID
                                   ReckoningUnit = tbCharge.ReckoningUnit.Trim(),//本公司的结算单位
                                   ReckoningUnits = tbCharge.ReckoningUnits.Trim(),//客户公司的结算单位
                               }).ToList();

                for (int i = 0; i < Charge.Count(); i++)
                {
                    var A = Charge[i].ReckoningUnit;
                    var B = Charge[i].ReckoningUnits;
                    if (!string.IsNullOrEmpty(A))
                    {
                        var C = myModels.SYS_SignCloseAccount.Where(m => m.ClientCode.Trim() == A).ToList();
                        if (C.Count()==0)
                        {
                            MyAccout.ClientCode = A;
                            myModels.SYS_SignCloseAccount.Add(MyAccout);
                            myModels.SaveChanges();
                        }
                    }
                    else if (!string.IsNullOrEmpty(B))
                    {
                        var C = myModels.SYS_SignCloseAccount.Where(m => m.ClientCode.Trim() == B).ToList();
                        if (C.Count() == 0)
                        {
                            MyAccout.ClientCode = B;
                            myModels.SYS_SignCloseAccount.Add(MyAccout);
                            myModels.SaveChanges();
                        }
                    }
                }
                              
                int BB = 0;  //声明一个int 类型变量
                for (int i = 0; i < Charge.Count(); i++)
                {
                    var j = Charge[i].ChargeID;
                    var ChargeID1 = (from tb in myModels.SYS_CloseAccount
                                     where tb.ChargeID == j
                                     select tb).Count();
                    if (ChargeID1==0)
                    {
                        BB++;
                    }
                }
                if (BB== Charge.Count())
                {

                    //for (int i = 0; i < Charge.Count(); i++)
                    //{
                    //    myInserts.ChargeID = Charge[i].ChargeID;
                    //    //var M = (from tb in myModels.SYS_SignCloseAccount
                    //    //                   where tb.ClientCode == A
                    //    //                   select tb).ToList();  //判断数据库有没有相同的客户代码
                    //    //myInserts.SignCloseAccountID = M[0].SignCloseAccountID;
                    //    myInserts.SignCloseAccountID = M[0].SignCloseAccountID;
                    //    myModels.SYS_CloseAccount.Add(myInserts);
                    //    myModels.SaveChanges();
                    //}
                    for (int i = 0; i < Charge.Count(); i++)
                    {
                        var A = Charge[i].ReckoningUnit;//本公司结算单位
                        var B = Charge[i].ReckoningUnits;//客户公司结算单位
                        if (!string.IsNullOrEmpty(A))//假如本公司的结算单位不为空，就执行下面的代码
                        {
                            var C = myModels.SYS_SignCloseAccount.Where(m => m.ClientCode.Trim() == A).Single();//声明一个变量接收用结算表里面的客户代码等于本公司的结算代码查出来的那一条的数据
                            myInserts.ChargeID = Charge[i].ChargeID;
                            myInserts.SignCloseAccountID = C.SignCloseAccountID;
                            myModels.SYS_CloseAccount.Add(myInserts);
                           
                        }
                        else if (!string.IsNullOrEmpty(B))
                        {
                            var C = myModels.SYS_SignCloseAccount.Where(m => m.ClientCode.Trim() == B).Single();
                            myInserts.ChargeID = Charge[i].ChargeID;
                            myInserts.SignCloseAccountID = C.SignCloseAccountID;
                            myModels.SYS_CloseAccount.Add(myInserts);
                           
                        }
                        myModels.SaveChanges();
                    }

                }
                
                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json("failed", JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 保存商务弃审
        /// </summary>
        /// <param name="EtrustID"></param>
        /// <returns></returns>
        public ActionResult SaveAbandon(int EtrustID)
        {
            try
            {
                SYS_Etrust mySubscibe = (from tbspecial in myModels.SYS_Etrust
                                         where tbspecial.EtrustID == EtrustID
                                         select tbspecial).Single();

                SYS_Commercel mySubscibe1 = (from tbspecial in myModels.SYS_Commercel
                                             where tbspecial.EtrustID == EtrustID  //通过审核表里的委托表ID等于委托表的委托单ID查询出来
                                             select tbspecial).Single();


                mySubscibe.AuditType = "未审核";

                mySubscibe1.CommerceQiExaminePID = Convert.ToInt32(Session["UserID"].ToString()); //用Session获取UserID里面的用户名
                mySubscibe1.CommerceQiExamineTime = DateTime.Now;

                myModels.Entry(mySubscibe).State = System.Data.Entity.EntityState.Modified;
                myModels.Entry(mySubscibe1).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();

                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json("failed", JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 保存新增结算方式数据
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertSettleWay(string SettleWay)
        {
            string strMsg = "failed";
            //实例化对象
            SYS_SettleWay mySettle = new SYS_SettleWay();
            //写入数据
            mySettle.SettleWay = SettleWay;
            //写入数据库
            myModels.SYS_SettleWay.Add(mySettle);
            //保存数据库
            if (myModels.SaveChanges() > 0)
            {
                strMsg = "success";
            }
            else
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存新增付款方式数据
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertPayWay(string PayWay)
        {
            string strMsg = "failed";
            //实例化对象
            SYS_PayWay myPayWay = new SYS_PayWay();
            //写入数据
            myPayWay.PayWay = PayWay;
            //写入数据库
            myModels.SYS_PayWay.Add(myPayWay);
            //保存数据库
            if (myModels.SaveChanges() > 0)
            {
                strMsg = "success";
            }
            else
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存标记对账模态框的代码
        /// </summary>
        /// <returns></returns>
        public ActionResult InsetSign(string myArray, SYS_Sign mySign)
        {
            string strMsg = "failed";
            //写入数据
            mySign.FinishType = "未核销";
            mySign.SignType = "审核";
            mySign.SignCloseAccountID = Convert.ToInt32(Session["SignCloseAccountID"]);//获取页面跳转时传过来的对账结算ID保存到对账表里面 

            //写入数据库
            myModels.SYS_Sign.Add(mySign);
            //保存数据库
            if (myModels.SaveChanges() > 0)
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < AAA.Length; i++)
                {
                    var C = Convert.ToInt32(AAA[i]);
                    List<SYS_CloseAccount> listClose = (from tbClose in myModels.SYS_CloseAccount
                                                        where tbClose.ChargeID == C
                                                        select tbClose).ToList();

                    if (listClose.Count() > 0)
                    {
                        for (int j = 0; j < listClose.Count(); j++)
                        {
                            listClose[j].SignID = mySign.SignID;
                            myModels.Entry(listClose[j]).State = System.Data.Entity.EntityState.Modified;
                            myModels.SaveChanges();
                        }
                    }
                }
                strMsg = "success";
            }
            else
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存财务审核代码
        /// </summary>
        /// <returns></returns>
        public ActionResult CaiWuShenHe(int SignID)
        {
            try
            {
                SYS_Sign mySign = (from tbSign in myModels.SYS_Sign
                                   where tbSign.SignID == SignID
                                   select tbSign).Single();
                mySign.SignType = "审核";//审核状态
                myModels.Entry(mySign).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json("failed", JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 保存财务弃审代码
        /// </summary>
        /// <returns></returns>
        public ActionResult CaiWuQiShen(int SignID)
        {
            try
            {
                SYS_Sign mySign = (from tbSign in myModels.SYS_Sign
                                   where tbSign.SignID == SignID
                                   select tbSign).Single();
                mySign.SignType = "制单";//制单状态
                myModels.Entry(mySign).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json("failed", JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 点击编辑模态框的确定按钮，保存修改财务结算右边表格的数据的代码 
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveUpdate(int SignID,string SignExplain,short SettleWayID,short PayWayID)
        {
            string strMsg = "failed";
            try
            {
                Models.SYS_Sign mySign = myModels.SYS_Sign.Where(m => m.SignID == SignID).Single();

                //修改的数据
                mySign.SignExplain = SignExplain;
                mySign.SettleWayID = SettleWayID;
                mySign.PayWayID = PayWayID;
                //获取和实质对象实体的状态=EntityState的枚举值
                myModels.Entry(mySign).State = System.Data.Entity.EntityState.Modified;
                //保存修改
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 财务结算删除按钮，点击上方删除按钮删除数据  (先修改再删除)右边
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteSign(int SignID)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                List<SYS_CloseAccount> listClose = (from tbClose in myModels.SYS_CloseAccount
                                            where tbClose.SignID == SignID
                                            select tbClose).ToList();   //如果两个表同时操作，可以传两张表同时有的ID来查出需要操作的数据
                for (int i = 0; i < listClose.Count(); i++)   //多条修改，用for循环，先修改，再删除
                {
                    //修改数据
                    listClose[i].SignID = null;

                    //获取和实质对象实体的状态=EntityState的枚举值
                    myModels.Entry(listClose[i]).State = System.Data.Entity.EntityState.Modified;
                }
          
                if (myModels.SaveChanges()>0)
                {
                    List<SYS_Sign> listSign = (from tbSign in myModels.SYS_Sign
                                               where tbSign.SignID == SignID
                                               select tbSign).ToList(); //查出要删除的数据

                    myModels.SYS_Sign.RemoveRange(listSign);
                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 点击确认按钮，新增明细单保存新的对账收费项目
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveInsertDetail(string myArray,decimal BCYShou, decimal BCYFu, decimal SettleMoney)
        {
            string strMsg = "failed";
            try
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < AAA.Length; i++)
                {
                    var C = Convert.ToInt32(AAA[i]);
                    List<SYS_CloseAccount> listClose = (from tbClose in myModels.SYS_CloseAccount
                                                        where tbClose.ChargeID == C
                                                        select tbClose).ToList();

                    for (int j = 0; j < listClose.Count(); j++)
                    {
                        listClose[j].SignID = Convert.ToInt32(Session["SignID"]); //接收跳转页面传过来的ID
                        myModels.Entry(listClose[j]).State = System.Data.Entity.EntityState.Modified;
                    }
                }
                if (myModels.SaveChanges() > 0)
                {
                    var SignID = Convert.ToInt32(Session["SignID"]);

                    SYS_Sign mySign = (from tbSign in myModels.SYS_Sign
                                       where tbSign.SignID == SignID
                                       select tbSign).Single();
                    //需要修改的数据
                    var AA = mySign.BCYShou;  //查出数据库原本的值
                    var BB = mySign.BCYFu;
                    var CC = mySign.SettleMoney;

                    mySign.BCYShou = AA + BCYShou;
                    mySign.BCYFu = BB + BCYFu;
                    mySign.SettleMoney = CC + SettleMoney;
                    myModels.Entry(mySign).State = System.Data.Entity.EntityState.Modified;

                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 点击表格数据行的编辑按钮，进行编辑保存
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveUpdateDetail(int ChargeID,decimal UnitPrice)
        {
            string strMsg = "failed";
            try
            {
                Models.SYS_Charge myCharge = myModels.SYS_Charge.Where(m => m.ChargeID == ChargeID).Single();

                //修改的数据
                myCharge.ChargeID = ChargeID;
                myCharge.UnitPrice = UnitPrice;

                //获取和实质对象实体的状态=EntityState的枚举值
                myModels.Entry(myCharge).State = System.Data.Entity.EntityState.Modified;
                //保存修改
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存点击下方删除按钮的代码(修改结算表的对账ID为空)
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveDeleteDetail(int CloseAccountID,int SignID)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                List<SYS_CloseAccount> listClose = (from tbClose in myModels.SYS_CloseAccount
                                                    where tbClose.CloseAccountID == CloseAccountID
                                                    select tbClose).ToList();   //如果两个表同时操作，可以传两张表同时有的ID来查出需要操作的数据

                var listCloseA = (from tbCloseA in myModels.SYS_CloseAccount
                                  join tbCharge in myModels.SYS_Charge on tbCloseA.ChargeID equals tbCharge.ChargeID
                                  join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                                  where tbCloseA.CloseAccountID == CloseAccountID
                                  select new ExamineVo
                                  {
                                      SettleAccounts = tbExpense.SettleAccounts.Trim(),
                                      UnitPrice = tbCharge.UnitPrice,
                                  }).ToList();
                if (listCloseA.Count() > 0)  //点击上方数据触发下方表格数据并同时把应收，应付和利润计算出来  Sum sel 总数
                {
                    listCloseA[0].YingShouZ = listCloseA.Where(m => m.SettleAccounts == "应收").Sum(sel => sel.UnitPrice);
                    listCloseA[0].YingFuZ = listCloseA.Where(m => m.SettleAccounts == "应付").Sum(sel => sel.UnitPrice);
                    listCloseA[0].LiRun = listCloseA[0].YingShouZ - listCloseA[0].YingFuZ;
                }

                for (int i = 0; i < listClose.Count(); i++)   //多条修改，用for循环，先修改
                {
                    //修改数据
                    listClose[i].SignID = null;

                    // 获取和实质对象实体的状态 = EntityState的枚举值
                    myModels.Entry(listClose[i]).State = System.Data.Entity.EntityState.Modified;
                }
                if (myModels.SaveChanges() > 0)
                {

                    SYS_Sign mySign = (from tbSign in myModels.SYS_Sign
                                       where tbSign.SignID == SignID
                                       select tbSign).Single(); //查出要修改的数据

                    var AA = mySign.BCYShou;
                    var BB = mySign.BCYFu;
                    var CC = mySign.SettleMoney;

                    mySign.BCYShou = AA - listCloseA[0].YingShouZ;
                    mySign.BCYFu = BB - listCloseA[0].YingFuZ;
                    mySign.SettleMoney = CC- listCloseA[0].LiRun;

                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存新增计费单
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveCalculate(SYS_Calculate myCalculate)
        {
            string strMsg = "failed";
            try
            {
                //写入数据，一个文本框的值(数据)保存到两个字段里
                myCalculate.StayCancelMoney = myCalculate.SFMoney;
                //写入数据库
                myModels.SYS_Calculate.Add(myCalculate);
                //保存数据库
                myModels.SaveChanges();

                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存修改计费单
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateCalculate(SYS_Calculate myCalculate)
        {
            string strMsg = "failed";
            try
            {
                SYS_Calculate myCalculatee = (from tbCalculate in myModels.SYS_Calculate
                                              where tbCalculate.CalculateID == myCalculate.CalculateID
                                              select tbCalculate).Single();

                //修改数据
                myCalculatee.CalculateID = myCalculate.CalculateID;
                myCalculatee.CalculateNumber = myCalculate.CalculateNumber;
                myCalculatee.SettleAccounts = myCalculate.SettleAccounts;
                myCalculatee.PayWayID = myCalculate.PayWayID;
                myCalculatee.Currency = myCalculate.Currency;
                myCalculatee.ClientID = myCalculate.ClientID;
                myCalculatee.MessageID = myCalculate.MessageID;
                myCalculatee.SFMoney = myCalculate.SFMoney;
                myCalculatee.StayCancelMoney = myCalculate.SFMoney;
                myCalculatee.SFDate = myCalculate.SFDate;
                myCalculatee.BankName = myCalculate.BankName;
                myCalculatee.BankAccount = myCalculate.BankAccount;
                myCalculatee.ChequeNumber = myCalculate.ChequeNumber;
                myCalculatee.Remark = myCalculate.Remark;

                myModels.Entry(myCalculatee).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                strMsg = "success";
                return Json(strMsg, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 删除计费单
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteShiShouShiFu(int CalculateID)
        {
            string strMsg = "failed";
            try
            {
                List<SYS_Calculate> listCalculate = (from tbCalculate in myModels.SYS_Calculate
                                                     where tbCalculate.CalculateID == CalculateID
                                                     select tbCalculate).ToList(); //查出要删除的数据

                myModels.SYS_Calculate.RemoveRange(listCalculate);
                myModels.SaveChanges();
                strMsg = "success";

            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 点击核销按钮，保存核销代码
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveVerification(int CalculateID,int SignID,decimal StayCancelMoney,decimal AlreadyCancelMoney)
        {
            string strMsg = "failed";
            try
            {
                SYS_Calculate myCalculate = (from tbCalculate in myModels.SYS_Calculate
                                             where tbCalculate.CalculateID == CalculateID
                                             select tbCalculate).Single(); //计费表

                SYS_Sign mySign = (from tbSign in myModels.SYS_Sign
                                   where tbSign.SignID == SignID
                                   select tbSign).Single();//对账表

                mySign.FinishType = "已核销";//已核销状态
                myCalculate.StayCancelMoney = StayCancelMoney;
                myCalculate.AlreadyCancelMoney = AlreadyCancelMoney;

                myModels.Entry(myCalculate).State = System.Data.Entity.EntityState.Modified;
                myModels.Entry(mySign).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        #endregion

        #region 查询下拉框

        /// <summary>
        /// 查询提还柜地下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelecrDiDian()
        {
            var listDiDian = (from tbDiDian in myModels.SYS_Mention
                              select new SelectVo
                              {
                                  id = tbDiDian.MentionID,
                                  text = tbDiDian.Abbreviation
                              }).ToList();
            listDiDian = Common.Tools.SetSelectJson(listDiDian);

            return Json(listDiDian, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询船名下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectShip()
        {
            var listShip = (from tbShip in myModels.SYS_Ship
                            select new SelectVo
                            {
                                id = tbShip.ShipID,
                                text = tbShip.ChineseName
                            }).ToList();
            listShip = Common.Tools.SetSelectJson(listShip);
            return Json(listShip, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定司机下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectChauffeur()
        {
            var listChauffeur = (from tbChauffeur in myModels.SYS_Chauffeur
                                 select new SelectVo
                                 {
                                     id = tbChauffeur.ChauffeurID,
                                     text = tbChauffeur.ChauffeurName,
                                 }).ToList();
            listChauffeur = Common.Tools.SetSelectJson(listChauffeur);
            return Json(listChauffeur, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定车牌下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectPlateNumbers()
        {
            var listPlate = (from tbPlate in myModels.SYS_VehicleInformation
                             select new SelectVo
                             {
                                 id = tbPlate.VehicleInformationID,
                                 text = tbPlate.PlateNumbers,
                             }).ToList();
            listPlate = Common.Tools.SetSelectJson(listPlate);
            return Json(listPlate, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定费用名称下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectExpense()
        {
            var listExpense = (from tbExpense in myModels.SYS_Expense
                               select new SelectVo
                               {
                                   id = tbExpense.ExpenseID,
                                   text = tbExpense.ExpenseName,
                               }).ToList();
            listExpense = Common.Tools.SetSelectJson(listExpense);
            return Json(listExpense, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定托架编号下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectBracket()
        {
            var listBracket = (from tbBracket in myModels.SYS_Bracket
                               select new SelectVo
                               {
                                   id = tbBracket.BracketID,
                                   text = tbBracket.BracketCode,
                               }).ToList();
            listBracket = Common.Tools.SetSelectJson(listBracket);
            return Json(listBracket, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定港口资料下拉框(进出口岸，目的港)
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectPort()
        {
            var listPort = (from tbPort in myModels.SYS_Port
                            select new SelectVo
                            {
                                id = tbPort.PortID,
                                text = tbPort.ChineseName,
                            }).ToList();
            listPort = Common.Tools.SetSelectJson(listPort);
            return Json(listPort, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询船公司下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectClient()
        {
            var listClient = (from tbClientType in myModels.SYS_ClientType
                              join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                              where tbClientType.ClientType.Trim() == "船公司"
                              select new SelectVo
                              {
                                  id = tbClient.ClientID,
                                  text = tbClient.ClientAbbreviation,
                              }).ToList();
            listClient = Common.Tools.SetSelectJson(listClient);
            return Json(listClient, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询拖车公司下拉框(承运公司)
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectClientt()
        {
            var listUndertake = (from tbUndertake in myModels.SYS_ClientType
                                 join tbClient in myModels.SYS_Client on tbUndertake.ClientID equals tbClient.ClientID
                                 where tbUndertake.ClientType.Trim() == "拖车公司"
                                 select new SelectVo
                                 {
                                     id = tbClient.ClientID,
                                     text = tbClient.ChineseName
                                 }).ToList();
            listUndertake = Common.Tools.SetSelectJson(listUndertake);
            return Json(listUndertake, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询结算方式下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectSettleWay()
        {
            var listSettleWay = (from tbSettleWay in myModels.SYS_SettleWay
                                 select new SelectVo
                                 {
                                     id = tbSettleWay.SettleWayID,
                                     text = tbSettleWay.SettleWay,
                                 }).ToList();
            listSettleWay = Common.Tools.SetSelectJson(listSettleWay);
            return Json(listSettleWay, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询付款方式下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectPayWay()
        {
            var listPayWay = (from tbPayWay in myModels.SYS_PayWay
                              select new SelectVo
                              {
                                  id = tbPayWay.PayWayID,
                                  text = tbPayWay.PayWay,
                              }).ToList();
            listPayWay = Common.Tools.SetSelectJson(listPayWay);
            return Json(listPayWay, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询客户信息表的所有客户名称
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectClientr()
        {
            var listClient = (from tbClient in myModels.SYS_Client
                              select new SelectVo
                              {
                                  id = tbClient.ClientID,
                                  text = tbClient.ClientAbbreviation,
                              }).ToList();
            listClient = Common.Tools.SetSelectJson(listClient);
            return Json(listClient, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询本公司名称下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectMessage()
        {
            var listMessage = (from tbMessage in myModels.SYS_Message
                               select new SelectVo
                               {
                                   id = tbMessage.MessageID,
                                   text = tbMessage.ChineseAbbreviation,
                               }).ToList();
            listMessage = Common.Tools.SetSelectJson(listMessage);
            return Json(listMessage, JsonRequestBehavior.AllowGet);
        }

        



        #endregion

        #region 打印数据
        /// <summary>
        /// 商务审核上方表格打印页面
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievement(string CarriageNumber, string Cupboard, string Seal, string IndentNumber, string ArriveFactoryTime, string LeftFactoryTime, string ClientAbbreviation, string EtrustNmber, int? TiGuiDiDianID, int? HuaiGuiDiDianID)
        {
            #region 数据查询
            var listExamine = from tbExamine in myModels.SYS_Commercel
                              join tbEtrust in myModels.SYS_Etrust on tbExamine.EtrustID equals tbEtrust.EtrustID
                              join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              join tbUser in myModels.SYS_User on tbExamine.CommerceQiExaminePID equals tbUser.UserID into From
                              from Add in From.DefaultIfEmpty()
                              join tbUser in myModels.SYS_User on tbExamine.CommerceExaminePID equals tbUser.UserID into From1
                              from Add1 in From1.DefaultIfEmpty()
                              where tbEtrust.AuditType.Trim() != "未审核"
                              orderby tbExamine.CommerceID descending
                              select new 
                              {
                                  EtrustID = tbExamine.EtrustID,//委托单号
                                  CommerceID = tbExamine.CommerceID,//审核ID
                                  EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                  CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                  EtrustType = tbEtrust.EtrustType,//状态
                                  ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                  CheXing = tbEtrust.CheXing,//车型
                                  CommerceExaminePID = tbExamine.CommerceExaminePID,//商务审核人 
                                  CommerceQiExaminePID = tbExamine.CommerceQiExaminePID,//商务弃审人
                                  FinancingExaninePID = tbExamine.FinancingExaninePID,//财务审核人
                                  FinancingQiExaninePID = tbExamine.FinancingQiExaninePID,//财务弃审人
                                  ShangWuQiShenTime = tbExamine.CommerceQiExamineTime.ToString(),//商务弃审时间
                                  FinancingQiExanineTime = tbExamine.FinancingQiExanineTime,//财务弃审时间
                                  ShangWuShenHeTime = tbExamine.CommerceExamineTime.ToString(),//商务审核时间
                                  FinancingExanineTime = tbExamine.FinancingExanineTime,//财务审核时间
                                  AuditType = tbEtrust.AuditType,//审核状态
                                  User = Add.AccountName,//商务弃审人 
                                  User1 = Add1.AccountName,//商务审核人
                                  Cupboard = tbEtrust.Cupboard,//柜号
                                  Seal = tbEtrust.Seal,//封条号
                                  IndentNumber = tbEtrust.IndentNumber,//订单号
                                  ArriveFactoryTime1 = tbEtrust.ArriveFactoryTime.ToString(),//到达工厂时间
                                  ArriveFactoryTime = tbEtrust.ArriveFactoryTime.ToString(),
                                  LeftFactoryTime1 = tbEtrust.LeftFactoryTime.ToString(),//离开工厂时间
                                  LeftFactoryTime = tbEtrust.LeftFactoryTime.ToString(),
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                              };
            if (!string.IsNullOrEmpty(CarriageNumber))
            {
                CarriageNumber = CarriageNumber.Trim();
                listExamine = listExamine.Where(p => p.CarriageNumber.Contains(CarriageNumber));
            }
            if (!string.IsNullOrEmpty(Cupboard))
            {
                Cupboard = Cupboard.Trim();
                listExamine = listExamine.Where(p => p.Cupboard.Contains(Cupboard));
            }
            if (!string.IsNullOrEmpty(Seal))
            {
                Seal = Seal.Trim();
                listExamine = listExamine.Where(p => p.Seal.Contains(Seal));
            }
            if (!string.IsNullOrEmpty(IndentNumber))
            {
                IndentNumber = IndentNumber.Trim();
                listExamine = listExamine.Where(p => p.IndentNumber.Contains(IndentNumber));
            }
            if (!string.IsNullOrEmpty(ArriveFactoryTime))
            {
                DateTime ArriveFactoryTime1 = Convert.ToDateTime(ArriveFactoryTime);
                listExamine = listExamine.Where(p => p.ArriveFactoryTime1.Contains(ArriveFactoryTime));
            }
            if (!string.IsNullOrEmpty(LeftFactoryTime))
            {
                DateTime LeftFactoryTime1 = Convert.ToDateTime(LeftFactoryTime);
                listExamine = listExamine.Where(p => p.LeftFactoryTime1.Contains(LeftFactoryTime));
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listExamine = listExamine.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation));
            }
            if (!string.IsNullOrEmpty(EtrustNmber))
            {
                EtrustNmber = EtrustNmber.Trim();
                listExamine = listExamine.Where(p => p.EtrustNmber.Contains(EtrustNmber));
            }
            if (TiGuiDiDianID > 0)
            {
                listExamine = listExamine.Where(m => m.TiGuiDiDianID == TiGuiDiDianID);
            }
            if (HuaiGuiDiDianID > 0)
            {
                listExamine = listExamine.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID);
            }

            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listExamine);

            //1、实例化数据集
            PrintReport.ExamineDB dbReport = new PrintReport.ExamineDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabExamine"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.Examine rp = new PrintReport.Examine();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\Commerce\\PrintReport\\Examine.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 商务审核下方表格打印页面
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementt(int? EtrustID)
        {
            #region 数据查询
            var listCharge = (from tbCharge in myModels.SYS_Charge
                              join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                              join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                              join tbOfferDetail in myModels.SYS_OfferDetail on tbEtrust.OfferDetailID equals tbOfferDetail.OfferDetailID
                              orderby tbCharge.ChargeID descending
                              select new ExamineVo
                              {
                                  EtrustID = tbEtrust.EtrustID,
                                  ChargeID = tbCharge.ChargeID,
                                  ReckoningUnit = "",
                                  ExpenseName = tbExpense.ExpenseName,
                                  SettleAccounts = tbExpense.SettleAccounts.Trim(),
                                  UnitPrice = tbCharge.UnitPrice,
                                  ChargeType = "正常",  //直接在查出来的表格给默认值，数据库不需要存在的字段
                                  BoxQuantity = tbOfferDetail.BoxQuantity,
                                  Money = tbOfferDetail.Money,
                                  Currency = tbOfferDetail.Currency,
                                  YingShouZ = 0,  //这是直接给用来计算的文本框的字段，数据库是没有的
                                  YingFuZ = 0,
                                  LiRun = 0,
                                  Bgs = tbCharge.ReckoningUnit, //本公司的结算代码
                                  Kh = tbCharge.ReckoningUnits  //客户公司的结算代码
                              }).ToList();

            if (EtrustID > 0)
            {
                listCharge = listCharge.Where(m => m.EtrustID == EtrustID).ToList();
            }
            if (listCharge.Count() > 0)  // Count 条数
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    if (!string.IsNullOrEmpty(listCharge[i].Bgs))
                    {
                        var B = listCharge[i].Bgs.Trim();  //TissueCode 本公司的组织代码  ChineseAbbreviation  客户简称
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == B).Select(m => m.ChineseAbbreviation).Single();
                        listCharge[i].ReckoningUnit = C;
                    }
                    else if (!string.IsNullOrEmpty(listCharge[i].Kh))
                    {
                        var B = listCharge[i].Kh.Trim();
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Select(m => m.ClientAbbreviation).Single();
                        listCharge[i].ReckoningUnit = C;
                    }
                }
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listCharge);

            //1、实例化数据集
            PrintReport.ChargeDB dbReport = new PrintReport.ChargeDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabCharge"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.Charge rp = new PrintReport.Charge();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\Commerce\\PrintReport\\Charge.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 新增结算单打印水晶报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementI(string EtrustNmber, string Cupboard, string Seal, string HuaiWeightTime, string ClientAbbreviation, int? TiGuiDiDianID, int? HuaiGuiDiDianID, int? ClientID, int? ShipID, string HangCi, int? PortID, int? GoalHarborID, int? UndertakeID, string EtrustType, int? ChauffeurID, int? VehicleInformationID, int? BracketID, string BookingSpace, string IndentNumber, string WorkCategory, string CheXing, string ReckoningUnit, int? ExpenseID, string SettleAccounts)
        {
            #region 数据查询
            var listCharge = (from tbCharge in myModels.SYS_Charge
                              join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                              join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                              join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              join tbVehicle in myModels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicle.VehicleInformationID
                              join tbSign in myModels.SYS_CloseAccount on tbCharge.ChargeID equals tbSign.ChargeID
                              where tbSign.SignID == null
                              orderby tbCharge.ChargeID descending
                              select new ExamineVo
                              {
                                  ChargeID = tbCharge.ChargeID, //收费ID
                                  EtrustID = tbEtrust.EtrustID,//委托单ID
                                  ExpenseName = tbExpense.ExpenseName, //费用名称
                                  ExpenseID = tbExpense.ExpenseID,//费用项目ID
                                  SettleAccounts = tbExpense.SettleAccounts,//收付款类型
                                  UnitPrice = tbCharge.UnitPrice,//单价
                                  Parities = "1",  //汇率
                                  Currency = "RMB",//币种
                                  YiJieMoney = "0",//已结金额
                                  CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                  EtrustType = tbEtrust.EtrustType,//委托单状态
                                  ClientAbbreviation = "",//客户简称
                                  CheXing = tbEtrust.CheXing,//车型
                                  WorkCategory = tbEtrust.WorkCategory,//作业类别
                                  EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                  Cupboard = tbEtrust.Cupboard,//柜号
                                  Seal = tbEtrust.Seal,//封条号
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                                  IndentNumber = tbEtrust.IndentNumber,//订单号
                                  HuaiWeightTime1 = tbEtrust.HuaiWeightTime.ToString(),//还柜时间
                                  HangCi = tbEtrust.HangCi,//航次
                                  ClientID = tbClient.ClientID,//客户信息(船公司)
                                  ShipID = tbEtrust.ShipID,//船名称
                                  PortID = tbEtrust.PortID,//进出口岸
                                  GoalHarborID = tbEtrust.GoalHarborID,//目的港
                                  UndertakeID = tbEtrust.UndertakeID,//承运公司
                                  ChauffeurID = tbVehicle.ChauffeurID,//司机资料ID
                                  VehicleInformationID = tbEtrust.VehicleInformationID,//车辆信息ID
                                  PlateNumbers = tbVehicle.PlateNumbers,//车牌号
                                  BracketID = tbVehicle.BracketID,//托架资料ID
                                  BookingSpace = tbEtrust.BookingSpace,//订舱号
                                  ReckoningUnit = "", //结算单位
                                  Bgs = tbCharge.ReckoningUnit.Trim(),//结算单位(本公司)
                                  Kh = tbCharge.ReckoningUnits.Trim(), //结算单位(客户)
                                  SignID = tbSign.SignID,//对账ID
                                  SignCloseAccountID = tbSign.SignCloseAccountID,//对账结算ID
                              }).ToList();

            if (listCharge.Count() > 0)  // Count 条数
            {
                var A = Convert.ToInt32(Session["SignCloseAccountID"].ToString()); //接收跳转页面时传过来的对账结算ID 
                var B = myModels.SYS_CloseAccount.Where(m => m.SignCloseAccountID == A).Select(m => m.ChargeID).ToList(); //让结算表里面的对账结算ID等于传过来的对账结算ID，然后查询出收费ID 
                if (B.Count() > 0)
                {
                    var C = B[0]; //声明一个变量接收上面查询出来的收费ID
                    var D = myModels.SYS_Charge.Where(m => m.ChargeID == C).ToList(); //声明一个变量接收收费表的收费ID等于C(前面声明一个变量接收上面查询出来的收费ID) 查询出来的收费数据，(一一对应的收费ID的那条数据查询出来)
                    var E = D[0].ReckoningUnit;  //声明一个变量接收收费ID对应的结算单位(本公司)
                    var F = D[0].ReckoningUnits; //声明一个变量接收收费ID对应的结算单位(客户)
                    if (!string.IsNullOrEmpty(E))
                    {
                        listCharge = listCharge.Where(m => m.Bgs == E.Trim()).ToList();  //本公司
                    }
                    else if (!string.IsNullOrEmpty(F))
                    {
                        listCharge = listCharge.Where(m => m.Kh == F.Trim()).ToList();  //客户
                    }
                }
            }
            if (listCharge.Count() > 0)  //两列数据放在表格的同一列显示出来写的判断，前面查询数据时，显示这一列的字段要让它为空，然后进行下面判断进行显示
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    var A = listCharge[i].Bgs; //本公司   Bgs前面表格查询需要显示出来的字段
                    var B = listCharge[i].Kh; //客户
                    if (!string.IsNullOrEmpty(A))
                    {
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == A).Single(); //查出本公司表里面的组织代码等于前面定义的变量查出那一条数据
                        listCharge[i].ClientAbbreviation = C.ChineseAbbreviation; //客户简称
                        listCharge[i].ReckoningUnit = C.ChineseName; //结算单位
                    }
                    else if (!string.IsNullOrEmpty(B))
                    {
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Single();
                        listCharge[i].ClientAbbreviation = C.ClientAbbreviation;
                        listCharge[i].ReckoningUnit = C.ChineseName;
                    }
                }
            }
            if (!string.IsNullOrEmpty(EtrustNmber))//委托单号
            {
                EtrustNmber = EtrustNmber.Trim();
                listCharge = listCharge.Where(p => p.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))//柜号
            {
                Cupboard = Cupboard.Trim();
                listCharge = listCharge.Where(p => p.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))//封条号
            {
                Seal = Seal.Trim();
                listCharge = listCharge.Where(p => p.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(HuaiWeightTime))//还柜时间
            {
                DateTime HuaiWeightTime1 = Convert.ToDateTime(HuaiWeightTime);
                listCharge = listCharge.Where(p => p.HuaiWeightTime1.Contains(HuaiWeightTime)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listCharge = listCharge.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (TiGuiDiDianID > 0)//提柜地点
            {
                listCharge = listCharge.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)//还柜地点
            {
                listCharge = listCharge.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (ClientID > 0)//船公司
            {
                listCharge = listCharge.Where(m => m.ClientID == ClientID).ToList();
            }
            if (ShipID > 0)//船名
            {
                listCharge = listCharge.Where(m => m.ShipID == ShipID).ToList();
            }
            if (!string.IsNullOrEmpty(HangCi))//航次
            {
                HangCi = HangCi.Trim();
                listCharge = listCharge.Where(p => p.HangCi.Contains(HangCi)).ToList();
            }
            if (PortID > 0)//进出口岸
            {
                listCharge = listCharge.Where(m => m.PortID == PortID).ToList();
            }
            if (GoalHarborID > 0)//目的港
            {
                listCharge = listCharge.Where(m => m.GoalHarborID == GoalHarborID).ToList();
            }
            if (UndertakeID > 0)//承运公司
            {
                listCharge = listCharge.Where(m => m.UndertakeID == UndertakeID).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustType))//状态
            {
                EtrustType = EtrustType.Trim();
                listCharge = listCharge.Where(p => p.EtrustType.Contains(EtrustType)).ToList();
            }
            if (ChauffeurID > 0)//司机
            {
                listCharge = listCharge.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (VehicleInformationID > 0)//车牌号
            {
                listCharge = listCharge.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (BracketID > 0)//托架编号
            {
                listCharge = listCharge.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(BookingSpace))//订舱号
            {
                BookingSpace = BookingSpace.Trim();
                listCharge = listCharge.Where(p => p.BookingSpace.Contains(BookingSpace)).ToList();
            }
            if (!string.IsNullOrEmpty(IndentNumber))//订单号
            {
                IndentNumber = IndentNumber.Trim();
                listCharge = listCharge.Where(p => p.IndentNumber.Contains(IndentNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(WorkCategory))//作业类别
            {
                WorkCategory = WorkCategory.Trim();
                listCharge = listCharge.Where(p => p.WorkCategory.Contains(WorkCategory)).ToList();
            }
            if (!string.IsNullOrEmpty(CheXing))//车型
            {
                CheXing = CheXing.Trim();
                listCharge = listCharge.Where(p => p.CheXing.Contains(CheXing)).ToList();
            }
            if (!string.IsNullOrEmpty(ReckoningUnit))//结算单位
            {
                ReckoningUnit = ReckoningUnit.Trim();
                listCharge = listCharge.Where(p => p.ReckoningUnit.Contains(ReckoningUnit)).ToList();
            }
            if (ExpenseID > 0)//费用名称
            {
                listCharge = listCharge.Where(m => m.ExpenseID == ExpenseID).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))//收付款类型
            {
                SettleAccounts = SettleAccounts.Trim();
                listCharge = listCharge.Where(p => p.SettleAccounts.Contains(SettleAccounts)).ToList();
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listCharge);

            //1、实例化数据集
            PrintReport.InsertDB dbReport = new PrintReport.InsertDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabInsert"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.Insert rp = new PrintReport.Insert();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\Commerce\\PrintReport\\Insert.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 对账成功后打印
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementII(int? ChargeID, string myArray)
        {
            #region 数据查询
            var listCharge = (from tbCharge in myModels.SYS_Charge
                              join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                              join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                              join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              join tbVehicle in myModels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicle.VehicleInformationID
                              join tbSign in myModels.SYS_CloseAccount on tbCharge.ChargeID equals tbSign.ChargeID
                              where tbSign.SignID == null
                              orderby tbCharge.ChargeID descending
                              select new ExamineVo
                              {
                                  ChargeID = tbCharge.ChargeID, //收费ID
                                  EtrustID = tbEtrust.EtrustID,//委托单ID
                                  ExpenseName = tbExpense.ExpenseName, //费用名称
                                  ExpenseID = tbExpense.ExpenseID,//费用项目ID
                                  SettleAccounts = tbExpense.SettleAccounts,//收付款类型
                                  UnitPrice = tbCharge.UnitPrice,//单价
                                  Parities = "1",  //汇率
                                  Currency = "RMB",//币种
                                  YiJieMoney = "0",//已结金额
                                  CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                  EtrustType = tbEtrust.EtrustType,//委托单状态
                                  ClientAbbreviation = "",//客户简称
                                  CheXing = tbEtrust.CheXing,//车型
                                  WorkCategory = tbEtrust.WorkCategory,//作业类别
                                  EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                  Cupboard = tbEtrust.Cupboard,//柜号
                                  Seal = tbEtrust.Seal,//封条号
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                                  IndentNumber = tbEtrust.IndentNumber,//订单号
                                  HuaiWeightTime1 = tbEtrust.HuaiWeightTime.ToString(),//还柜时间
                                  HangCi = tbEtrust.HangCi,//航次
                                  ClientID = tbClient.ClientID,//客户信息(船公司)
                                  ShipID = tbEtrust.ShipID,//船名称
                                  PortID = tbEtrust.PortID,//进出口岸
                                  GoalHarborID = tbEtrust.GoalHarborID,//目的港
                                  UndertakeID = tbEtrust.UndertakeID,//承运公司
                                  ChauffeurID = tbVehicle.ChauffeurID,//司机资料ID
                                  VehicleInformationID = tbEtrust.VehicleInformationID,//车辆信息ID
                                  PlateNumbers = tbVehicle.PlateNumbers,//车牌号
                                  BracketID = tbVehicle.BracketID,//托架资料ID
                                  BookingSpace = tbEtrust.BookingSpace,//订舱号
                                  ReckoningUnit = "", //结算单位
                                  Bgs = tbCharge.ReckoningUnit.Trim(),//结算单位(本公司)
                                  Kh = tbCharge.ReckoningUnits.Trim(), //结算单位(客户)
                                  SignID = tbSign.SignID,//对账ID
                                  SignCloseAccountID = tbSign.SignCloseAccountID,//对账结算ID
                                  CabinetType = tbEtrust.CabinetType,//箱型
                              }).ToList();

            if (listCharge.Count() > 0)  //两列数据放在表格的同一列显示出来写的判断，前面查询数据时，显示这一列的字段要让它为空，然后进行下面判断进行显示
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    var A = listCharge[i].Bgs; //本公司   Bgs前面表格查询需要显示出来的字段
                    var B = listCharge[i].Kh; //客户
                    if (!string.IsNullOrEmpty(A)) //字符串不为空
                    {
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == A).Single(); //查出本公司表里面的组织代码等于前面定义的变量查出那一条数据
                        listCharge[i].ClientAbbreviation = C.ChineseAbbreviation; //客户简称
                        listCharge[i].ReckoningUnit = C.ChineseName; //结算单位
                    }
                    else if (!string.IsNullOrEmpty(B))
                    {
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Single();
                        listCharge[i].ClientAbbreviation = C.ClientAbbreviation;
                        listCharge[i].ReckoningUnit = C.ChineseName;
                    }
                }
            }
            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    var F = listCharge[i].ChargeID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listCharge.Remove(listCharge[i]);
                        i = 0;
                    }
                }
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listCharge);

            //1、实例化数据集
            PrintReport.IndexDB dbReport = new PrintReport.IndexDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabIndex"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.Index rp = new PrintReport.Index();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\Commerce\\PrintReport\\Index.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 财务结算右边表格打印水晶报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementC(int? SignCloseAccountID, string myArray)
        {
            #region 数据查询

            var listSign = (from tbSign in myModels.SYS_Sign
                           join tbClose in myModels.SYS_SignCloseAccount on tbSign.SignCloseAccountID equals tbClose.SignCloseAccountID
                           join tbPay in myModels.SYS_PayWay on tbSign.PayWayID equals tbPay.PayWayID
                           join tbSettle in myModels.SYS_SettleWay on tbSign.SettleWayID equals tbSettle.SettleWayID
                           where tbSign.FinishType.Trim() == "未核销"
                           orderby tbSign.SignCloseAccountID descending
                           select new ExamineVo
                           {
                               SignCloseAccountID = tbClose.SignCloseAccountID,//对账结算ID
                               SignID = tbSign.SignID,//标记对账ID
                               SignType = tbSign.SignType.Trim(),//对账状态   需要用来筛选的字段在查询时最好要去空格，要不然后面有可能出错
                               SettleNumber = tbSign.SettleNumber,//结算单号
                               PayWay = tbPay.PayWay,//付款方式
                               SettleWay = tbSettle.SettleWay,//结算方式
                               YuSettleDate1 = tbSign.YuSettleDate.ToString(),//预结款日期
                               FinishType = tbSign.FinishType, //结算状态(已核销，未核销)
                               SignExplain = tbSign.SignExplain,//对账说明
                               SignDate = tbSign.SignDate.ToString(),//对账日期
                               SignDate1 = tbSign.SignDate.ToString(),//转换对账日期
                               PayWayID = tbSign.PayWayID,//付款方式ID
                               SettleWayID = tbSign.SettleWayID,//结算方式ID
                               SettleMoney = tbSign.SettleMoney,//结算金额
                           }).ToList();

            if (SignCloseAccountID > 0)
            {
                listSign = listSign.Where(m => m.SignCloseAccountID == SignCloseAccountID).ToList();
            }
            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listSign.Count(); i++)
                {
                    var F = listSign[i].SignCloseAccountID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listSign.Remove(listSign[i]);
                        i = 0;
                    }
                }
            }

            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listSign);

            //1、实例化数据集
            PrintReport.SignDB dbReport = new PrintReport.SignDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabSign"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.Sign rp = new PrintReport.Sign();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\Commerce\\PrintReport\\Sign.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 实收实付上方表格打印代码
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintShiShouShiFu(string CalculateNumber, int? ClientID, int? MessageID, string BankName, string SFDate, int? PayWayID, string Currency, string SFMoney, string AlreadyCancelMoney, string StayCancelMoney)
        {
            #region 查询数据
            var listCalculate = (from tbCalculate in myModels.SYS_Calculate
                                 join tbPayWay in myModels.SYS_PayWay on tbCalculate.PayWayID equals tbPayWay.PayWayID
                                 orderby tbCalculate.CalculateID descending
                                 select new ExamineVo
                                 {
                                     CalculateID = tbCalculate.CalculateID,//计费ID
                                     SignID = tbCalculate.SignID,//对账ID
                                     CalculateState = "已审核",//计费状态
                                     CalculateNumber = tbCalculate.CalculateNumber.Trim(),//收付款单号
                                     SettleAccounts = tbCalculate.SettleAccounts.Trim(),//收付款类型
                                     SFMoney = tbCalculate.SFMoney,//收付金额
                                     AlreadyCancelMoney = tbCalculate.AlreadyCancelMoney,//已核销金额
                                     StayCancelMoney = tbCalculate.StayCancelMoney,//待核销金额
                                     Currency = tbCalculate.Currency.Trim(),//币种
                                     SFDate1 = tbCalculate.SFDate.ToString(),//收付日期
                                     BankName = tbCalculate.BankName.Trim(),//银行名称
                                     BankAccount = tbCalculate.BankAccount.Trim(),//银行账号
                                     PayWay = tbPayWay.PayWay.Trim(),//付款方式
                                     ChequeNumber = tbCalculate.ChequeNumber.Trim(),//支票号
                                     ClientID = tbCalculate.ClientID,//客户ID  结算单位多条件查询
                                     PayWayID = tbCalculate.PayWayID,//付款方式ID
                                     ReckoningUnit = "",
                                     MessageID = tbCalculate.MessageID,
                                 }).ToList();

            if (listCalculate.Count() > 0) // Count 条数 两列放在同一列显示，在上面声明一个空的字段在表格显示，然后判断查询该显示哪张表的名称
            {
                for (int i = 0; i < listCalculate.Count(); i++)
                {
                    if (listCalculate[i].MessageID > 0)
                    {
                        var B = listCalculate[i].MessageID;  //listCalculate[i] 前面查询要用ToList()
                        var C = myModels.SYS_Message.Where(m => m.MessageID == B).Select(m => m.ChineseAbbreviation).Single();
                        listCalculate[i].ReckoningUnit = C;
                    }
                    else if (listCalculate[i].ClientID > 0)
                    {
                        var B = listCalculate[i].ClientID;
                        var C = myModels.SYS_Client.Where(m => m.ClientID == B).Select(m => m.ClientAbbreviation).Single();
                        listCalculate[i].ReckoningUnit = C;
                    }
                }
            }


            if (!string.IsNullOrEmpty(CalculateNumber))//收付款单号
            {
                CalculateNumber = CalculateNumber.Trim();
                listCalculate = listCalculate.Where(p => p.CalculateNumber.Contains(CalculateNumber)).ToList();
            }
            if (ClientID > 0)//结算单位
            {
                listCalculate = listCalculate.Where(m => m.ClientID == ClientID).ToList();
            }
            if (MessageID > 0)//结算单位
            {
                listCalculate = listCalculate.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(BankName))//银行名称
            {
                BankName = BankName.Trim();
                listCalculate = listCalculate.Where(p => p.BankName.Contains(BankName)).ToList();
            }
            if (!string.IsNullOrEmpty(SFDate))//收付日期
            {
                DateTime SFDate1 = Convert.ToDateTime(SFDate);
                listCalculate = listCalculate.Where(p => p.SFDate1.Contains(SFDate)).ToList();
            }
            if (PayWayID > 0)//付款方式ID
            {
                listCalculate = listCalculate.Where(m => m.PayWayID == PayWayID).ToList();
            }
            if (!string.IsNullOrEmpty(Currency))//币种
            {
                Currency = Currency.Trim();
                listCalculate = listCalculate.Where(p => p.Currency.Contains(Currency)).ToList();
            }
            if (!string.IsNullOrEmpty(SFMoney))//收付金额
            {
                SFMoney = SFMoney.Trim();
                listCalculate = listCalculate.Where(p => p.SFMoney.ToString().Contains(SFMoney)).ToList();
            }
            if (!string.IsNullOrEmpty(AlreadyCancelMoney))//已核销金额
            {
                AlreadyCancelMoney = AlreadyCancelMoney.Trim();
                listCalculate = listCalculate.Where(p => p.AlreadyCancelMoney.ToString().Contains(AlreadyCancelMoney)).ToList();
            }
            if (!string.IsNullOrEmpty(StayCancelMoney))//待核销金额
            {
                StayCancelMoney = StayCancelMoney.Trim();
                listCalculate = listCalculate.Where(p => p.StayCancelMoney.ToString().Contains(StayCancelMoney)).ToList();
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listCalculate);

            //1、实例化数据集
            PrintReport.ShiShouShiFuDB dbReport = new PrintReport.ShiShouShiFuDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabShiShouShiFu"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.ShiShouShiFu rp = new PrintReport.ShiShouShiFu();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\Commerce\\PrintReport\\ShiShouShiFu.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }




        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable的对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null)
                return dtReturn;//安全性检查
            //循环遍历集合，使用反射获取类型的属性信息
            foreach (T rec in varlist)
            {
                //使用反射获取T类型的属性信息，返回一个PropertyInfo类型的集合
                #region 
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    //循环PropertyInfo数组
                    foreach (PropertyInfo pi in oProps)
                    {
                        //得到属性的类型
                        Type colType = pi.PropertyType;
                        //如果属性为泛型类型
                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            //获取泛型类型的参数
                            colType = colType.GetGenericArguments()[0];
                        }
                        //将类型的属性名称与属性类型作为DataTable 的列数据
                        dtReturn.Columns.Add(pi.Name, colType);
                    }
                }
                #endregion
                //新建一个用于添加到DataTable中的DataRow对象
                DataRow dr = dtReturn.NewRow();
                //循环遍历属性集合
                foreach (PropertyInfo pi in oProps)
                {
                    //为DataRow中的指定列赋值
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue(rec, null);
                }
                //将具有结果值的DataRow添加到DataTable集合中
                dtReturn.Rows.Add(dr);
            }
            return dtReturn;//返回DataTable对象
        }
        #endregion

        #region 导出数据
        /// <summary>
        /// 上方表格导出数据
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportYuDingInfor(string CarriageNumber, string Cupboard, string Seal, string IndentNumber, string ArriveFactoryTime, string LeftFactoryTime, string ClientAbbreviation, string EtrustNmber, int? TiGuiDiDianID, int? HuaiGuiDiDianID)
        {
            //查询分页数据
            List<ExamineVo> listExamine = (from tbExamine in myModels.SYS_Commercel
                              join tbEtrust in myModels.SYS_Etrust on tbExamine.EtrustID equals tbEtrust.EtrustID
                              join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              join tbUser in myModels.SYS_User on tbExamine.CommerceQiExaminePID equals tbUser.UserID into From
                              from Add in From.DefaultIfEmpty()
                              join tbUser in myModels.SYS_User on tbExamine.CommerceExaminePID equals tbUser.UserID into From1
                              from Add1 in From1.DefaultIfEmpty()
                              where tbEtrust.AuditType.Trim() != "未审核"
                              orderby tbExamine.CommerceID descending
                              select new ExamineVo
                              {
                                  EtrustID = tbExamine.EtrustID,//委托单号
                                  CommerceID = tbExamine.CommerceID,//审核ID
                                  EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                  CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                  EtrustType = tbEtrust.EtrustType,//状态
                                  ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                  CheXing = tbEtrust.CheXing,//车型
                                  CommerceExaminePID = tbExamine.CommerceExaminePID,//商务审核人 
                                  CommerceQiExaminePID = tbExamine.CommerceQiExaminePID,//商务弃审人
                                  FinancingExaninePID = tbExamine.FinancingExaninePID,//财务审核人
                                  FinancingQiExaninePID = tbExamine.FinancingQiExaninePID,//财务弃审人
                                  ShangWuQiShenTime = tbExamine.CommerceQiExamineTime.ToString(),//商务弃审时间
                                  FinancingQiExanineTime = tbExamine.FinancingQiExanineTime,//财务弃审时间
                                  ShangWuShenHeTime = tbExamine.CommerceExamineTime.ToString(),//商务审核时间
                                  FinancingExanineTime = tbExamine.FinancingExanineTime,//财务审核时间
                                  AuditType = tbEtrust.AuditType,//审核状态
                                  User = Add.AccountName,//商务弃审人 
                                  User1 = Add1.AccountName,//商务审核人
                                  Cupboard = tbEtrust.Cupboard,//柜号
                                  Seal = tbEtrust.Seal,//封条号
                                  IndentNumber = tbEtrust.IndentNumber,//订单号
                                  ArriveFactoryTime1 = tbEtrust.ArriveFactoryTime.ToString(),//到达工厂时间
                                  //ArriveFactoryTime = tbEtrust.ArriveFactoryTime.ToString(),
                                  LeftFactoryTime1 = tbEtrust.LeftFactoryTime.ToString(),//离开工厂时间
                                  LeftFactoryTime = tbEtrust.LeftFactoryTime.ToString(),
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                              }).ToList();
            if (!string.IsNullOrEmpty(CarriageNumber))
            {
                CarriageNumber = CarriageNumber.Trim();
                listExamine = listExamine.Where(p => p.CarriageNumber.Contains(CarriageNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))
            {
                Cupboard = Cupboard.Trim();
                listExamine = listExamine.Where(p => p.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))
            {
                Seal = Seal.Trim();
                listExamine = listExamine.Where(p => p.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(IndentNumber))
            {
                IndentNumber = IndentNumber.Trim();
                listExamine = listExamine.Where(p => p.IndentNumber.Contains(IndentNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(ArriveFactoryTime))
            {
                DateTime ArriveFactoryTime1 = Convert.ToDateTime(ArriveFactoryTime);
                listExamine = listExamine.Where(p => p.ArriveFactoryTime1.Contains(ArriveFactoryTime)).ToList();
            }
            if (!string.IsNullOrEmpty(LeftFactoryTime))
            {
                DateTime LeftFactoryTime1 = Convert.ToDateTime(LeftFactoryTime);
                listExamine = listExamine.Where(p => p.LeftFactoryTime1.Contains(LeftFactoryTime)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listExamine = listExamine.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustNmber))
            {
                EtrustNmber = EtrustNmber.Trim();
                listExamine = listExamine.Where(p => p.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (TiGuiDiDianID > 0)
            {
                listExamine = listExamine.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)
            {
                listExamine = listExamine.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("委托单号");
            row1.CreateCell(1).SetCellValue("运输单号");
            row1.CreateCell(2).SetCellValue("状态");
            row1.CreateCell(3).SetCellValue("客户简称");
            row1.CreateCell(4).SetCellValue("柜(车)型");
            row1.CreateCell(5).SetCellValue("柜号");
            row1.CreateCell(6).SetCellValue("封条号");
            row1.CreateCell(7).SetCellValue("订单号");
            row1.CreateCell(8).SetCellValue("到厂时间");
            row1.CreateCell(9).SetCellValue("离厂时间");
            row1.CreateCell(10).SetCellValue("商务审核人");
            row1.CreateCell(11).SetCellValue("商务弃审人");
            row1.CreateCell(12).SetCellValue("财务审核人");
            row1.CreateCell(13).SetCellValue("财务弃审人");
            row1.CreateCell(14).SetCellValue("商务审核时间");
            row1.CreateCell(15).SetCellValue("商务弃审时间");
            row1.CreateCell(16).SetCellValue("财务审核时间");
            row1.CreateCell(17).SetCellValue("财务弃审时间");
            row1.CreateCell(18).SetCellValue("审核否");
            //4、循环写入数据
            for (var i = 0; i < listExamine.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExamine[i].EtrustNmber);
                rowTemp.CreateCell(1).SetCellValue(listExamine[i].CarriageNumber);
                rowTemp.CreateCell(2).SetCellValue(listExamine[i].EtrustType);
                rowTemp.CreateCell(3).SetCellValue(listExamine[i].ClientAbbreviation);
                rowTemp.CreateCell(4).SetCellValue(listExamine[i].CheXing);
                rowTemp.CreateCell(5).SetCellValue(listExamine[i].Cupboard);
                rowTemp.CreateCell(6).SetCellValue(listExamine[i].Seal);
                rowTemp.CreateCell(7).SetCellValue(listExamine[i].IndentNumber);
                rowTemp.CreateCell(8).SetCellValue(listExamine[i].ArriveFactoryTime1);
                rowTemp.CreateCell(9).SetCellValue(listExamine[i].LeftFactoryTime1);
                rowTemp.CreateCell(10).SetCellValue(listExamine[i].User1);
                rowTemp.CreateCell(11).SetCellValue(listExamine[i].User);
                rowTemp.CreateCell(12).SetCellValue(listExamine[i].FinancingExaninePID.ToString());
                rowTemp.CreateCell(13).SetCellValue(listExamine[i].FinancingQiExaninePID.ToString());
                rowTemp.CreateCell(14).SetCellValue(listExamine[i].ShangWuShenHeTime);
                rowTemp.CreateCell(15).SetCellValue(listExamine[i].ShangWuQiShenTime);
                rowTemp.CreateCell(16).SetCellValue(listExamine[i].FinancingExanineTime.ToString());
                rowTemp.CreateCell(17).SetCellValue(listExamine[i].FinancingQiExanineTime.ToString());
                rowTemp.CreateCell(18).SetCellValue(listExamine[i].AuditType);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司审核情况报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);
            //6、文件名


            return File(exStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 下方表格导出数据
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportYuDingInforr(int? EtrustID)
        {
            var listCharge = (from tbCharge in myModels.SYS_Charge
                              join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                              join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                              join tbOfferDetail in myModels.SYS_OfferDetail on tbEtrust.OfferDetailID equals tbOfferDetail.OfferDetailID
                              orderby tbCharge.ChargeID descending
                              select new ExamineVo
                              {
                                  EtrustID = tbEtrust.EtrustID,
                                  ChargeID = tbCharge.ChargeID,
                                  ReckoningUnit = "",
                                  ExpenseName = tbExpense.ExpenseName,
                                  SettleAccounts = tbExpense.SettleAccounts.Trim(),
                                  UnitPrice = tbCharge.UnitPrice,
                                  ChargeType = "正常",  //直接在查出来的表格给默认值，数据库不需要存在的字段
                                  BoxQuantity = tbOfferDetail.BoxQuantity,
                                  Money = tbOfferDetail.Money,
                                  Currency = tbOfferDetail.Currency,
                                  YingShouZ = 0,  //这是直接给用来计算的文本框的字段，数据库是没有的
                                  YingFuZ = 0,
                                  LiRun = 0,
                                  Bgs = tbCharge.ReckoningUnit, //本公司的结算代码
                                  Kh = tbCharge.ReckoningUnits  //客户公司的结算代码
                              }).ToList();

            if (EtrustID > 0)
            {
                listCharge = listCharge.Where(m => m.EtrustID == EtrustID).ToList();
            }
            if (listCharge.Count() > 0)  // Count 条数
            {
                for (int i = 0; i < listCharge.Count(); i++)
                {
                    if (!string.IsNullOrEmpty(listCharge[i].Bgs))
                    {
                        var B = listCharge[i].Bgs.Trim();  //TissueCode 本公司的组织代码  ChineseAbbreviation  客户简称
                        var C = myModels.SYS_Message.Where(m => m.TissueCode.Trim() == B).Select(m => m.ChineseAbbreviation).Single();
                        listCharge[i].ReckoningUnit = C;
                    }
                    else if (!string.IsNullOrEmpty(listCharge[i].Kh))
                    {
                        var B = listCharge[i].Kh.Trim();
                        var C = myModels.SYS_Client.Where(m => m.ClientCode.Trim() == B).Select(m => m.ClientAbbreviation).Single();
                        listCharge[i].ReckoningUnit = C;
                    }
                }
            }
            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("结算单位");
            row1.CreateCell(1).SetCellValue("费用名称");
            row1.CreateCell(2).SetCellValue("费用类型");
            row1.CreateCell(3).SetCellValue("收付款类型");
            row1.CreateCell(4).SetCellValue("单价");
            row1.CreateCell(5).SetCellValue("箱量");
            row1.CreateCell(6).SetCellValue("金额");
            row1.CreateCell(7).SetCellValue("币种");
            //4、循环写入数据
            for (var i = 0; i < listCharge.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listCharge[i].ReckoningUnit);
                rowTemp.CreateCell(1).SetCellValue(listCharge[i].ExpenseName);
                rowTemp.CreateCell(2).SetCellValue(listCharge[i].ChargeType);
                rowTemp.CreateCell(3).SetCellValue(listCharge[i].SettleAccounts);
                rowTemp.CreateCell(4).SetCellValue(listCharge[i].UnitPrice.ToString());
                rowTemp.CreateCell(5).SetCellValue(listCharge[i].BoxQuantity);
                rowTemp.CreateCell(6).SetCellValue(listCharge[i].Money.ToString());
                rowTemp.CreateCell(7).SetCellValue(listCharge[i].Currency);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司审核收费情况报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);
            //6、文件名

            return File(exStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 导出新增结算单数据
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportInsert(string EtrustNmber, string Cupboard, string Seal, string HuaiWeightTime, string ClientAbbreviation, int? TiGuiDiDianID, int? HuaiGuiDiDianID, int? ClientID, int? ShipID, string HangCi, int? PortID, int? GoalHarborID, int? UndertakeID, string EtrustType, int? ChauffeurID, int? VehicleInformationID, int? BracketID, string BookingSpace, string IndentNumber, string WorkCategory, string CheXing, string ReckoningUnit, int? ExpenseID, string SettleAccounts)
        {
            List<ExamineVo> listCharge = (from tbCharge in myModels.SYS_Charge
                             join tbEtrust in myModels.SYS_Etrust on tbCharge.EtrustID equals tbEtrust.EtrustID
                             join tbExpense in myModels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                             join tbClient in myModels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                             join tbVehicle in myModels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicle.VehicleInformationID
                             orderby tbCharge.ChargeID descending
                             select new ExamineVo
                             {
                                 ChargeID = tbCharge.ChargeID, //收费ID
                                 EtrustID = tbEtrust.EtrustID,//委托单ID
                                 ExpenseName = tbExpense.ExpenseName, //费用名称
                                 ExpenseID = tbExpense.ExpenseID,//费用项目ID
                                 SettleAccounts = tbExpense.SettleAccounts,//收付款类型
                                 UnitPrice = tbCharge.UnitPrice,//单价
                                 Parities = "1",  //汇率
                                 Currency = "RMB",//币种
                                 YiJieMoney = "0",//已结金额
                                 CarriageNumber = tbEtrust.CarriageNumber,//运输单号
                                 EtrustType = tbEtrust.EtrustType,//委托单状态
                                 ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                 CheXing = tbEtrust.CheXing,//车型
                                 WorkCategory = tbEtrust.WorkCategory,//作业类别
                                 EtrustNmber = tbEtrust.EtrustNmber,//委托单号
                                 Cupboard = tbEtrust.Cupboard,//柜号
                                 Seal = tbEtrust.Seal,//封条号
                                 TiGuiDiDianID = tbEtrust.TiGuiDiDianID,//提柜地点
                                 HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,//还柜地点
                                 IndentNumber = tbEtrust.IndentNumber,//订单号
                                 HuaiWeightTime1 = tbEtrust.HuaiWeightTime.ToString(),//还柜时间
                                 HangCi = tbEtrust.HangCi,//航次
                                 ClientID = tbClient.ClientID,//客户信息(船公司)
                                 ShipID = tbEtrust.ShipID,//船名称
                                 PortID = tbEtrust.PortID,//进出口岸
                                 GoalHarborID = tbEtrust.GoalHarborID,//目的港
                                 UndertakeID = tbEtrust.UndertakeID,//承运公司
                                 ChauffeurID = tbVehicle.ChauffeurID,//司机资料ID
                                 VehicleInformationID = tbEtrust.VehicleInformationID,//车辆信息ID
                                 PlateNumbers = tbVehicle.PlateNumbers,//车牌号
                                 BracketID = tbVehicle.BracketID,//托架资料ID
                                 BookingSpace = tbEtrust.BookingSpace,//订舱号
                                 ReckoningUnit = tbCharge.ReckoningUnit,//结算单位
                             }).ToList();

            if (!string.IsNullOrEmpty(EtrustNmber))//委托单号
            {
                EtrustNmber = EtrustNmber.Trim();
                listCharge = listCharge.Where(p => p.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))//柜号
            {
                Cupboard = Cupboard.Trim();
                listCharge = listCharge.Where(p => p.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))//封条号
            {
                Seal = Seal.Trim();
                listCharge = listCharge.Where(p => p.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(HuaiWeightTime))//还柜时间
            {
                DateTime HuaiWeightTime1 = Convert.ToDateTime(HuaiWeightTime);
                listCharge = listCharge.Where(p => p.HuaiWeightTime1.Contains(HuaiWeightTime)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listCharge = listCharge.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (TiGuiDiDianID > 0)//提柜地点
            {
                listCharge = listCharge.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)//还柜地点
            {
                listCharge = listCharge.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (ClientID > 0)//船公司
            {
                listCharge = listCharge.Where(m => m.ClientID == ClientID).ToList();
            }
            if (ShipID > 0)//船名
            {
                listCharge = listCharge.Where(m => m.ShipID == ShipID).ToList();
            }
            if (!string.IsNullOrEmpty(HangCi))//航次
            {
                HangCi = HangCi.Trim();
                listCharge = listCharge.Where(p => p.HangCi.Contains(HangCi)).ToList();
            }
            if (PortID > 0)//进出口岸
            {
                listCharge = listCharge.Where(m => m.PortID == PortID).ToList();
            }
            if (GoalHarborID > 0)//目的港
            {
                listCharge = listCharge.Where(m => m.GoalHarborID == GoalHarborID).ToList();
            }
            if (UndertakeID > 0)//承运公司
            {
                listCharge = listCharge.Where(m => m.UndertakeID == UndertakeID).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustType))//状态
            {
                EtrustType = EtrustType.Trim();
                listCharge = listCharge.Where(p => p.EtrustType.Contains(EtrustType)).ToList();
            }
            if (ChauffeurID > 0)//司机
            {
                listCharge = listCharge.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (VehicleInformationID > 0)//车牌号
            {
                listCharge = listCharge.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (BracketID > 0)//托架编号
            {
                listCharge = listCharge.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(BookingSpace))//订舱号
            {
                BookingSpace = BookingSpace.Trim();
                listCharge = listCharge.Where(p => p.BookingSpace.Contains(BookingSpace)).ToList();
            }
            if (!string.IsNullOrEmpty(IndentNumber))//订单号
            {
                IndentNumber = IndentNumber.Trim();
                listCharge = listCharge.Where(p => p.IndentNumber.Contains(IndentNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(WorkCategory))//作业类别
            {
                WorkCategory = WorkCategory.Trim();
                listCharge = listCharge.Where(p => p.WorkCategory.Contains(WorkCategory)).ToList();
            }
            if (!string.IsNullOrEmpty(CheXing))//车型
            {
                CheXing = CheXing.Trim();
                listCharge = listCharge.Where(p => p.CheXing.Contains(CheXing)).ToList();
            }
            if (!string.IsNullOrEmpty(ReckoningUnit))//结算单位
            {
                ReckoningUnit = ReckoningUnit.Trim();
                listCharge = listCharge.Where(p => p.ReckoningUnit.Contains(ReckoningUnit)).ToList();
            }
            if (ExpenseID > 0)//费用名称
            {
                listCharge = listCharge.Where(m => m.ExpenseID == ExpenseID).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))//收付款类型
            {
                SettleAccounts = SettleAccounts.Trim();
                listCharge = listCharge.Where(p => p.SettleAccounts.Contains(SettleAccounts)).ToList();
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("费用名称");
            row1.CreateCell(1).SetCellValue("收付类型");
            row1.CreateCell(2).SetCellValue("金额");
            row1.CreateCell(3).SetCellValue("已结算");
            row1.CreateCell(4).SetCellValue("未结算");
            row1.CreateCell(5).SetCellValue("汇率");
            row1.CreateCell(6).SetCellValue("币种");
            row1.CreateCell(7).SetCellValue("订单号");
            row1.CreateCell(8).SetCellValue("委托单号");
            row1.CreateCell(9).SetCellValue("柜号");
            row1.CreateCell(10).SetCellValue("封条号");
            row1.CreateCell(11).SetCellValue("业务状态");
            row1.CreateCell(12).SetCellValue("客户简称");
            row1.CreateCell(13).SetCellValue("结算单位");
            row1.CreateCell(14).SetCellValue("柜(车)型");
            row1.CreateCell(15).SetCellValue("作业类别");
            row1.CreateCell(16).SetCellValue("还柜时间");
            //4、循环写入数据
            for (var i = 0; i < listCharge.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listCharge[i].ExpenseName);
                rowTemp.CreateCell(1).SetCellValue(listCharge[i].SettleAccounts);
                rowTemp.CreateCell(2).SetCellValue(listCharge[i].UnitPrice.ToString());
                rowTemp.CreateCell(3).SetCellValue(listCharge[i].YiJieMoney);
                rowTemp.CreateCell(4).SetCellValue(listCharge[i].UnitPrice.ToString());
                rowTemp.CreateCell(5).SetCellValue(listCharge[i].Parities);
                rowTemp.CreateCell(6).SetCellValue(listCharge[i].Currency);
                rowTemp.CreateCell(7).SetCellValue(listCharge[i].IndentNumber);
                rowTemp.CreateCell(8).SetCellValue(listCharge[i].EtrustNmber);
                rowTemp.CreateCell(9).SetCellValue(listCharge[i].Cupboard);
                rowTemp.CreateCell(10).SetCellValue(listCharge[i].Seal);
                rowTemp.CreateCell(11).SetCellValue(listCharge[i].EtrustType);
                rowTemp.CreateCell(12).SetCellValue(listCharge[i].ClientAbbreviation);
                rowTemp.CreateCell(13).SetCellValue(listCharge[i].ReckoningUnit);
                rowTemp.CreateCell(14).SetCellValue(listCharge[i].CheXing);
                rowTemp.CreateCell(15).SetCellValue(listCharge[i].WorkCategory);
                rowTemp.CreateCell(16).SetCellValue(listCharge[i].HuaiWeightTime1);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司新增结算单报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);
            //6、文件名
            return File(exStream, "application/vnd.ms-excel", fileName);
        }


        /// <summary>
        /// 财务结算右边表格导出EXCEL表格
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportSign(int? SignCloseAccountID, string myArray)
        {
            List<ExamineVo> listSign = (from tbSign in myModels.SYS_Sign
                           join tbClose in myModels.SYS_SignCloseAccount on tbSign.SignCloseAccountID equals tbClose.SignCloseAccountID
                           join tbPay in myModels.SYS_PayWay on tbSign.PayWayID equals tbPay.PayWayID
                           join tbSettle in myModels.SYS_SettleWay on tbSign.SettleWayID equals tbSettle.SettleWayID
                           where tbSign.FinishType.Trim() == "未核销"
                           orderby tbSign.SignCloseAccountID descending
                           select new ExamineVo
                           {
                               SignCloseAccountID = tbClose.SignCloseAccountID,//对账结算ID
                               SignID = tbSign.SignID,//标记对账ID
                               SignType = tbSign.SignType.Trim(),//对账状态   需要用来筛选的字段在查询时最好要去空格，要不然后面有可能出错
                               SettleNumber = tbSign.SettleNumber,//结算单号
                               PayWay = tbPay.PayWay,//付款方式
                               SettleWay = tbSettle.SettleWay,//结算方式
                               YuSettleDate1 = tbSign.YuSettleDate.ToString(),//预结款日期
                               FinishType = tbSign.FinishType, //结算状态(已核销，未核销)
                               SignExplain = tbSign.SignExplain,//对账说明
                               SignDate = tbSign.SignDate.ToString(),//对账日期
                               SignDate1 = tbSign.SignDate.ToString(),//转换对账日期
                               PayWayID = tbSign.PayWayID,//付款方式ID
                               SettleWayID = tbSign.SettleWayID,//结算方式ID
                               SettleMoney = tbSign.SettleMoney,//结算金额
                           }).ToList();

            if (SignCloseAccountID > 0)
            {
                listSign = listSign.Where(m => m.SignCloseAccountID == SignCloseAccountID).ToList();
            }
            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listSign.Count(); i++)
                {
                    var F = listSign[i].SignCloseAccountID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listSign.Remove(listSign[i]);
                        i = 0;
                    }
                }
            }


            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("对账状态");
            row1.CreateCell(1).SetCellValue("对账说明");
            row1.CreateCell(2).SetCellValue("结算单号");
            row1.CreateCell(3).SetCellValue("结算金额");
            row1.CreateCell(4).SetCellValue("付款方式");
            row1.CreateCell(5).SetCellValue("结算方式");
            row1.CreateCell(6).SetCellValue("对账日期");
            row1.CreateCell(7).SetCellValue("核销状态");
            //4、循环写入数据
            for (var i = 0; i < listSign.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listSign[i].SignType);
                rowTemp.CreateCell(1).SetCellValue(listSign[i].SignExplain);
                rowTemp.CreateCell(2).SetCellValue(listSign[i].SettleNumber);
                rowTemp.CreateCell(3).SetCellValue(listSign[i].SettleMoney.ToString());
                rowTemp.CreateCell(4).SetCellValue(listSign[i].PayWay);
                rowTemp.CreateCell(5).SetCellValue(listSign[i].SettleWay);
                rowTemp.CreateCell(6).SetCellValue(listSign[i].SignDate1);
                rowTemp.CreateCell(7).SetCellValue(listSign[i].FinishType);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司对账单报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);
            //6、文件名
            return File(exStream, "application/vnd.ms-excel", fileName);

        }

        /// <summary>
        /// 实收实付上方表格导出数据的保存方法
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportShiShouShiFu(string CalculateNumber, int? ClientID, int? MessageID, string BankName, string SFDate, int? PayWayID, string Currency, string SFMoney, string AlreadyCancelMoney, string StayCancelMoney)
        {
            var listCalculate = (from tbCalculate in myModels.SYS_Calculate
                                 join tbPayWay in myModels.SYS_PayWay on tbCalculate.PayWayID equals tbPayWay.PayWayID
                                 orderby tbCalculate.CalculateID descending
                                 select new ExamineVo
                                 {
                                     CalculateID = tbCalculate.CalculateID,//计费ID
                                     SignID = tbCalculate.SignID,//对账ID
                                     CalculateState = "已审核",//计费状态
                                     CalculateNumber = tbCalculate.CalculateNumber.Trim(),//收付款单号
                                     SettleAccounts = tbCalculate.SettleAccounts.Trim(),//收付款类型
                                     SFMoney = tbCalculate.SFMoney,//收付金额
                                     AlreadyCancelMoney = tbCalculate.AlreadyCancelMoney,//已核销金额
                                     StayCancelMoney = tbCalculate.StayCancelMoney,//待核销金额
                                     Currency = tbCalculate.Currency.Trim(),//币种
                                     SFDate1 = tbCalculate.SFDate.ToString(),//收付日期
                                     BankName = tbCalculate.BankName.Trim(),//银行名称
                                     BankAccount = tbCalculate.BankAccount.Trim(),//银行账号
                                     PayWay = tbPayWay.PayWay.Trim(),//付款方式
                                     ChequeNumber = tbCalculate.ChequeNumber.Trim(),//支票号
                                     ClientID = tbCalculate.ClientID,//客户ID  结算单位多条件查询
                                     PayWayID = tbCalculate.PayWayID,//付款方式ID
                                     ReckoningUnit = "",
                                     MessageID = tbCalculate.MessageID,
                                 }).ToList();

            if (listCalculate.Count() > 0) // Count 条数 两列放在同一列显示，在上面声明一个空的字段在表格显示，然后判断查询该显示哪张表的名称
            {
                for (int i = 0; i < listCalculate.Count(); i++)
                {
                    if (listCalculate[i].MessageID > 0)
                    {
                        var B = listCalculate[i].MessageID;  //listCalculate[i] 前面查询要用ToList()
                        var C = myModels.SYS_Message.Where(m => m.MessageID == B).Select(m => m.ChineseAbbreviation).Single();
                        listCalculate[i].ReckoningUnit = C;
                    }
                    else if (listCalculate[i].ClientID > 0)
                    {
                        var B = listCalculate[i].ClientID;
                        var C = myModels.SYS_Client.Where(m => m.ClientID == B).Select(m => m.ClientAbbreviation).Single();
                        listCalculate[i].ReckoningUnit = C;
                    }
                }
            }


            if (!string.IsNullOrEmpty(CalculateNumber))//收付款单号
            {
                CalculateNumber = CalculateNumber.Trim();
                listCalculate = listCalculate.Where(p => p.CalculateNumber.Contains(CalculateNumber)).ToList();
            }
            if (ClientID > 0)//结算单位
            {
                listCalculate = listCalculate.Where(m => m.ClientID == ClientID).ToList();
            }
            if (MessageID > 0)//结算单位
            {
                listCalculate = listCalculate.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(BankName))//银行名称
            {
                BankName = BankName.Trim();
                listCalculate = listCalculate.Where(p => p.BankName.Contains(BankName)).ToList();
            }
            if (!string.IsNullOrEmpty(SFDate))//收付日期
            {
                DateTime SFDate1 = Convert.ToDateTime(SFDate);
                listCalculate = listCalculate.Where(p => p.SFDate1.Contains(SFDate)).ToList();
            }
            if (PayWayID > 0)//付款方式ID
            {
                listCalculate = listCalculate.Where(m => m.PayWayID == PayWayID).ToList();
            }
            if (!string.IsNullOrEmpty(Currency))//币种
            {
                Currency = Currency.Trim();
                listCalculate = listCalculate.Where(p => p.Currency.Contains(Currency)).ToList();
            }
            if (!string.IsNullOrEmpty(SFMoney))//收付金额
            {
                SFMoney = SFMoney.Trim();
                listCalculate = listCalculate.Where(p => p.SFMoney.ToString().Contains(SFMoney)).ToList();
            }
            if (!string.IsNullOrEmpty(AlreadyCancelMoney))//已核销金额
            {
                AlreadyCancelMoney = AlreadyCancelMoney.Trim();
                listCalculate = listCalculate.Where(p => p.AlreadyCancelMoney.ToString().Contains(AlreadyCancelMoney)).ToList();
            }
            if (!string.IsNullOrEmpty(StayCancelMoney))//待核销金额
            {
                StayCancelMoney = StayCancelMoney.Trim();
                listCalculate = listCalculate.Where(p => p.StayCancelMoney.ToString().Contains(StayCancelMoney)).ToList();
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("状态");
            row1.CreateCell(1).SetCellValue("收付款单号");
            row1.CreateCell(2).SetCellValue("收付款类型");
            row1.CreateCell(3).SetCellValue("结算单位");
            row1.CreateCell(4).SetCellValue("收付金额");
            row1.CreateCell(5).SetCellValue("已核销金额");
            row1.CreateCell(6).SetCellValue("待核销金额");
            row1.CreateCell(7).SetCellValue("付款方式");
            row1.CreateCell(8).SetCellValue("币种");
            row1.CreateCell(9).SetCellValue("收付日期");
            row1.CreateCell(10).SetCellValue("银行名称");
            row1.CreateCell(11).SetCellValue("账号");
            row1.CreateCell(12).SetCellValue("支票号");
            //4、循环写入数据
            for (var i = 0; i < listCalculate.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listCalculate[i].CalculateState);
                rowTemp.CreateCell(1).SetCellValue(listCalculate[i].CalculateNumber);
                rowTemp.CreateCell(2).SetCellValue(listCalculate[i].SettleAccounts);
                rowTemp.CreateCell(3).SetCellValue(listCalculate[i].ReckoningUnit);
                rowTemp.CreateCell(4).SetCellValue(listCalculate[i].SFMoney.ToString());
                rowTemp.CreateCell(5).SetCellValue(listCalculate[i].AlreadyCancelMoney.ToString());
                rowTemp.CreateCell(6).SetCellValue(listCalculate[i].StayCancelMoney.ToString());
                rowTemp.CreateCell(7).SetCellValue(listCalculate[i].PayWay);
                rowTemp.CreateCell(8).SetCellValue(listCalculate[i].Currency);
                rowTemp.CreateCell(9).SetCellValue(listCalculate[i].SFDate1);
                rowTemp.CreateCell(10).SetCellValue(listCalculate[i].BankName);
                rowTemp.CreateCell(11).SetCellValue(listCalculate[i].BankAccount);
                rowTemp.CreateCell(12).SetCellValue(listCalculate[i].ChequeNumber);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司计费单报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);
            //6、文件名
            return File(exStream, "application/vnd.ms-excel", fileName);
        }



        #endregion


        /// <summary>
        /// 结算单号自动生成
        /// </summary>
        /// <returns></returns>
        public ActionResult getCode()
        {
            string strCurrentCode = "";//定义当前编码的字符串
            //查询所有部门信息，以DepartmentCode排序
            var listFlightNuber = (from tbHangBan in myModels.SYS_Sign
                                   orderby tbHangBan.SettleNumber
                                   select tbHangBan).ToList();
            if (listFlightNuber.Count > 0)//判断listDep中是否有数据
            {   //获取部门的数据的条数
                int count = listFlightNuber.Count;
                //获取最后一个科室模型//Employee_Department 需要添加引用using HMSBootstrap.Models;
                SYS_Sign modelNUM = listFlightNuber[count - 1];
                // Models.Employee_Department modelDep = listDep[count-1]; //不需要添加引用
                //获取最后一个科室编码
                int intCode = Convert.ToInt32(modelNUM.SettleNumber.Substring(9, 4));//Substring(1,4)截取字符串,1代表字符串起始位置，4表示截取长度
                //最后一个科室编码+1                              A  0000 0001  A0000 0000 0007
                intCode++;
                //对新的科室编码格式化
                strCurrentCode = intCode.ToString();
                var year = DateTime.Now.Year.ToString();
                var month = DateTime.Now.Month.ToString().Length == 1 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                var day = DateTime.Now.Day.ToString().Length == 1 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                for (int i = 0; i < 4; i++)
                {   //循环判断strCurrentCode的长度是否大于4，如果不是就在strCurrentCode前面加"0",否则strCurrentCode=strCurrentCode
                    strCurrentCode = strCurrentCode.Length < 4 ? "0" + strCurrentCode : strCurrentCode;//三目运算符
                }
                strCurrentCode = "JS" + year + month + day + strCurrentCode;
            }
            else
            {
                strCurrentCode = "JS000000000001";//如果count=0,科室编码strCurrentCode从D0001开始
            }
            return Json(strCurrentCode, JsonRequestBehavior.AllowGet);//返回Json格式的数据

        }

        /// <summary>
        /// 自动生成收付款单号
        /// </summary>
        /// <returns></returns>
        public ActionResult getCodee()
        {
            string strCurrentCode = "";//定义当前编码的字符串
            //查询所有部门信息，以DepartmentCode排序
            var listFlightNuber = (from tbHangBan in myModels.SYS_Calculate
                                   orderby tbHangBan.CalculateNumber
                                   select tbHangBan).ToList();
            if (listFlightNuber.Count > 0)//判断listDep中是否有数据
            {   //获取部门的数据的条数
                int count = listFlightNuber.Count;
                //获取最后一个科室模型//Employee_Department 需要添加引用using HMSBootstrap.Models;
                SYS_Calculate modelNUM = listFlightNuber[count - 1];
                // Models.Employee_Department modelDep = listDep[count-1]; //不需要添加引用
                //获取最后一个科室编码
                int intCode = Convert.ToInt32(modelNUM.CalculateNumber.Substring(9, 4));//Substring(1,4)截取字符串,1代表字符串起始位置，4表示截取长度
                //最后一个科室编码+1                              A  0000 0001  A0000 0000 0007
                intCode++;
                //对新的科室编码格式化
                strCurrentCode = intCode.ToString();
                var year = DateTime.Now.Year.ToString();
                var month = DateTime.Now.Month.ToString().Length == 1 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                var day = DateTime.Now.Day.ToString().Length == 1 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                for (int i = 0; i < 4; i++)
                {   //循环判断strCurrentCode的长度是否大于4，如果不是就在strCurrentCode前面加"0",否则strCurrentCode=strCurrentCode
                    strCurrentCode = strCurrentCode.Length < 4 ? "0" + strCurrentCode : strCurrentCode;//三目运算符
                }
                strCurrentCode = "FK" + year + month + day + strCurrentCode;
            }
            else
            {
                strCurrentCode = "FK000000000001";//如果count=0,科室编码strCurrentCode从D0001开始
            }
            return Json(strCurrentCode, JsonRequestBehavior.AllowGet);//返回Json格式的数据

        }


    }
}