
using Hotel.VO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SeaTransportation.Vo;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Areas.Forwarding.Controllers
{
    public class StatementController : Controller
    {
        Models.SeaTransportationEntities mymodels = new Models.SeaTransportationEntities();
        // GET: Forwarding/Statement
        #region 视图
        //车辆调度报表
        public ActionResult VehicleSchedulingReport()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //车辆作业明细表
        public ActionResult VehicleOperationSchedule()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //车辆作业汇总表
        public ActionResult VehicleOperationSummarySheet()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //司机作业明细表
        public ActionResult DriverWorkSchedule()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //司机作业汇总表
        public ActionResult DriverJobSummary()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //应收运费明细报表
        public ActionResult AccountsReceivableDetailFreight()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //应收运费汇总表
        public ActionResult SummaryFreightReceivable()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //应付运费明细表
        public ActionResult ScheduleFreightPayable()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //应付运费汇总表
        public ActionResult SummaryFreightPayable()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //#region
        ////应付报关费明细表
        //public ActionResult ScheduleCustomsChargesPayable()
        //{
        //    return View();
        //}
        ////应付报关费汇总表
        //public ActionResult SummaryCustomsChargesPayable()
        //{
        //    return View();
        //}
        ////船名航次明细表
        //public ActionResult NameVoyageSchedule()
        //{
        //    return View();
        //}
        ////船名航次汇总表
        //public ActionResult ShipNameVoyageSummarySheet()
        //{
        //    return View();
        //}
        //#endregion
        #endregion
        #region 方法
        //查询报表
        public ActionResult SelectStatement(BsgridPage BsgridPage, bool? VehicleSchedulingReport, int? EtrustID, string EtrustNmber, string Seal, string ArriveTime, string ArriveTimes, int? ClientAbbreviation, int? TiGuiDiDianID, int? HuaiGuiDiDianID, string Cupboard,int? ClientTypeID,int? ClientTypesID,int? VehicleInformationID,int? ChauffeurID,bool? V,bool? W,bool? Y,bool? V1,bool? W1,bool? V2)
        {
            var listSelectStatement = (from tbEtrust in mymodels.SYS_Etrust
                              join tbClient in mymodels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                              into tbClient
                              from Client in tbClient.DefaultIfEmpty()
                              join tbClientScontacts in mymodels.SYS_ClientScontacts on tbEtrust.ClientScontactsID equals tbClientScontacts.ClientScontactsID
                              into tbClientScontacts
                              from ClientScontacts in tbClientScontacts.DefaultIfEmpty()
                              join tbClientSite in mymodels.SYS_ClientSite on tbEtrust.ClientSiteID equals tbClientSite.ClientSiteID
                              into tbClientSite
                              from ClientSite in tbClientSite.DefaultIfEmpty()
                              join tbClientType in mymodels.SYS_ClientType on tbEtrust.UndertakeID equals tbClientType.ClientTypeID
                              into tbClientType
                              from ClientType in tbClientType.DefaultIfEmpty()
                              join tbClients in mymodels.SYS_Client on ClientType.ClientID equals tbClients.ClientID
                              into tbClients
                              from Clients in tbClients.DefaultIfEmpty()
                              join tbVehicleInformation in mymodels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicleInformation.VehicleInformationID
                              into tbVehicleInformation
                              from VehicleInformation in tbVehicleInformation.DefaultIfEmpty()
                              join tbChauffeur in mymodels.SYS_Chauffeur on VehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                              into tbChauffeur
                              from Chauffeur in tbChauffeur.DefaultIfEmpty()
                              join tbBracket in mymodels.SYS_Bracket on VehicleInformation.BracketID equals tbBracket.BracketID
                              into tbBracket
                              from Bracket in tbBracket.DefaultIfEmpty()
                              join tbMention in mymodels.SYS_Mention on tbEtrust.TiGuiDiDianID equals tbMention.MentionID
                              into tbMention
                              from Mention in tbMention.DefaultIfEmpty()
                              join tbMentions in mymodels.SYS_Mention on tbEtrust.HuaiGuiDiDianID equals tbMentions.MentionID
                              into tbMentions
                              from Mentions in tbMentions.DefaultIfEmpty()
                              join tbGatedot in mymodels.SYS_Gatedot on Mention.GatedotID equals tbGatedot.GatedotID
                              into tbGatedot
                              from Gatedot in tbGatedot.DefaultIfEmpty()
                              join tbGatedots in mymodels.SYS_Gatedot on Mentions.GatedotID equals tbGatedots.GatedotID
                              into tbGatedots
                              from Gatedots in tbGatedots.DefaultIfEmpty()
                              join tbShip in mymodels.SYS_Ship on tbEtrust.ShipID equals tbShip.ShipID
                              into tbShip
                              from Ship in tbShip.DefaultIfEmpty()
                              join tbClinetTypes in mymodels.SYS_ClientType on Ship.ClientTypeID equals tbClinetTypes.ClientTypeID
                              into tbClinetTypes
                              from ClinetTypes in tbClinetTypes.DefaultIfEmpty()
                              join tbClientl in mymodels.SYS_Client on ClinetTypes.ClientID equals tbClientl.ClientID
                              into tbClientl
                              from Clientl in tbClientl.DefaultIfEmpty()
                              join tbGatedotl in mymodels.SYS_Gatedot on ClientSite.GatedotID equals tbGatedotl.GatedotID
                              into tbGatedotl
                              from Gatedotl in tbGatedotl.DefaultIfEmpty()
                              join tbStaff in mymodels.SYS_Staff on tbEtrust.StaffID equals tbStaff.StaffID
                              into tbStaff
                              from staff in tbStaff.DefaultIfEmpty()
                              join tbMessage in mymodels.SYS_Message on tbEtrust.MessageID equals tbMessage.MessageID
                              into tbMessage
                              from Message in tbMessage.DefaultIfEmpty()
                             // join tbCharge in mymodels.SYS_Charge on tbEtrust.EtrustID equals tbCharge.EtrustID
                             // into tbCharge
                             // from Charge in tbCharge.DefaultIfEmpty()
                             // join tbExpense in mymodels.SYS_Expense on Charge.ExpenseID equals tbExpense.ExpenseID
                             // into tbExpense
                             //from Expense in tbExpense.DefaultIfEmpty()
                              orderby tbEtrust.EtrustID descending
                              select new EtrustVo
                              {
                                  ChauffeurID= Chauffeur.ChauffeurID,
                                  MessageName = Message.ChineseName.Trim(),
                                  AssistBusiness = tbEtrust.AssistBusiness.Trim(),
                                  OddNumber = tbEtrust.OddNumber.Trim(),
                                  EtrustID = tbEtrust.EtrustID,
                                  Undertake = Clients.ClientAbbreviation.Trim(),
                                  EtrustNmber = tbEtrust.EtrustNmber.Trim(),
                                  WorkCategory = tbEtrust.WorkCategory.Trim(),
                                  CheXing = tbEtrust.CheXing,
                                  VehicleCode = VehicleInformation.VehicleCode,
                                  ChauffeurName = Chauffeur.ChauffeurName.Trim(),
                                  ChauffeurNumber = Chauffeur.ChauffeurNumber,
                                  BracketCode = Bracket.BracketCode,
                                  arriveTime = tbEtrust.ArriveTime.ToString(),
                                  ArriveTime = tbEtrust.ArriveTime,
                                  HangCi = tbEtrust.HangCi.Trim(),
                                  CabinetType = tbEtrust.CabinetType.Trim(),
                                  BookingSpace = tbEtrust.BookingSpace.Trim(),
                                  Seal = tbEtrust.Seal.Trim(),
                                  CarriageNumber = tbEtrust.CarriageNumber.Trim(),
                                  EtrustType = tbEtrust.EtrustType.Trim(),
                                  ClientAbbreviation = Client.ClientAbbreviation,
                                  ContactsName = ClientSite.ContactsName.Trim(),
                                  ContactsPhone = ClientSite.ContactsPhone.Trim(),
                                  contactsName = ClientScontacts.contactsName.Trim(),
                                  contactsPhone = ClientScontacts.contactsPhone.Trim(),
                                  ClientID = Client.ClientID,
                                  Cupboard = tbEtrust.Cupboard.Trim(),
                                  TiGuiDiDianID = tbEtrust.TiGuiDiDianID,
                                  HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,
                                  AuditType = tbEtrust.AuditType.Trim(),
                                  MessageID = tbEtrust.MessageID,
                                  ClientCode = tbEtrust.ClientCode.Trim(),
                                  ChineseName = Client.ChineseName.Trim(),
                                  ClientScontactsID = tbEtrust.ClientScontactsID,
                                  ClientMobile = ClientScontacts.contactsPhone.Trim(),
                                  StaffID = tbEtrust.StaffID,
                                  PlanHandTime = tbEtrust.PlanHandTime,
                                  planHandTime = tbEtrust.PlanHandTime.ToString(),
                                  HuoZhuName = tbEtrust.HuoZhuName.Trim(),
                                  IndentNumber = tbEtrust.IndentNumber.Trim(),
                                  ShipID = tbEtrust.ShipID,
                                  CollectionMoney = tbEtrust.CollectionMoney,
                                  CargoClassify = tbEtrust.CargoClassify.Trim(),
                                  CargoWeight = tbEtrust.CargoWeight.Trim(),
                                  ShuiBl = tbEtrust.ShuiBl,
                                  ManagementFee = tbEtrust.ManagementFee,
                                  T = Gatedot.GatedotName.Trim(),
                                  H = Gatedots.GatedotName.Trim(),
                                  PortID = tbEtrust.PortID,
                                  GoalHarborID = tbEtrust.GoalHarborID,
                                  FactoryName = ClientSite.FactoryName.Trim(),
                                  FactoryCode = ClientSite.FactoryCode.Trim(),
                                  ClientSite = ClientSite.ClientSite.Trim(),
                                  Chuangs = Clientl.ClientAbbreviation.Trim(),
                                  HuiChangTime = tbEtrust.HuiChangTime,
                                  huiChangTime = tbEtrust.HuiChangTime.ToString(),
                                  ClientSiteID = ClientSite.ClientSiteID,
                                  GatedotName = Gatedotl.GatedotName.Trim(),
                                  SpecialNeed = tbEtrust.SpecialNeed.Trim(),
                                  AttentionMatter = tbEtrust.AttentionMatter.Trim(),
                                  KaiCangTime = tbEtrust.KaiCangTime,
                                  kaiCangTime = tbEtrust.KaiCangTime.ToString(),
                                  StaffName = staff.StaffName.Trim(),
                                  TG = Mention.Abbreviation.Trim(),
                                  HG = Mentions.Abbreviation.Trim(),
                                  OfferDetailID = tbEtrust.OfferDetailID,
                                  YSLX = Gatedot.GatedotName.Trim() + " - " + Gatedotl.GatedotName.Trim() + " - " + Gatedots.GatedotName.Trim(),
                                  VehicleInformationID = tbEtrust.VehicleInformationID,
                                  Shij = Chauffeur.ChauffeurName.Trim(),
                                  ICNumber = Chauffeur.ICNumber.Trim(),
                                  PhoneOne = Chauffeur.PhoneOne.Trim(),
                                  BracketType = Bracket.BracketType.Trim(),
                                  UndertakeID = tbEtrust.UndertakeID,
                                  paiCheDate = tbEtrust.PaiCheDate.ToString(),
                                  FollowExplain = tbEtrust.FollowExplain.Trim(),
                                  PlateNumbers= VehicleInformation.PlateNumbers.Trim(),
                                  BracketTag= Bracket.BracketTag.Trim(),
                                  Name = Chauffeur.ChauffeurName.Trim()
                                  // SettleAccounts= Expense.SettleAccounts.Trim()
                              }).ToList();
            if (ClientTypeID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.UndertakeID == ClientTypeID).ToList();
            }
            if (ClientTypesID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.UndertakeID == ClientTypesID).ToList();
            }
            if (VehicleInformationID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (ChauffeurID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (EtrustID >= 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
            }
            if (ClientAbbreviation > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.ClientID == ClientAbbreviation).ToList();
            }
            if (TiGuiDiDianID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (!string.IsNullOrEmpty(EtrustNmber))
            {
                listSelectStatement = listSelectStatement.Where(m => m.EtrustNmber.Contains(EtrustNmber)).ToList();
            }
            if (!string.IsNullOrEmpty(Seal))
            {
                listSelectStatement = listSelectStatement.Where(m => m.Seal.Contains(Seal)).ToList();
            }
            if (!string.IsNullOrEmpty(Cupboard))
            {
                listSelectStatement = listSelectStatement.Where(m => m.Cupboard.Contains(Cupboard)).ToList();
            }
            if (!string.IsNullOrEmpty(ArriveTime) && string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveTime);
                    listSelectStatement = listSelectStatement.Where(m => m.ArriveTime == Time).ToList();
                }
                catch (Exception)
                {
                    listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(ArriveTimes);
                    listSelectStatement = listSelectStatement.Where(m => m.ArriveTime <= Times).ToList();
                }
                catch (Exception)
                {
                    listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveTime);
                    DateTime Times = Convert.ToDateTime(ArriveTimes);
                    listSelectStatement = listSelectStatement.Where(m => m.ArriveTime >= Time && m.ArriveTime <= Times).ToList();
                }
                catch (Exception)
                {
                    listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            if (VehicleSchedulingReport == true)
            {
                listSelectStatement = listSelectStatement.Where(m => m.CarriageNumber != null).ToList();
            }
            //if (AccountsReceivableDetailFreight==true)
            //{
            //    listSelectStatement = listSelectStatement.Where(m => m.SettleAccounts =="应收").ToList();
            //}
            if (W==true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listSelectStatement.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("车辆调度报表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                row1.CreateCell(0).SetCellValue("运输单号");
                row1.CreateCell(1).SetCellValue("状态");
                row1.CreateCell(2).SetCellValue("客户简称");
                row1.CreateCell(3).SetCellValue("工厂名称");
                row1.CreateCell(4).SetCellValue("工厂联系人");
                row1.CreateCell(5).SetCellValue("工厂联系电话");
                row1.CreateCell(6).SetCellValue("到厂时间");
                row1.CreateCell(7).SetCellValue("箱型");
                row1.CreateCell(8).SetCellValue("订舱号S / 0");
                row1.CreateCell(9).SetCellValue("封条号");
                row1.CreateCell(10).SetCellValue("航次");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].CarriageNumber);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].EtrustType);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].ContactsName);
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].ContactsPhone);
                    rowTemp.CreateCell(6).SetCellValue(listExaminee[i].arriveTimes);
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].CabinetType);
                    rowTemp.CreateCell(8).SetCellValue(listExaminee[i].BookingSpace);
                    rowTemp.CreateCell(9).SetCellValue(listExaminee[i].Seal);
                    rowTemp.CreateCell(10).SetCellValue(listExaminee[i].HangCi);
                }
                //输出的文件名称
                string fileName = "车辆调度报表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (V == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listSelectStatement.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("车辆作业明细表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                row1.CreateCell(0).SetCellValue("运输单号");
                row1.CreateCell(1).SetCellValue("托架尾牌号");
                row1.CreateCell(2).SetCellValue("车牌号");
                row1.CreateCell(3).SetCellValue("承运车队");
                row1.CreateCell(4).SetCellValue("客户简称");
                row1.CreateCell(5).SetCellValue("到厂时间");
                row1.CreateCell(6).SetCellValue("箱型");
                row1.CreateCell(8).SetCellValue("订舱号S / 0");
                row1.CreateCell(9).SetCellValue("封条号");
                row1.CreateCell(10).SetCellValue("航次");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].CarriageNumber);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].BracketTag);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].PlateNumbers);
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].Undertake);
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].arriveTimes);
                    rowTemp.CreateCell(6).SetCellValue(listExaminee[i].CabinetType);
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].BookingSpace);
                    rowTemp.CreateCell(8).SetCellValue(listExaminee[i].Seal);
                    rowTemp.CreateCell(9).SetCellValue(listExaminee[i].HangCi);
                }
                //输出的文件名称
                string fileName = "车辆作业明细表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (Y == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listSelectStatement.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("司机作业明细表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                 row1.CreateCell(0).SetCellValue("司机姓名");
                row1.CreateCell(1).SetCellValue("司机编号");
                row1.CreateCell(2).SetCellValue("运输单号");
                row1.CreateCell(3).SetCellValue("托架尾牌号");
                row1.CreateCell(4).SetCellValue("车牌号");
                row1.CreateCell(5).SetCellValue("承运车队");
                row1.CreateCell(6).SetCellValue("客户简称");
                row1.CreateCell(7).SetCellValue("到厂时间");
                row1.CreateCell(8).SetCellValue("箱型");
                row1.CreateCell(9).SetCellValue("订舱号S / 0 ");
                row1.CreateCell(10).SetCellValue("封条号");
                row1.CreateCell(11).SetCellValue("航次");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ChauffeurName);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ChauffeurNumber);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].CarriageNumber);
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].BracketTag);
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].PlateNumbers);
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].Undertake);
                    rowTemp.CreateCell(6).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].arriveTimes);
                    rowTemp.CreateCell(8).SetCellValue(listExaminee[i].CabinetType);
                    rowTemp.CreateCell(9).SetCellValue(listExaminee[i].BookingSpace);
                    rowTemp.CreateCell(10).SetCellValue(listExaminee[i].Seal);
                    rowTemp.CreateCell(11).SetCellValue(listExaminee[i].HangCi);
                }
                //输出的文件名称
                string fileName = "司机作业明细表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (V1 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(listSelectStatement);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.VehicleSchedulingReport rp = new Print.VehicleSchedulingReport();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\VehicleSchedulingReport.rpt";//  \转义字符
                                                                 //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            if (W1 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(listSelectStatement);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.VehicleOperationSchedule rp = new Print.VehicleOperationSchedule();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\VehicleOperationSchedule.rpt";//  \转义字符
                                                                              //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            if (V2 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(listSelectStatement);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.DriverWorkSchedule rp = new Print.DriverWorkSchedule();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\DriverWorkSchedule.rpt";//  \转义字符
                                                                              //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            int totalRow = listSelectStatement.Count();
            List<EtrustVo> list = listSelectStatement
                                        .Skip(BsgridPage.GetStartIndex())
                                        .Take(BsgridPage.pageSize)
                                        .ToList();
            Bsgrid<EtrustVo> bsgrid = new Bsgrid<EtrustVo>()
            {
                success = true,
                totalRows = totalRow,
                curPage = BsgridPage.curPage,
                data = list
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        //查询车辆作业汇总表
        public ActionResult SelectVehicleOperationSummarySheet(BsgridPage BsgridPage,int? ClientTypeID,int? ClientTypesID,int? VehicleInformationID,int? ChauffeurID,bool? X,bool? V1)
        {
            var VehicleOperationSummarySheet = (from tbEtrust in mymodels.SYS_Etrust
                                                join tbOfferDetail in mymodels.SYS_OfferDetail on tbEtrust.OfferDetailID equals tbOfferDetail.OfferDetailID
                                                join tbVehicleInformation in mymodels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicleInformation.VehicleInformationID
                                                join tbChauffeur in mymodels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                                                select new EtrustVo
                                                {
                                                    M = tbOfferDetail.Money,
                                                    S = (tbOfferDetail.Money * tbEtrust.ShuiBl) / 100,
                                                    J = 0,
                                                    C=(tbOfferDetail.Money * tbChauffeur.DeductRatio) / 100,
                                                   PlateNumbers =tbVehicleInformation.PlateNumbers,
                                                   VehicleCode=tbVehicleInformation.VehicleCode.Trim(),
                                                    ManagementFee=tbEtrust.ManagementFee,
                                                   VehicleInformationID =tbEtrust.VehicleInformationID,
                                                   ChauffeurID=tbChauffeur.ChauffeurID,
                                                   UndertakeID=tbEtrust.UndertakeID
                                               }).ToList();
            if (VehicleOperationSummarySheet.Count()>0)
            {
                var A = mymodels.SYS_Etrust.Where(m=>m.VehicleInformationID>0).Select(m => m.VehicleInformationID).ToList().Distinct().ToArray();
                for (int i = 0; i < A.Count(); i++)
                {
                    var C = VehicleOperationSummarySheet.Where(m => m.VehicleInformationID == A[i]).ToList();
                    var M = C.Sum(m => m.M + m.M);
                    var S = C.Sum(m => m.S + m.S);
                    var c = C.Sum(m => m.C + m.C);
                    var ManagementFee= C.Sum(m => m.ManagementFee + m.ManagementFee);
                    for (int n = 0; n < C.Count(); n++)
                    {
                        C[n].M = M;
                        C[n].S = S;
                        C[n].C = c;
                        C[n].ManagementFee = ManagementFee;
                        C[n].J = M+S+c+ ManagementFee;
                    }
                    for (int m = 0; m < (C.Count() - 1); m++)
                    {
                        VehicleOperationSummarySheet.Remove(C[m]);
                    }
                }
            }
            if (VehicleInformationID>0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (ChauffeurID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (ClientTypeID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.UndertakeID == ClientTypeID).ToList();
            }
            if (ClientTypesID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.UndertakeID == ClientTypesID).ToList();
            }
            if (X == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = VehicleOperationSummarySheet.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("车辆作业汇总表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值

              

                row1.CreateCell(0).SetCellValue("车牌号");
                row1.CreateCell(1).SetCellValue("车辆代码");
                row1.CreateCell(2).SetCellValue("管理费");
                row1.CreateCell(3).SetCellValue("税金");
                row1.CreateCell(4).SetCellValue("司机提成");
                row1.CreateCell(5).SetCellValue("运费");
                row1.CreateCell(6).SetCellValue("合计");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].PlateNumbers);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].VehicleCode);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].ManagementFee.ToString());
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].S.ToString());
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].C.ToString());
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].M.ToString());
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].J.ToString());
                }
                //输出的文件名称
                string fileName = "车辆作业汇总表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (V1 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(VehicleOperationSummarySheet);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.VehicleOperationSummarySheet rp = new Print.VehicleOperationSummarySheet();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\VehicleOperationSummarySheet.rpt";//  \转义字符
                                                                              //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            int totalRow = VehicleOperationSummarySheet.Count();
            List<EtrustVo> list = VehicleOperationSummarySheet
                                        .Skip(BsgridPage.GetStartIndex())
                                        .Take(BsgridPage.pageSize).Distinct()
                                        .ToList();
            Bsgrid<EtrustVo> bsgrid = new Bsgrid<EtrustVo>()
            {
                success = true,
                totalRows = totalRow,
                curPage = BsgridPage.curPage,
                data = list
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        //查询司机作业汇总表
        public ActionResult SelectDriverJobSummary(BsgridPage BsgridPage, int? ClientTypeID, int? ClientTypesID, int? VehicleInformationID, int? ChauffeurID,bool? Z,bool? Z1)
        {
            var VehicleOperationSummarySheet = (from tbEtrust in mymodels.SYS_Etrust
                                                join tbOfferDetail in mymodels.SYS_OfferDetail on tbEtrust.OfferDetailID equals tbOfferDetail.OfferDetailID
                                                join tbVehicleInformation in mymodels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicleInformation.VehicleInformationID
                                                join tbChauffeur in mymodels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                                                select new EtrustVo
                                                {
                                                    M = tbOfferDetail.Money,
                                                    S = (tbOfferDetail.Money * tbEtrust.ShuiBl) / 100,
                                                    J = 0,
                                                    C = (tbOfferDetail.Money * tbChauffeur.DeductRatio) / 100,
                                                    PlateNumbers = tbVehicleInformation.PlateNumbers,
                                                    VehicleCode = tbVehicleInformation.VehicleCode.Trim(),
                                                    ManagementFee = tbEtrust.ManagementFee,
                                                    VehicleInformationID = tbEtrust.VehicleInformationID,
                                                    ChauffeurID = tbChauffeur.ChauffeurID,
                                                    UndertakeID = tbEtrust.UndertakeID,
                                                    ChauffeurName=tbChauffeur.ChauffeurName.Trim(),
                                                    ChauffeurNumber=tbChauffeur.ChauffeurNumber,
                                                    Name = tbChauffeur.ChauffeurName.Trim(),
                                                }).ToList();
            if (VehicleOperationSummarySheet.Count() > 0)
            {
                var A = mymodels.SYS_VehicleInformation.Where(m => m.ChauffeurID > 0).Select(m => m.ChauffeurID).ToList().Distinct().ToArray();
                for (int i = 0; i < A.Count(); i++)
                {
                    var C = VehicleOperationSummarySheet.Where(m => m.ChauffeurID == A[i]).ToList();
                    if (C.Count()>0)
                    {
                        var M = C.Sum(m => m.M + m.M);
                        var S = C.Sum(m => m.S + m.S);
                        var c = C.Sum(m => m.C + m.C);
                        var ManagementFee = C.Sum(m => m.ManagementFee + m.ManagementFee);
                        for (int n = 0; n < C.Count(); n++)
                        {
                            C[n].M = M;
                            C[n].S = S;
                            C[n].C = c;
                            C[n].ManagementFee = ManagementFee;
                            C[n].J = M + S + c + ManagementFee;
                        }
                        for (int m = 0; m < (C.Count() - 1); m++)
                        {
                            VehicleOperationSummarySheet.Remove(C[m]);
                        }
                    }
                }
            }
            if (VehicleInformationID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (ChauffeurID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (ClientTypeID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.UndertakeID == ClientTypeID).ToList();
            }
            if (ClientTypesID > 0)
            {
                VehicleOperationSummarySheet = VehicleOperationSummarySheet.Where(m => m.UndertakeID == ClientTypesID).ToList();
            }
            if (Z == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = VehicleOperationSummarySheet.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("司机作业汇总表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                row1.CreateCell(0).SetCellValue("司机姓名");
                row1.CreateCell(1).SetCellValue("司机代码");
                row1.CreateCell(2).SetCellValue("管理费");
                row1.CreateCell(3).SetCellValue("税金");
                row1.CreateCell(4).SetCellValue("司机提成");
                row1.CreateCell(5).SetCellValue("运费");
                row1.CreateCell(6).SetCellValue("合计");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ChauffeurName);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ChauffeurNumber);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].ManagementFee.ToString());
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].S.ToString());
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].C.ToString());
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].M.ToString());
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].J.ToString());
                }
                //输出的文件名称
                string fileName = "司机作业汇总表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (Z1 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(VehicleOperationSummarySheet);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.DriverJobSummary rp = new Print.DriverJobSummary();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\DriverJobSummary.rpt";//  \转义字符
                                                                       //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            int totalRow = VehicleOperationSummarySheet.Count();
            List<EtrustVo> list = VehicleOperationSummarySheet
                                        .Skip(BsgridPage.GetStartIndex())
                                        .Take(BsgridPage.pageSize).Distinct()
                                        .ToList();
            Bsgrid<EtrustVo> bsgrid = new Bsgrid<EtrustVo>()
            {
                success = true,
                totalRows = totalRow,
                curPage = BsgridPage.curPage,
                data = list
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        //查询应收付运费明细表
        public ActionResult SelectAccountsReceivableDetailFreight(BsgridPage BsgridPage, bool? VehicleSchedulingReport, int? EtrustID, string EtrustNmber, string Seal, string ArriveTime, string ArriveTimes, int? ClientAbbreviation, int? TiGuiDiDianID, int? HuaiGuiDiDianID, string Cupboard, int? ClientTypeID, int? ClientTypesID, int? VehicleInformationID, int? ChauffeurID, bool? AccountsReceivableDetailFreight,bool? ScheduleFreightPayable,bool? V,bool?W,bool? V1)
        {
            var listSelectStatement = (from tbEtrust in mymodels.SYS_Etrust
                                       join tbClient in mymodels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                                       into tbClient
                                       from Client in tbClient.DefaultIfEmpty()
                                       join tbClientScontacts in mymodels.SYS_ClientScontacts on tbEtrust.ClientScontactsID equals tbClientScontacts.ClientScontactsID
                                       into tbClientScontacts
                                       from ClientScontacts in tbClientScontacts.DefaultIfEmpty()
                                       join tbClientSite in mymodels.SYS_ClientSite on tbEtrust.ClientSiteID equals tbClientSite.ClientSiteID
                                       into tbClientSite
                                       from ClientSite in tbClientSite.DefaultIfEmpty()
                                       join tbClientType in mymodels.SYS_ClientType on tbEtrust.UndertakeID equals tbClientType.ClientTypeID
                                       into tbClientType
                                       from ClientType in tbClientType.DefaultIfEmpty()
                                       join tbClients in mymodels.SYS_Client on ClientType.ClientID equals tbClients.ClientID
                                       into tbClients
                                       from Clients in tbClients.DefaultIfEmpty()
                                       join tbVehicleInformation in mymodels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicleInformation.VehicleInformationID
                                       into tbVehicleInformation
                                       from VehicleInformation in tbVehicleInformation.DefaultIfEmpty()
                                       join tbChauffeur in mymodels.SYS_Chauffeur on VehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                                       into tbChauffeur
                                       from Chauffeur in tbChauffeur.DefaultIfEmpty()
                                       join tbBracket in mymodels.SYS_Bracket on VehicleInformation.BracketID equals tbBracket.BracketID
                                       into tbBracket
                                       from Bracket in tbBracket.DefaultIfEmpty()
                                       join tbMention in mymodels.SYS_Mention on tbEtrust.TiGuiDiDianID equals tbMention.MentionID
                                       into tbMention
                                       from Mention in tbMention.DefaultIfEmpty()
                                       join tbMentions in mymodels.SYS_Mention on tbEtrust.HuaiGuiDiDianID equals tbMentions.MentionID
                                       into tbMentions
                                       from Mentions in tbMentions.DefaultIfEmpty()
                                       join tbGatedot in mymodels.SYS_Gatedot on Mention.GatedotID equals tbGatedot.GatedotID
                                       into tbGatedot
                                       from Gatedot in tbGatedot.DefaultIfEmpty()
                                       join tbGatedots in mymodels.SYS_Gatedot on Mentions.GatedotID equals tbGatedots.GatedotID
                                       into tbGatedots
                                       from Gatedots in tbGatedots.DefaultIfEmpty()
                                       join tbShip in mymodels.SYS_Ship on tbEtrust.ShipID equals tbShip.ShipID
                                       into tbShip
                                       from Ship in tbShip.DefaultIfEmpty()
                                       join tbClinetTypes in mymodels.SYS_ClientType on Ship.ClientTypeID equals tbClinetTypes.ClientTypeID
                                       into tbClinetTypes
                                       from ClinetTypes in tbClinetTypes.DefaultIfEmpty()
                                       join tbClientl in mymodels.SYS_Client on ClinetTypes.ClientID equals tbClientl.ClientID
                                       into tbClientl
                                       from Clientl in tbClientl.DefaultIfEmpty()
                                       join tbGatedotl in mymodels.SYS_Gatedot on ClientSite.GatedotID equals tbGatedotl.GatedotID
                                       into tbGatedotl
                                       from Gatedotl in tbGatedotl.DefaultIfEmpty()
                                       join tbStaff in mymodels.SYS_Staff on tbEtrust.StaffID equals tbStaff.StaffID
                                       into tbStaff
                                       from staff in tbStaff.DefaultIfEmpty()
                                       join tbMessage in mymodels.SYS_Message on tbEtrust.MessageID equals tbMessage.MessageID
                                       into tbMessage
                                       from Message in tbMessage.DefaultIfEmpty()
                                       join tbCharge in mymodels.SYS_Charge on tbEtrust.EtrustID equals tbCharge.EtrustID
                                       into tbCharge
                                       from Charge in tbCharge.DefaultIfEmpty()
                                       join tbExpense in mymodels.SYS_Expense on Charge.ExpenseID equals tbExpense.ExpenseID
                                       into tbExpense
                                       from Expense in tbExpense.DefaultIfEmpty()
                                       orderby tbEtrust.EtrustID descending
                                       select new EtrustVo
                                       {
                                           ChauffeurID = Chauffeur.ChauffeurID,
                                           MessageName = Message.ChineseName.Trim(),
                                           AssistBusiness = tbEtrust.AssistBusiness.Trim(),
                                           OddNumber = tbEtrust.OddNumber.Trim(),
                                           EtrustID = tbEtrust.EtrustID,
                                           Undertake = Clients.ClientAbbreviation,
                                           EtrustNmber = tbEtrust.EtrustNmber.Trim(),
                                           WorkCategory = tbEtrust.WorkCategory.Trim(),
                                           CheXing = tbEtrust.CheXing,
                                           VehicleCode = VehicleInformation.VehicleCode,
                                           ChauffeurName = Chauffeur.ChauffeurName.Trim(),
                                           ChauffeurNumber = Chauffeur.ChauffeurNumber,
                                           BracketCode = Bracket.BracketCode,
                                           arriveTime = tbEtrust.ArriveTime.ToString(),
                                           ArriveTime = tbEtrust.ArriveTime,
                                           HangCi = tbEtrust.HangCi.Trim(),
                                           CabinetType = tbEtrust.CabinetType.Trim(),
                                           BookingSpace = tbEtrust.BookingSpace.Trim(),
                                           Seal = tbEtrust.Seal.Trim(),
                                           CarriageNumber = tbEtrust.CarriageNumber,
                                           EtrustType = tbEtrust.EtrustType,
                                           ClientAbbreviation = Client.ClientAbbreviation,
                                           ContactsName = ClientSite.ContactsName.Trim(),
                                           ContactsPhone = ClientSite.ContactsPhone.Trim(),
                                           contactsName = ClientScontacts.contactsName,
                                           contactsPhone = ClientScontacts.contactsPhone,
                                           ClientID = Client.ClientID,
                                           Cupboard = tbEtrust.Cupboard.Trim(),
                                           TiGuiDiDianID = tbEtrust.TiGuiDiDianID,
                                           HuaiGuiDiDianID = tbEtrust.HuaiGuiDiDianID,
                                           AuditType = tbEtrust.AuditType.Trim(),
                                           MessageID = tbEtrust.MessageID,
                                           ClientCode = tbEtrust.ClientCode.Trim(),
                                           ChineseName = Client.ChineseName.Trim(),
                                           ClientScontactsID = tbEtrust.ClientScontactsID,
                                           ClientMobile = ClientScontacts.contactsPhone.Trim(),
                                           StaffID = tbEtrust.StaffID,
                                           PlanHandTime = tbEtrust.PlanHandTime,
                                           planHandTime = tbEtrust.PlanHandTime.ToString(),
                                           HuoZhuName = tbEtrust.HuoZhuName.Trim(),
                                           IndentNumber = tbEtrust.IndentNumber.Trim(),
                                           ShipID = tbEtrust.ShipID,
                                           CollectionMoney = tbEtrust.CollectionMoney,
                                           CargoClassify = tbEtrust.CargoClassify.Trim(),
                                           CargoWeight = tbEtrust.CargoWeight.Trim(),
                                           ShuiBl = tbEtrust.ShuiBl,
                                           ManagementFee = tbEtrust.ManagementFee,
                                           T = Gatedot.GatedotName.Trim(),
                                           H = Gatedots.GatedotName.Trim(),
                                           PortID = tbEtrust.PortID,
                                           GoalHarborID = tbEtrust.GoalHarborID,
                                           FactoryName = ClientSite.FactoryName.Trim(),
                                           FactoryCode = ClientSite.FactoryCode.Trim(),
                                           ClientSite = ClientSite.ClientSite.Trim(),
                                           Chuangs = Clientl.ClientAbbreviation.Trim(),
                                           HuiChangTime = tbEtrust.HuiChangTime,
                                           huiChangTime = tbEtrust.HuiChangTime.ToString(),
                                           ClientSiteID = ClientSite.ClientSiteID,
                                           GatedotName = Gatedotl.GatedotName.Trim(),
                                           SpecialNeed = tbEtrust.SpecialNeed.Trim(),
                                           AttentionMatter = tbEtrust.AttentionMatter.Trim(),
                                           KaiCangTime = tbEtrust.KaiCangTime,
                                           kaiCangTime = tbEtrust.KaiCangTime.ToString(),
                                           StaffName = staff.StaffName.Trim(),
                                           TG = Mention.Abbreviation.Trim(),
                                           HG = Mentions.Abbreviation.Trim(),
                                           OfferDetailID = tbEtrust.OfferDetailID,
                                           YSLX = Gatedot.GatedotName.Trim() + " - " + Gatedotl.GatedotName.Trim() + " - " + Gatedots.GatedotName.Trim(),
                                           VehicleInformationID = tbEtrust.VehicleInformationID,
                                           Shij = Chauffeur.ChauffeurName.Trim(),
                                           ICNumber = Chauffeur.ICNumber.Trim(),
                                           PhoneOne = Chauffeur.PhoneOne.Trim(),
                                           BracketType = Bracket.BracketType.Trim(),
                                           UndertakeID = tbEtrust.UndertakeID,
                                           paiCheDate = tbEtrust.PaiCheDate.ToString(),
                                           FollowExplain = tbEtrust.FollowExplain.Trim(),
                                           PlateNumbers = VehicleInformation.PlateNumbers.Trim(),
                                           BracketTag = Bracket.BracketTag.Trim(),
                                           SettleAccounts = Expense.SettleAccounts.Trim(),
                                           ExpenseName=Expense.ExpenseName.Trim(),
                                           UnitPrice=Charge.UnitPrice
                                       }).ToList();
            if (ClientTypeID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.UndertakeID == ClientTypeID).ToList();
            }
            if (ClientTypesID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.UndertakeID == ClientTypesID).ToList();
            }
            if (VehicleInformationID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.VehicleInformationID == VehicleInformationID).ToList();
            }
            if (ChauffeurID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (EtrustID >= 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
            }
            if (ClientAbbreviation > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.ClientID == ClientAbbreviation).ToList();
            }
            if (TiGuiDiDianID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
            }
            if (HuaiGuiDiDianID > 0)
            {
                listSelectStatement = listSelectStatement.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
            }
            if (!string.IsNullOrEmpty(ArriveTime) && string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveTime);
                    listSelectStatement = listSelectStatement.Where(m => m.ArriveTime == Time).ToList();
                }
                catch (Exception)
                {
                    listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(ArriveTimes);
                    listSelectStatement = listSelectStatement.Where(m => m.ArriveTime <= Times).ToList();
                }
                catch (Exception)
                {
                    listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveTime);
                    DateTime Times = Convert.ToDateTime(ArriveTimes);
                    listSelectStatement = listSelectStatement.Where(m => m.ArriveTime >= Time && m.ArriveTime <= Times).ToList();
                }
                catch (Exception)
                {
                    listSelectStatement = listSelectStatement.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            if (VehicleSchedulingReport == true)
            {
                listSelectStatement = listSelectStatement.Where(m => m.CarriageNumber != null).ToList();
            }
            if (AccountsReceivableDetailFreight == true)
            {
                listSelectStatement = listSelectStatement.Where(m => m.SettleAccounts == "应收").ToList();
            }
            if (ScheduleFreightPayable==true)
            {
                listSelectStatement = listSelectStatement.Where(m => m.SettleAccounts == "应付").ToList();
            }

            if (V == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listSelectStatement.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("应收运费明细报表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                 row1.CreateCell(0).SetCellValue("费用类型");
                row1.CreateCell(1).SetCellValue("费用名称");
                row1.CreateCell(2).SetCellValue("作业类别");
                row1.CreateCell(3).SetCellValue("金额");
                row1.CreateCell(4).SetCellValue("状态");
                row1.CreateCell(5).SetCellValue("客户简称");
                row1.CreateCell(6).SetCellValue("工厂名称");
                row1.CreateCell(7).SetCellValue("工厂地址");
                row1.CreateCell(8).SetCellValue("工厂联系人");
                row1.CreateCell(9).SetCellValue("工厂联系电话");
                row1.CreateCell(10).SetCellValue("到厂时间");
                row1.CreateCell(11).SetCellValue("箱型");
                row1.CreateCell(12).SetCellValue("运输单号");
                row1.CreateCell(13).SetCellValue("订舱号S / 0 ");
                row1.CreateCell(14).SetCellValue("封条号");
                row1.CreateCell(15).SetCellValue("航次");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].SettleAccounts);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ExpenseName);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].WorkCategory);
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].UnitPrice.ToString());
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].EtrustType);
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(6).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].ClientSite);
                    rowTemp.CreateCell(8).SetCellValue(listExaminee[i].ContactsName);
                    rowTemp.CreateCell(9).SetCellValue(listExaminee[i].ContactsPhone);
                    rowTemp.CreateCell(10).SetCellValue(listExaminee[i].arriveTimes);
                    rowTemp.CreateCell(11).SetCellValue(listExaminee[i].CabinetType);
                    rowTemp.CreateCell(12).SetCellValue(listExaminee[i].CarriageNumber);
                    rowTemp.CreateCell(13).SetCellValue(listExaminee[i].BookingSpace);
                    rowTemp.CreateCell(14).SetCellValue(listExaminee[i].Seal);
                    rowTemp.CreateCell(15).SetCellValue(listExaminee[i].HangCi);
                }
                //输出的文件名称
                string fileName = "应收运费明细报表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (W == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listSelectStatement.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("应付运费明细报表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                row1.CreateCell(0).SetCellValue("费用类型");
                row1.CreateCell(1).SetCellValue("费用名称");
                row1.CreateCell(2).SetCellValue("作业类别");
                row1.CreateCell(3).SetCellValue("金额");
                row1.CreateCell(4).SetCellValue("状态");
                row1.CreateCell(5).SetCellValue("客户简称");
                row1.CreateCell(6).SetCellValue("工厂名称");
                row1.CreateCell(7).SetCellValue("工厂地址");
                row1.CreateCell(8).SetCellValue("工厂联系人");
                row1.CreateCell(9).SetCellValue("工厂联系电话");
                row1.CreateCell(10).SetCellValue("到厂时间");
                row1.CreateCell(11).SetCellValue("箱型");
                row1.CreateCell(12).SetCellValue("运输单号");
                row1.CreateCell(13).SetCellValue("订舱号S / 0 ");
                row1.CreateCell(14).SetCellValue("封条号");
                row1.CreateCell(15).SetCellValue("航次");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].SettleAccounts);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ExpenseName);
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].WorkCategory);
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].UnitPrice.ToString());
                    rowTemp.CreateCell(4).SetCellValue(listExaminee[i].EtrustType);
                    rowTemp.CreateCell(5).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(6).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(7).SetCellValue(listExaminee[i].ClientSite);
                    rowTemp.CreateCell(8).SetCellValue(listExaminee[i].ContactsName);
                    rowTemp.CreateCell(9).SetCellValue(listExaminee[i].ContactsPhone);
                    rowTemp.CreateCell(10).SetCellValue(listExaminee[i].arriveTimes);
                    rowTemp.CreateCell(11).SetCellValue(listExaminee[i].CabinetType);
                    rowTemp.CreateCell(12).SetCellValue(listExaminee[i].CarriageNumber);
                    rowTemp.CreateCell(13).SetCellValue(listExaminee[i].BookingSpace);
                    rowTemp.CreateCell(14).SetCellValue(listExaminee[i].Seal);
                    rowTemp.CreateCell(15).SetCellValue(listExaminee[i].HangCi);
                }
                //输出的文件名称
                string fileName = "应付运费明细报表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (V1 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(listSelectStatement);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.AccountsReceivableDetailFreight rp = new Print.AccountsReceivableDetailFreight();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\AccountsReceivableDetailFreight.rpt";//  \转义字符
                                                                              //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            int totalRow = listSelectStatement.Count();
            List<EtrustVo> list = listSelectStatement
                                        .Skip(BsgridPage.GetStartIndex())
                                        .Take(BsgridPage.pageSize)
                                        .ToList();
            Bsgrid<EtrustVo> bsgrid = new Bsgrid<EtrustVo>()
            {
                success = true,
                totalRows = totalRow,
                curPage = BsgridPage.curPage,
                data = list
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        //查询应收付运费总汇
        public ActionResult SelectSummaryFreightReceivable(BsgridPage BsgridPage,bool? SummaryFreightReceivable,bool? SummaryFreightPayable,string ArriveTime,string ArriveTimes,int? ClientID, int? EtrustID,int? MessageID,bool? V,bool? W,bool? V1)
        {
            var listsummaryFreightReceivable = (from tbEtrust in mymodels.SYS_Etrust
                                            join tbCharge in mymodels.SYS_Charge on tbEtrust.EtrustID equals tbCharge.EtrustID
                                            join tbExpense in mymodels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                                            join tbClient in mymodels.SYS_Client on tbEtrust.ClientCode equals  tbClient.ClientCode
                                                select new EtrustVo
                                                {
                                                   SettleAccounts=tbExpense.SettleAccounts.Trim(),
                                                   UnitPrice=tbCharge.UnitPrice,
                                                   ClientAbbreviation=tbClient.ClientAbbreviation.Trim(),
                                                    ClientCode=tbEtrust.ClientCode.Trim(),
                                                    arriveTime = tbEtrust.ArriveTime.ToString(),
                                                    ArriveTime = tbEtrust.ArriveTime,
                                                    EtrustID=tbEtrust.EtrustID,
                                                    MessageID=tbEtrust.MessageID,
                                                    ClientID = tbClient.ClientID,
                                                    J =0,
                                                   S=0
                                                }).ToList();
            if (SummaryFreightReceivable == true)
            {
                listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.SettleAccounts == "应收").ToList();
            }
            if (SummaryFreightPayable == true)
            {
                listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.SettleAccounts == "应付").ToList();
            }
            if (ClientID>0)
            {
                listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.ClientID == ClientID).ToList();
            }
            if (MessageID > 0)
            {
                listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(ArriveTime) && string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveTime);
                    listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.ArriveTime == Time).ToList();
                }
                catch (Exception)
                {
                    listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(ArriveTimes);
                    listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.ArriveTime <= Times).ToList();
                }
                catch (Exception)
                {
                    listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(ArriveTime);
                    DateTime Times = Convert.ToDateTime(ArriveTimes);
                    listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.ArriveTime >= Time && m.ArriveTime <= Times).ToList();
                }
                catch (Exception)
                {
                    listsummaryFreightReceivable = listsummaryFreightReceivable.Where(m => m.EtrustID == EtrustID).ToList();
                }
            }
            if (listsummaryFreightReceivable.Count() > 0)
            {
                var A = mymodels.SYS_Etrust.Where(m => m.ClientCode!="").Select(m => m.ClientCode.Trim()).ToList().Distinct().ToArray();
                for (int i = 0; i < A.Count(); i++)
                {
                    var C = listsummaryFreightReceivable.Where(m => m.ClientCode == A[i]).ToList();
                    var J = C.Count();
                    var S = C.Sum(m => m.UnitPrice + m.UnitPrice);
                    for (int n = 0; n < C.Count(); n++)
                    {
                        C[n].J = J;
                        C[n].S = S;
                    }
                    for (int m = 0; m < (C.Count() - 1); m++)
                    {
                        listsummaryFreightReceivable.Remove(C[m]);
                    }
                }
            }

            if (W == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listsummaryFreightReceivable.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("应付运费汇总报表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                    //给标题的每一个单元格赋值
                    row1.CreateCell(0).SetCellValue("客户");
                row1.CreateCell(1).SetCellValue("做柜数量");
                row1.CreateCell(2).SetCellValue("运费");
                row1.CreateCell(3).SetCellValue("合计");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].J.ToString());
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].S.ToString());
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].S.ToString());
                }
                //输出的文件名称
                string fileName = "应付运费汇总报表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (V == true)
            {
                //查询数据
                List<EtrustVo> listExaminee = listsummaryFreightReceivable.ToList();
                //二：代码创建一个Excel表格（这里称为工作簿）
                //创建Excel文件的对象 工作簿(调用NPOI文件)
                HSSFWorkbook excelBook = new HSSFWorkbook();
                //创建Excel工作表 Sheet=考生信息
                ISheet sheet1 = excelBook.CreateSheet("应收运费汇总报表");
                //给Sheet（考生信息）添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                //给标题的每一个单元格赋值
                row1.CreateCell(0).SetCellValue("客户");
                row1.CreateCell(1).SetCellValue("做柜数量");
                row1.CreateCell(2).SetCellValue("运费");
                row1.CreateCell(3).SetCellValue("合计");
                for (int i = 0; i < listExaminee.Count; i++)
                {
                    IRow rowTemp = sheet1.CreateRow(i + 1);
                    rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ClientAbbreviation);
                    rowTemp.CreateCell(1).SetCellValue(listExaminee[i].J.ToString());
                    rowTemp.CreateCell(2).SetCellValue(listExaminee[i].S.ToString());
                    rowTemp.CreateCell(3).SetCellValue(listExaminee[i].S.ToString());
                }
                //输出的文件名称
                string fileName = "应收运费汇总报表" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
                //把Excel转为流，输出
                //创建文件流
                System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
                //将工作薄写入文件流
                excelBook.Write(bookStream);

                //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
                bookStream.Seek(0, System.IO.SeekOrigin.Begin);
                //Stream对象,文件类型,文件名称
                return File(bookStream, "application/vnd.ms-excel", fileName);
            }
            if (V1 == true)
            {
                //把linq类型的数据listResult转化为DataTable类型数据
                DataTable dt = LINQToDataTable(listsummaryFreightReceivable);
                //第一步：实例化数据集
                Print.Etrust dbReport = new Print.Etrust();
                //第二步：将dt的数据放入数据集的数据表中
                dbReport.Tables["Statement"].Merge(dt);
                //第三步：实例化报表模板
                Print.SummaryFreightReceivable rp = new Print.SummaryFreightReceivable();
                //第四步：获取报表物理文件地址     
                string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                    + "Areas\\Forwarding\\Print\\SummaryFreightReceivable.rpt";//  \转义字符
                                                                                      //第五步：把报表文件加载到ReportDocument
                rp.Load(strRptPath);
                //第六步：设置报表数据源
                rp.SetDataSource(dbReport);
                //第七步：把ReportDocument转化为文件流
                Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                return File(stream, "application/pdf");
            }
            int totalRow = listsummaryFreightReceivable.Count();
            List<EtrustVo> list = listsummaryFreightReceivable
                                        .Skip(BsgridPage.GetStartIndex())
                                        .Take(BsgridPage.pageSize)
                                        .ToList();
            Bsgrid<EtrustVo> bsgrid = new Bsgrid<EtrustVo>()
            {
                success = true,
                totalRows = totalRow,
                curPage = BsgridPage.curPage,
                data = list
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        //做柜数量
        public ActionResult SelectZuoGui() {
            var ZGS = mymodels.SYS_Etrust.Where(m => m.VehicleInformationID > 0).Select(m => m.VehicleInformationID).Distinct().Count();
            return Json(ZGS, JsonRequestBehavior.AllowGet);
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型 
        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //将类型的属性名称与属性类型作为DataTable的列数据
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
    }
}