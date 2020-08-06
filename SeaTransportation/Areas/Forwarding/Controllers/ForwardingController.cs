using Hotel.VO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SeaTransportation.Models;
using SeaTransportation.Vo;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Mvc;

namespace SeaTransportation.Areas.Forwarding.Controllers
{
    public class ForwardingController : Controller
    {
        Models.SeaTransportationEntities mymodels = new Models.SeaTransportationEntities();
        // GET: Forwarding/Forwarding

        #region 视图
        //Forwarding
        public ActionResult Forwarding()
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
        #endregion

        #region 查询
        //委托单查询
        public ActionResult SelectEtrust(BsgridPage BsgridPage,bool? Carriage,int? EtrustID,string EtrustNmber,string Seal,string ArriveTime,string ArriveTimes,int? ClientAbbreviation,int? TiGuiDiDianID,int? HuaiGuiDiDianID,string Cupboard,bool? Update,bool? EtrustDerive,bool? CarriageDerive,bool? ChagreDerive,bool? EtrustPrint,bool? tbetrust,int? A,bool? A1)
        {
            try
            {
                #region 委托
                if (Carriage == true)
                {
                    var Staff = Session["StaffName"].ToString();
                    var listEtrust = (from tbEtrust in mymodels.SYS_Etrust
                                      join tbClient in mymodels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                                      into tbClient from Client in tbClient.DefaultIfEmpty()
                                      join tbClientScontacts in mymodels.SYS_ClientScontacts on tbEtrust.ClientScontactsID equals tbClientScontacts.ClientScontactsID
                                      into tbClientScontacts from ClientScontacts in tbClientScontacts.DefaultIfEmpty()
                                      join tbClientSite in mymodels.SYS_ClientSite on tbEtrust.ClientSiteID equals tbClientSite.ClientSiteID
                                      into tbClientSite from ClientSite in tbClientSite.DefaultIfEmpty()
                                      join tbClientType in mymodels.SYS_ClientType on tbEtrust.UndertakeID equals tbClientType.ClientTypeID
                                      into tbClientType from ClientType in tbClientType.DefaultIfEmpty()
                                      join tbClients in mymodels.SYS_Client on ClientType.ClientID equals tbClients.ClientID
                                      into tbClients from Clients in tbClients.DefaultIfEmpty()
                                      join tbVehicleInformation in mymodels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicleInformation.VehicleInformationID 
                                      into tbVehicleInformation  from VehicleInformation in tbVehicleInformation.DefaultIfEmpty()
                                      join tbChauffeur in mymodels.SYS_Chauffeur on VehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                                      into tbChauffeur from Chauffeur in tbChauffeur.DefaultIfEmpty()
                                      join tbBracket in mymodels.SYS_Bracket on VehicleInformation.BracketID equals tbBracket.BracketID
                                      into tbBracket from Bracket in tbBracket.DefaultIfEmpty()
                                      join tbMention in mymodels.SYS_Mention on tbEtrust.TiGuiDiDianID equals tbMention.MentionID
                                      into tbMention from Mention in tbMention.DefaultIfEmpty()
                                      join tbMentions in mymodels.SYS_Mention on tbEtrust.HuaiGuiDiDianID equals tbMentions.MentionID
                                      into tbMentions from Mentions in tbMentions.DefaultIfEmpty()
                                      join tbGatedot in mymodels.SYS_Gatedot on Mention.GatedotID equals tbGatedot.GatedotID
                                      into tbGatedot from Gatedot in tbGatedot.DefaultIfEmpty()
                                      join tbGatedots in mymodels.SYS_Gatedot on Mentions.GatedotID equals tbGatedots.GatedotID
                                      into tbGatedots from Gatedots in tbGatedots.DefaultIfEmpty()
                                      join tbShip in mymodels.SYS_Ship on tbEtrust.ShipID equals tbShip.ShipID
                                      into tbShip from Ship in tbShip.DefaultIfEmpty()
                                      join tbClinetTypes in mymodels.SYS_ClientType on Ship.ClientTypeID equals tbClinetTypes.ClientTypeID
                                      into tbClinetTypes from ClinetTypes in tbClinetTypes.DefaultIfEmpty()
                                      join tbClientl in mymodels.SYS_Client on ClinetTypes.ClientID equals tbClientl.ClientID
                                      into tbClientl from Clientl in tbClientl.DefaultIfEmpty()
                                      join tbGatedotl in mymodels.SYS_Gatedot on ClientSite.GatedotID equals tbGatedotl.GatedotID
                                      into tbGatedotl from Gatedotl in tbGatedotl.DefaultIfEmpty()
                                      join tbStaff in mymodels.SYS_Staff on tbEtrust.StaffID equals tbStaff.StaffID
                                      into tbStaff from staff in tbStaff.DefaultIfEmpty()
                                      join tbMessage in mymodels.SYS_Message on tbEtrust.MessageID equals tbMessage.MessageID
                                      into tbMessage from Message in tbMessage.DefaultIfEmpty()
                                      orderby tbEtrust.EtrustID  descending
                                      select new EtrustVo
                                      {
                                          MessageName=Message.ChineseName.Trim(),
                                          AssistBusiness=tbEtrust.AssistBusiness.Trim(),
                                          OddNumber=tbEtrust.OddNumber.Trim(),
                                          EtrustID = tbEtrust.EtrustID,
                                          Undertake = Clients.ClientAbbreviation,
                                          EtrustNmber = tbEtrust.EtrustNmber.Trim(),
                                          WorkCategory = tbEtrust.WorkCategory.Trim(),
                                          CheXing = tbEtrust.CheXing,
                                          VehicleCode = VehicleInformation.VehicleCode,
                                          ChauffeurName = Chauffeur.ChauffeurName,
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
                                          Staff = Staff.Trim(),
                                          KaiCangTime=tbEtrust.KaiCangTime,
                                          kaiCangTime = tbEtrust.KaiCangTime.ToString(),
                                          StaffName =staff.StaffName.Trim(),
                                          TG=Mention.Abbreviation.Trim(),
                                          HG = Mentions.Abbreviation.Trim(),
                                          OfferDetailID = tbEtrust.OfferDetailID,
                                          YSLX= Gatedot.GatedotName.Trim() + " - " + Gatedotl.GatedotName.Trim() + " - "+ Gatedots.GatedotName.Trim(),
                                          VehicleInformationID = tbEtrust.VehicleInformationID,
                                          Shij = Chauffeur.ChauffeurName.Trim(),
                                          ICNumber = Chauffeur.ICNumber.Trim(),
                                          PhoneOne = Chauffeur.PhoneOne.Trim(),
                                          BracketType = Bracket.BracketType.Trim(),
                                          UndertakeID=tbEtrust.UndertakeID,
                                          paiCheDate=tbEtrust.PaiCheDate.ToString(),
                                          FollowExplain=tbEtrust.FollowExplain.Trim()
                                      }).ToList();
                    if (EtrustID >= 0)
                    {
                        listEtrust = listEtrust.Where(m => m.EtrustID == EtrustID).ToList();
                    }
                    if (Update==true)
                    {
                        return Json(listEtrust, JsonRequestBehavior.AllowGet);
                    }
                    if (ClientAbbreviation >0)
                    {
                        listEtrust = listEtrust.Where(m => m.ClientID == ClientAbbreviation).ToList();
                    }
                    if (TiGuiDiDianID > 0)
                    {
                        listEtrust = listEtrust.Where(m => m.TiGuiDiDianID == TiGuiDiDianID).ToList();
                    }
                    if (HuaiGuiDiDianID > 0)
                    {
                        listEtrust = listEtrust.Where(m => m.HuaiGuiDiDianID == HuaiGuiDiDianID).ToList();
                    }
                    if (!string.IsNullOrEmpty(EtrustNmber))
                    {
                        listEtrust = listEtrust.Where(m => m.EtrustNmber.Contains(EtrustNmber)).ToList();
                    }
                    if (!string.IsNullOrEmpty(Seal))
                    {
                        listEtrust = listEtrust.Where(m => m.Seal.Contains(Seal)).ToList();
                    }
                    if (!string.IsNullOrEmpty(Cupboard))
                    {
                        listEtrust = listEtrust.Where(m => m.Cupboard.Contains(Cupboard)).ToList();
                    }
                    if (!string.IsNullOrEmpty(ArriveTime) && string.IsNullOrEmpty(ArriveTimes))
                    {
                        try
                        {
                            DateTime Time = Convert.ToDateTime(ArriveTime);
                            listEtrust = listEtrust.Where(m => m.ArriveTime == Time).ToList();
                        }
                        catch (Exception)
                        {
                            listEtrust = listEtrust.Where(m => m.EtrustID == EtrustID).ToList();
                        }
                    }
                    else if (string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
                    {
                        try
                        {
                            DateTime Times = Convert.ToDateTime(ArriveTimes);
                            listEtrust = listEtrust.Where(m => m.ArriveTime <= Times).ToList();
                        }
                        catch (Exception)
                        {
                            listEtrust = listEtrust.Where(m => m.EtrustID == EtrustID).ToList();
                        }
                    }else if (!string.IsNullOrEmpty(ArriveTime) && !string.IsNullOrEmpty(ArriveTimes))
                    {
                        try
                        {
                            DateTime Time = Convert.ToDateTime(ArriveTime);
                            DateTime Times = Convert.ToDateTime(ArriveTimes);
                            listEtrust = listEtrust.Where(m => m.ArriveTime >= Time && m.ArriveTime <= Times).ToList();
                        }
                        catch (Exception)
                        {
                            listEtrust = listEtrust.Where(m => m.EtrustID == EtrustID).ToList();
                        }
                    }
                    if (EtrustDerive==true)
                    {
                        //查询数据
                        List<EtrustVo> listExaminee = listEtrust.ToList();
                        //二：代码创建一个Excel表格（这里称为工作簿）
                        //创建Excel文件的对象 工作簿(调用NPOI文件)
                        HSSFWorkbook excelBook = new HSSFWorkbook();
                        //创建Excel工作表 Sheet=考生信息
                        ISheet sheet1 = excelBook.CreateSheet("委托单信息");
                        //给Sheet（考生信息）添加第一行的头部标题
                        IRow row1 = sheet1.CreateRow(0);
                        //给标题的每一个单元格赋值
                        row1.CreateCell(0).SetCellValue("状态");
                        row1.CreateCell(1).SetCellValue("委托单号");
                        row1.CreateCell(2).SetCellValue("作业类别");
                        row1.CreateCell(3).SetCellValue("车型");
                        row1.CreateCell(4).SetCellValue("到厂时间");
                        row1.CreateCell(5).SetCellValue("客户简称");
                        row1.CreateCell(6).SetCellValue("客户联系人");
                        row1.CreateCell(7).SetCellValue("客户联系人电话");
                        row1.CreateCell(8).SetCellValue("工厂联系人");
                        row1.CreateCell(9).SetCellValue("工厂联系电话");
                        row1.CreateCell(10).SetCellValue("箱型");
                        row1.CreateCell(11).SetCellValue("订舱号S / 0");
                        row1.CreateCell(12).SetCellValue("封条号");
                        row1.CreateCell(13).SetCellValue("航次");
                        for (int i = 0; i < listExaminee.Count; i++)
                        {
                            IRow rowTemp = sheet1.CreateRow(i + 1);
                            rowTemp.CreateCell(0).SetCellValue(listExaminee[i].EtrustType);
                            rowTemp.CreateCell(1).SetCellValue(listExaminee[i].EtrustNmber);
                            rowTemp.CreateCell(2).SetCellValue(listExaminee[i].WorkCategory);
                            rowTemp.CreateCell(3).SetCellValue(listExaminee[i].CheXing);
                            rowTemp.CreateCell(4).SetCellValue(listExaminee[i].arriveTimes);
                            rowTemp.CreateCell(5).SetCellValue(listExaminee[i].ClientAbbreviation);
                            rowTemp.CreateCell(6).SetCellValue(listExaminee[i].contactsName);
                            rowTemp.CreateCell(7).SetCellValue(listExaminee[i].contactsPhone);
                            rowTemp.CreateCell(8).SetCellValue(listExaminee[i].ContactsName);
                            rowTemp.CreateCell(9).SetCellValue(listExaminee[i].ContactsPhone);
                            rowTemp.CreateCell(10).SetCellValue(listExaminee[i].CabinetType);
                            rowTemp.CreateCell(11).SetCellValue(listExaminee[i].BookingSpace);
                            rowTemp.CreateCell(12).SetCellValue(listExaminee[i].Seal);
                            rowTemp.CreateCell(13).SetCellValue(listExaminee[i].HangCi);
                        }
                        //输出的文件名称
                        string fileName = "委托单信息" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
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
                    if (CarriageDerive == true)
                    {
                        listEtrust = listEtrust.Where(m => m.EtrustType.Trim() != "新单").ToList();
                        //查询数据
                        List<EtrustVo> listExaminee = listEtrust.ToList();
                        //二：代码创建一个Excel表格（这里称为工作簿）
                        //创建Excel文件的对象 工作簿(调用NPOI文件)
                        HSSFWorkbook excelBook = new HSSFWorkbook();
                        //创建Excel工作表 Sheet=考生信息
                        ISheet sheet1 = excelBook.CreateSheet("运输单信息");
                        //给Sheet（考生信息）添加第一行的头部标题
                        IRow row1 = sheet1.CreateRow(0);
                        //给标题的每一个单元格赋值
                        row1.CreateCell(0).SetCellValue("委托单号");
                        row1.CreateCell(1).SetCellValue("运输单号");
                        row1.CreateCell(2).SetCellValue("状态");
                        row1.CreateCell(3).SetCellValue("客户简称");
                        row1.CreateCell(4).SetCellValue("车型");
                        row1.CreateCell(5).SetCellValue("承运公司");
                        row1.CreateCell(6).SetCellValue("司机姓名");
                        row1.CreateCell(7).SetCellValue("托架代码");
                        row1.CreateCell(8).SetCellValue("车辆代码");
                        for (int i = 0; i < listExaminee.Count; i++)
                        {
                            IRow rowTemp = sheet1.CreateRow(i + 1);
                            rowTemp.CreateCell(0).SetCellValue(listExaminee[i].EtrustNmber);
                            rowTemp.CreateCell(1).SetCellValue(listExaminee[i].CarriageNumber);
                            rowTemp.CreateCell(2).SetCellValue(listExaminee[i].EtrustType);
                            rowTemp.CreateCell(3).SetCellValue(listExaminee[i].ClientAbbreviation);
                            rowTemp.CreateCell(4).SetCellValue(listExaminee[i].CheXing);
                            rowTemp.CreateCell(5).SetCellValue(listExaminee[i].Undertake);
                            rowTemp.CreateCell(6).SetCellValue(listExaminee[i].ChauffeurName);
                            rowTemp.CreateCell(7).SetCellValue(listExaminee[i].BracketCode);
                            rowTemp.CreateCell(8).SetCellValue(listExaminee[i].VehicleCode);
                        }
                        //输出的文件名称
                        string fileName = "运输单信息" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
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
                    if (EtrustPrint==true)
                    {
                        //把linq类型的数据listResult转化为DataTable类型数据
                        DataTable dt = LINQToDataTable(listEtrust);
                        //第一步：实例化数据集
                        Print.Etrust dbReport = new Print.Etrust();
                        //第二步：将dt的数据放入数据集的数据表中
                        dbReport.Tables["SYS_Etrust"].Merge(dt);
                        //第三步：实例化报表模板
                        Print.Forwarding rp = new Print.Forwarding();
                        //第四步：获取报表物理文件地址     
                        string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                            + "Areas\\Forwarding\\Print\\Forwarding.rpt";//  \转义字符
                                                                                                 //第五步：把报表文件加载到ReportDocument
                        rp.Load(strRptPath);
                        //第六步：设置报表数据源
                        rp.SetDataSource(dbReport);
                        //第七步：把ReportDocument转化为文件流
                        Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                        return File(stream, "application/pdf");
                    }
                    if (tbetrust == true)
                    {
                        //把linq类型的数据listResult转化为DataTable类型数据
                        DataTable dt = LINQToDataTable(listEtrust);
                        //第一步：实例化数据集
                        Print.Etrust dbReport = new Print.Etrust();
                        //第二步：将dt的数据放入数据集的数据表中
                        dbReport.Tables["tbEtrust"].Merge(dt);
                        //第三步：实例化报表模板
                        Print.tbEtrust rp = new Print.tbEtrust();
                        //第四步：获取报表物理文件地址     
                        string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                            + "Areas\\Forwarding\\Print\\tbEtrust.rpt";//  \转义字符
                                                                         //第五步：把报表文件加载到ReportDocument
                        rp.Load(strRptPath);
                        //第六步：设置报表数据源
                        rp.SetDataSource(dbReport);
                        //第七步：把ReportDocument转化为文件流
                        Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                        return File(stream, "application/pdf");
                    }
                    if (A1 == true)
                    {
                        listEtrust = listEtrust.Where(m => m.EtrustID == A).ToList();
                        //把linq类型的数据listResult转化为DataTable类型数据
                        DataTable dt = LINQToDataTable(listEtrust);
                        //第一步：实例化数据集
                        Print.Etrust dbReport = new Print.Etrust();
                        //第二步：将dt的数据放入数据集的数据表中
                        dbReport.Tables["tbCarriage"].Merge(dt);
                        //第三步：实例化报表模板
                        Print.tbCarriage rp = new Print.tbCarriage();
                        //第四步：获取报表物理文件地址     
                        string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                            + "Areas\\Forwarding\\Print\\tbCarriage.rpt";//  \转义字符
                                                                       //第五步：把报表文件加载到ReportDocument
                        rp.Load(strRptPath);
                        //第六步：设置报表数据源
                        rp.SetDataSource(dbReport);
                        //第七步：把ReportDocument转化为文件流
                        Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                        return File(stream, "application/pdf");
                    }
                    int totalRow = listEtrust.Count();
                    List<EtrustVo> list = listEtrust
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
                else
                #endregion
                #region 收费
                if (Carriage == false)
                {
                    var listEtrust = (from tbEtrust in mymodels.SYS_Etrust
                                      join tbClient in mymodels.SYS_Client on tbEtrust.ClientCode equals tbClient.ClientCode
                                      into tbClient from Client in tbClient.DefaultIfEmpty()
                                      join tbClientScontacts in mymodels.SYS_ClientScontacts on tbEtrust.ClientScontactsID equals tbClientScontacts.ClientScontactsID
                                      into tbClientScontacts from ClientScontacts in tbClientScontacts.DefaultIfEmpty()
                                      join tbClientSite in mymodels.SYS_ClientSite on tbEtrust.ClientSiteID equals tbClientSite.ClientSiteID
                                      into tbClientSite from ClientSite in tbClientSite.DefaultIfEmpty()
                                      join tbClientType in mymodels.SYS_ClientType on tbEtrust.UndertakeID equals tbClientType.ClientTypeID
                                      into tbClientType from ClientType in tbClientType.DefaultIfEmpty()
                                      join tbClients in mymodels.SYS_Client on ClientType.ClientID equals tbClients.ClientID
                                      into tbClients from Clients in tbClients.DefaultIfEmpty()
                                      join tbVehicleInformation in mymodels.SYS_VehicleInformation on tbEtrust.VehicleInformationID equals tbVehicleInformation.VehicleInformationID
                                      into tbVehicleInformation from VehicleInformation in tbVehicleInformation.DefaultIfEmpty()
                                      join tbChauffeur in mymodels.SYS_Chauffeur on VehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                                      into tbChauffeur from Chauffeur in tbChauffeur.DefaultIfEmpty()
                                      join tbBracket in mymodels.SYS_Bracket on VehicleInformation.BracketID equals tbBracket.BracketID
                                      into tbBracket from Bracket in tbBracket.DefaultIfEmpty()
                                      join tbCharge in mymodels.SYS_Charge on tbEtrust.EtrustID equals tbCharge.EtrustID
                                      into tbCharge from Charge in tbCharge.DefaultIfEmpty()
                                      join tbExpense in mymodels.SYS_Expense on Charge.ExpenseID equals tbExpense.ExpenseID
                                      join tbOfferDetail in mymodels.SYS_OfferDetail on tbEtrust.OfferDetailID equals tbOfferDetail.OfferDetailID
                                      into tbOfferDetail from OfferDetail in tbOfferDetail.DefaultIfEmpty()
                                      orderby tbEtrust.EtrustID descending
                                      select new EtrustVo
                                      {
                                          VehicleInformationID = tbEtrust.VehicleInformationID,
                                          AssistBusiness = tbEtrust.AssistBusiness.Trim(),
                                          ClientID = Client.ClientID,
                                          EtrustID = tbEtrust.EtrustID,
                                          Undertake = Clients.ClientAbbreviation,
                                          EtrustNmber = tbEtrust.EtrustNmber.Trim(),
                                          WorkCategory = tbEtrust.WorkCategory,
                                          CheXing = tbEtrust.CheXing.Trim(),
                                          VehicleCode = VehicleInformation.VehicleCode.Trim(),
                                          ChauffeurName = Chauffeur.ChauffeurName.Trim(),
                                          ChauffeurNumber = Chauffeur.ChauffeurNumber.Trim(),
                                          BracketCode = Bracket.BracketCode.Trim(),
                                          arriveTime = tbEtrust.ArriveTime.ToString(),
                                          HangCi = tbEtrust.HangCi.Trim(),
                                          CabinetType = tbEtrust.CabinetType.Trim(),
                                          BookingSpace = tbEtrust.BookingSpace,
                                          Seal = tbEtrust.Seal.Trim(),
                                          CarriageNumber = tbEtrust.CarriageNumber.Trim(),
                                          EtrustType = tbEtrust.EtrustType.Trim(),
                                          ClientAbbreviation = Client.ClientAbbreviation.Trim(),
                                          ContactsName = ClientSite.ContactsName.Trim(),
                                          ContactsPhone = ClientSite.ContactsPhone.Trim(),
                                          contactsName = ClientScontacts.contactsName.Trim(),
                                          contactsPhone = ClientScontacts.contactsPhone.Trim(),
                                          ExpenseName = tbExpense.ExpenseName.Trim(),
                                          SettleAccounts = tbExpense.SettleAccounts.Trim(),
                                          UnitPrice = Charge.UnitPrice,
                                          BoxQuantity = "",
                                          Currency = "RMB",
                                          ReckoningUnit = Charge.ReckoningUnit.Trim(),
                                          ReckoningUnits = Charge.ReckoningUnits.Trim(),
                                          Reckoning = "",
                                          AuditType = tbEtrust.AuditType.Trim(),
                                          ClientCode=tbEtrust.ClientCode.Trim(),
                                          ChineseName = Client.ChineseName.Trim(),
                                          ClientScontactsID=tbEtrust.ClientScontactsID,
                                          ClientMobile=ClientScontacts.contactsPhone.Trim(),
                                          PlanHandTime =tbEtrust.PlanHandTime,
                                          planHandTime=tbEtrust.PlanHandTime.ToString(),
                                          ClientSiteID=tbEtrust.ClientSiteID,
                                          arriveTimes=tbEtrust.ArriveTime.ToString()
                                      }).ToList();
                    if (EtrustID != null)
                    {
                        listEtrust = listEtrust.Where(m => m.EtrustID == EtrustID).ToList();
                    }
                    if (listEtrust.Count() > 0)
                    {

                        for (int i = 0; i < listEtrust.Count(); i++)
                        {
                            var ReckoningUnit = listEtrust[i].ReckoningUnit;
                            var ReckoningUnits = listEtrust[i].ReckoningUnits;
                            if (!string.IsNullOrEmpty(ReckoningUnit))
                            {
                                try
                                {
                                    listEtrust[i].Reckoning = mymodels.SYS_Message.Where(m => m.TissueCode.Trim() == ReckoningUnit).Select(m => m.ChineseAbbreviation).Single();
                                }
                                catch (Exception)
                                {
                                    listEtrust[i].Reckoning = "";
                                }
                            }
                            else if (!string.IsNullOrEmpty(ReckoningUnits))
                            {
                                try
                                {
                                    listEtrust[i].Reckoning = mymodels.SYS_Client.Where(m => m.ClientCode.Trim() == ReckoningUnits).Select(m => m.ClientAbbreviation).Single();
                                }
                                catch (Exception)
                                {
                                    listEtrust[i].Reckoning = "";
                                }
                            }
                        }
                    }
                    if (listEtrust.Count() > 0)
                    {
                        for (int i = 0; i < listEtrust.Count(); i++)
                        {
                            if (listEtrust[i].CabinetType.Trim()=="20GP")
                            {
                                listEtrust[i].BoxQuantity = "1000";
                            }
                            else if (listEtrust[i].CabinetType.Trim() == "40GP")
                            {
                                listEtrust[i].BoxQuantity = "2000";
                            }
                            else if (listEtrust[i].CabinetType.Trim() == "60GP")
                            {
                                listEtrust[i].BoxQuantity = "3000";
                            }
                            else if (listEtrust[i].CabinetType.Trim() == "80GP")
                            {
                                listEtrust[i].BoxQuantity = "4000";
                            }
                            else if (listEtrust[i].CabinetType.Trim() == "100GP")
                            {
                                listEtrust[i].BoxQuantity = "5000";
                            }
                        }
                    }
                    if (ChagreDerive==true)
                    {
                        //查询数据
                        List<EtrustVo> listExaminee = listEtrust.ToList();
                        //二：代码创建一个Excel表格（这里称为工作簿）
                        //创建Excel文件的对象 工作簿(调用NPOI文件)
                        HSSFWorkbook excelBook = new HSSFWorkbook();
                        //创建Excel工作表 Sheet=考生信息
                        ISheet sheet1 = excelBook.CreateSheet("委托单信息");
                        //给Sheet（考生信息）添加第一行的头部标题
                        IRow row1 = sheet1.CreateRow(0);
                        //给标题的每一个单元格赋值
                        row1.CreateCell(0).SetCellValue("委托单号");
                        row1.CreateCell(1).SetCellValue("费用名称");
                        row1.CreateCell(2).SetCellValue("收付款类型");
                        row1.CreateCell(3).SetCellValue("结算单位");
                        row1.CreateCell(4).SetCellValue("单价");
                        row1.CreateCell(5).SetCellValue("箱型");
                        row1.CreateCell(6).SetCellValue("箱量");
                        row1.CreateCell(7).SetCellValue("金额");
                        row1.CreateCell(8).SetCellValue("币种");
                        for (int i = 0; i < listExaminee.Count; i++)
                        {
                            IRow rowTemp = sheet1.CreateRow(i + 1);
                            rowTemp.CreateCell(0).SetCellValue(listExaminee[i].EtrustNmber);
                            rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ExpenseName);
                            rowTemp.CreateCell(2).SetCellValue(listExaminee[i].SettleAccounts);
                            rowTemp.CreateCell(3).SetCellValue(listExaminee[i].Reckoning);
                            rowTemp.CreateCell(4).SetCellValue(Convert.ToDouble(listExaminee[i].UnitPrice));
                            rowTemp.CreateCell(5).SetCellValue(listExaminee[i].CabinetType);
                            rowTemp.CreateCell(6).SetCellValue(listExaminee[i].BoxQuantity);
                            rowTemp.CreateCell(7).SetCellValue(Convert.ToDouble(listExaminee[i].UnitPrice));
                            rowTemp.CreateCell(8).SetCellValue(listExaminee[i].Currency);
                        }
                        //输出的文件名称
                        string fileName = "委托单信息" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
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
                    if (A1 == true)
                    {
                        listEtrust = listEtrust.Where(m => m.EtrustID == A).ToList();
                        //把linq类型的数据listResult转化为DataTable类型数据
                        DataTable dt = LINQToDataTable(listEtrust);
                        //第一步：实例化数据集
                        Print.Etrust dbReport = new Print.Etrust();
                        //第二步：将dt的数据放入数据集的数据表中
                        dbReport.Tables["tbCharge"].Merge(dt);
                        //第三步：实例化报表模板
                        Print.tbCharge rp = new Print.tbCharge();
                        //第四步：获取报表物理文件地址     
                        string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                            + "Areas\\Forwarding\\Print\\tbCharge.rpt";//  \转义字符
                                                                         //第五步：把报表文件加载到ReportDocument
                        rp.Load(strRptPath);
                        //第六步：设置报表数据源
                        rp.SetDataSource(dbReport);
                        //第七步：把ReportDocument转化为文件流
                        Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                        return File(stream, "application/pdf");
                    }
                    int totalRow = listEtrust.Count();
                    List<EtrustVo> list = listEtrust
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
                #endregion
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
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
        //费用项目查询
        public ActionResult SelectExpense(BsgridPage BsgridPage,int? ExpenseID) {
            var Expense = (mymodels.SYS_Expense).ToList();
            int totalRow = Expense.Count();
            List<SYS_Expense> list = Expense
                                        .Skip(BsgridPage.GetStartIndex())
                                        .Take(BsgridPage.pageSize)
                                        .ToList();
            Bsgrid<SYS_Expense> bsgrid = new Bsgrid<SYS_Expense>()
            {
                success = true,
                totalRows = totalRow,
                curPage = BsgridPage.curPage,
                data = list
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        //客户查询
        public ActionResult SelectClient(int? ClientID) {
            var Client = mymodels.SYS_Client.Where(m => m.ClientID == ClientID).Select(m => m.ChineseName.Trim()).ToList();
            return Json(Client, JsonRequestBehavior.AllowGet);
        }
        //客户联系人查询
        public ActionResult SelectClientScontacts(int? ClientScontactsID)
        {
            var ClientScontacts = mymodels.SYS_ClientScontacts.Where(m => m.ClientScontactsID == ClientScontactsID).Select(m => m.contactsPhone.Trim()).ToList();
            return Json(ClientScontacts, JsonRequestBehavior.AllowGet);
        }
        //客户地址查询
        public ActionResult SelectClientSite(int? ClientSiteID)
        {
            var Factory = (from tbClientSite in mymodels.SYS_ClientSite
                          join tbGatedot in mymodels.SYS_Gatedot on tbClientSite.GatedotID equals tbGatedot.GatedotID
                          where tbClientSite.ClientSiteID == ClientSiteID
                          select new
                          {
                            FactoryName=tbClientSite.FactoryName.Trim(),
                            ClientSite = tbClientSite.ClientSite.Trim(),
                            GatedotName=tbGatedot.GatedotName.Trim(),
                            ContactsName = tbClientSite.ContactsName.Trim(),
                            ContactsPhone = tbClientSite.ContactsPhone.Trim(),
                          }).ToList();
            return Json(Factory, JsonRequestBehavior.AllowGet);
        }
        //门点查询
        public ActionResult SelectMention(int? MentionID)
        {
            var GatedotName = (from tbMention in mymodels.SYS_Mention
                          join tbGatedot in mymodels.SYS_Gatedot on tbMention.GatedotID equals tbGatedot.GatedotID
                          where tbMention.MentionID == MentionID
                          select new
                          {
                              GatedotName = tbGatedot.GatedotName.Trim(),
                              Abbreviation = tbMention.Abbreviation.Trim()
                          }).ToList();
            return Json(GatedotName, JsonRequestBehavior.AllowGet);
        }
        //船公司查询
        public ActionResult SelectShip(int? ShipID)
        {
            var ChineseName = (from tbShip in mymodels.SYS_Ship
                          join tbClientType in mymodels.SYS_ClientType on tbShip.ClientTypeID equals tbClientType.ClientTypeID
                          join tbClient in mymodels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                          where tbShip.ShipID== ShipID
                          select new
                          {
                              ChineseName = tbClient.ChineseName.Trim()
                          }).ToList();
            return Json(ChineseName, JsonRequestBehavior.AllowGet);
        }
          //派车查询
        public ActionResult SelectSendCar(int? VehicleInformationID)
        {
            var SendCar = (from tbVehicleInformation in mymodels.SYS_VehicleInformation
                               join tbChauffeur in mymodels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                               join tbBracket in mymodels.SYS_Bracket on tbVehicleInformation.BracketID equals tbBracket.BracketID
                               where tbVehicleInformation.VehicleInformationID == VehicleInformationID
                               select new
                               {
                                   BracketCode = tbBracket.BracketCode.Trim(),
                                   Shij = tbChauffeur.ChauffeurName.Trim(),
                                   ICNumber = tbChauffeur.ICNumber.Trim(),
                                   PhoneOne = tbChauffeur.PhoneOne.Trim(),
                                   BracketType = tbBracket.BracketType.Trim()
                               }).ToList();
            return Json(SendCar, JsonRequestBehavior.AllowGet);
        }
        //费用项目查询
        public ActionResult chagre(int? EtrustID,int? ExpenseID) {
            try
            {
                var mychagre = mymodels.SYS_Charge.Where(m => m.EtrustID == EtrustID && m.ExpenseID == ExpenseID).ToList();
                if (mychagre.Count()>0)
                {
                    if (mychagre[0].UnitPrice > 0)
                    {
                        var Money = mychagre[0].UnitPrice;
                        return Json(Money, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        var Money = 0;
                        return Json(Money, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                   var  Money = 0;
                    return Json(Money, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
               
                throw;
            }
          
        }
        #endregion

        #region 新增
        //新增委托单
        public ActionResult InsertEtrust(SYS_Etrust SYS_Etrust,int? ClientID)
        {
            var Data = "";
            try
            {
                SYS_Etrust.StaffID = Convert.ToInt32(Session["StaffID"].ToString());
                SYS_Etrust.ClientCode = mymodels.SYS_Client.Where(m => m.ClientID == ClientID).Select(m => m.ClientCode.Trim()).Single();
                SYS_Etrust.CheXing = "自备柜";
                SYS_Etrust.AuditType = "未审核";
                SYS_Etrust.EtrustType = "新单";
                if (mymodels.SYS_Etrust.Where(M=>M.EtrustNmber==SYS_Etrust.EtrustNmber).Count()==0)
                {
                        var S = false;
                        if (SYS_Etrust.ShuiBl > 0)
                        {
                             S = true;
                        }
                        var GatedotID = mymodels.SYS_ClientSite.Where(m => m.ClientSiteID == SYS_Etrust.ClientSiteID).Select(m => m.GatedotID).Single();
                       var OfferCount = (from tbOfferDetail in mymodels.SYS_OfferDetail
                                         join tbHaulWay in mymodels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                         join tbOffer in mymodels.SYS_Offer on tbOfferDetail.OfferID equals tbOffer.OfferID
                                         where tbHaulWay.MentionAreaID == SYS_Etrust.TiGuiDiDianID && tbHaulWay.AlsoTankAreaID == SYS_Etrust.HuaiGuiDiDianID && tbHaulWay.LoadingAreaID == GatedotID &&
                                         tbOffer.WhetherShui == S && tbOfferDetail.CabinetType == SYS_Etrust.CabinetType
                                         where tbOffer.OfferType.Trim() == "客户标准运费" || tbOffer.OfferType.Trim() == "客户应收费用"
                                         select new  EtrustVo{
                                            money=tbOfferDetail.Money,
                                            clientID=tbOffer.ClientID,
                                            offerDetailID=tbOfferDetail.OfferDetailID,
                                             OfferType=tbOffer.OfferType.Trim()
                                         }).ToList();
                        if (OfferCount.Count()>0)
                        {
                            var offercount = OfferCount.Where(m => m.clientID == ClientID).ToList();
                            if (offercount.Count()>0)
                            {
                                    SYS_Etrust.OfferDetailID = offercount[0].offerDetailID;
                                    if (SYS_Etrust.CollectionMoney > 0)
                                    {
                                        SYS_Etrust.WhetherShou = true;
                                    }
                                    if (SYS_Etrust.ShuiBl > 0)
                                    {
                                        SYS_Etrust.WhetherShui = true;
                                    }
                                    if (SYS_Etrust.ManagementFee > 0)
                                    {
                                        SYS_Etrust.WhetherBen = true;
                                    }
                                      mymodels.SYS_Etrust.Add(SYS_Etrust);
                                    //mymodels.SaveChanges();
                                    if (mymodels.SaveChanges()>0)
                                    {
                                            SYS_Charge Charge = new SYS_Charge();
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = offercount[0].money;
                                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                            Charge.ReckoningUnit = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                            if (SYS_Etrust.CollectionMoney>0)
                                        {
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = (offercount[0].money * SYS_Etrust.CollectionMoney)/100;
                                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                            Charge.ReckoningUnit = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                        }
                                        if (SYS_Etrust.ShuiBl > 0)
                                        {
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = (offercount[0].money * SYS_Etrust.ShuiBl) / 100;
                                            Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                            Charge.ReckoningUnits = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                        }
                                        if (SYS_Etrust.ManagementFee > 0)
                                        {
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = (offercount[0].money * SYS_Etrust.ManagementFee) / 100;
                                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                            Charge.ReckoningUnit = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                        }
                                           mymodels.SaveChanges();
                                            Data = "新增委托单成功";
                                }
                            }
                            else
                            {
                                var offercounts = OfferCount.Where(m => m.OfferType == "客户标准运费").ToList();
                                SYS_Etrust.OfferDetailID = offercounts[0].offerDetailID;
                                if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    SYS_Etrust.WhetherShou = true;
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    SYS_Etrust.WhetherShui = true;
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    SYS_Etrust.WhetherBen = true;
                                }
                            mymodels.SYS_Etrust.Add(SYS_Etrust);
                                //mymodels.SaveChanges();
                            if (mymodels.SaveChanges() > 0)
                                {
                                            SYS_Charge Charge = new SYS_Charge();
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = offercounts[0].money;
                                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                            Charge.ReckoningUnit = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                        if (SYS_Etrust.CollectionMoney > 0)
                                        {
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.CollectionMoney) / 100;
                                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                            Charge.ReckoningUnit = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                             mymodels.SaveChanges();
                                        }
                                        if (SYS_Etrust.ShuiBl > 0)
                                        {
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.ShuiBl) / 100;
                                            Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                            Charge.ReckoningUnits = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                        }
                                        if (SYS_Etrust.ManagementFee > 0)
                                        {
                                            Charge.EtrustID = SYS_Etrust.EtrustID;
                                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                            Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.ManagementFee) / 100;
                                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                            Charge.ReckoningUnit = null;
                                            mymodels.SYS_Charge.Add(Charge);
                                            mymodels.SaveChanges();
                                        }
                                            mymodels.SaveChanges();
                                            Data = "新增委托单成功";
                                }
                            }
                        }
                        else
                        {
                            if (SYS_Etrust.CollectionMoney > 0)
                            {
                                SYS_Etrust.WhetherShou = true;
                            }
                            if (SYS_Etrust.ShuiBl > 0)
                            {
                                SYS_Etrust.WhetherShui = true;
                            }
                            if (SYS_Etrust.ManagementFee > 0)
                            {
                                SYS_Etrust.WhetherBen = true;
                            }
                           mymodels.SYS_Etrust.Add(SYS_Etrust);
                            //mymodels.SaveChanges();
                            if (mymodels.SaveChanges() > 0)
                            {
                                SYS_Charge Charge = new SYS_Charge();
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                Charge.ReckoningUnit = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                            if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                    Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                    Charge.ReckoningUnits = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                 mymodels.SaveChanges();
                                Data = "新增委托单成功,无报价";
                        }
                      }
                }
                else
                {
                    Data = "委托单号重复";      
                }
            }
            catch (Exception)
            {
                Data = "新增委托单失败";
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //新增客户
        public ActionResult InsertClient(SYS_Client SYS_Client,bool? WTDW,bool? FHR,bool? CGS, bool? HZDL, bool? SHR, bool? TCGS, bool? BXGS, bool? CDGS, bool? TZR, bool? WLGS, bool? BGH)
        {
                var Data = "新增失败";
                try
                {
                    if (mymodels.SYS_Client.Where(m => m.ClientCode.Trim() == SYS_Client.ClientCode.Trim()).Count() == 0)
                    {
                        if (mymodels.SYS_Client.Where(m => m.ClientAbbreviation.Trim() == SYS_Client.ClientAbbreviation.Trim()).Count() == 0)
                        {
                            SYS_Client.StaffID = Convert.ToInt32(Session["StaffID"].ToString());
                            SYS_Client.WhetherStart = true;
                            mymodels.SYS_Client.Add(SYS_Client);
                            SYS_ClientType SYS_ClientType = new SYS_ClientType();
                            if (mymodels.SaveChanges() > 0)
                            {

                                if (WTDW == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "委托单位";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (FHR == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "发货人";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (CGS == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "船公司";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (HZDL == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "合作代理";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (SHR == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "收货人";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (TCGS == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "拖车公司";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (BXGS == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "保险公司";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (CDGS == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "船代公司";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (TZR == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "通知人";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (WLGS == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "物流公司";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                if (BGH == true)
                                {
                                    SYS_ClientType.ClientID = SYS_Client.ClientID;
                                    SYS_ClientType.ClientType = "报关行";
                                    mymodels.SYS_ClientType.Add(SYS_ClientType);
                                    mymodels.SaveChanges();
                                };
                                Data = "新增客户成功";
                            }
                        }
                        else
                        {
                            Data = "客户简称重复";
                        }
                    }
                    else
                    {
                        Data = "客户代码重复";
                    }
                }
                catch (Exception)
                {
                    Data = "新增失败";
                }
                return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //新增客户联系人
        public ActionResult InsertClientScontacts(SYS_ClientScontacts SYS_ClientScontacts)
        {
                var Data = "新增失败";
                try
                {
                    if (mymodels.SYS_ClientScontacts.Where(m => m.contactsPhone.Trim() == SYS_ClientScontacts.contactsPhone.Trim()).Count() == 0)
                    {
                        mymodels.SYS_ClientScontacts.Add(SYS_ClientScontacts);
                        if (mymodels.SaveChanges() > 0)
                        {
                            Data = "新增客户联系人成功";
                        }
                    }
                    else
                    {
                        Data = "联系人号码重复";
                    }
                }
                catch (Exception)
                {

                    Data = "新增失败";
                }
                return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //新增客户地址
        public ActionResult InsertClientSite(SYS_ClientSite SYS_ClientSite)
       {
            var Data = "";
            try
            {
                if (mymodels.SYS_ClientSite.Where(m=>m.FactoryCode.Trim()==SYS_ClientSite.FactoryCode.Trim()).Count()==0)
                {
                    mymodels.SYS_ClientSite.Add(SYS_ClientSite);
                    mymodels.SaveChanges();
                    Data = "新增客户地址成功";
                }
                else
                {
                    Data = "工厂代码重复";
                }
            }
            catch (Exception)
            {
                Data = "新增客户地址失败";
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //新增港口
        public ActionResult InsertPort(SYS_Port SYS_Port,string Phones)
        {
            var Data = "";
            try
            {
                if (mymodels.SYS_Port.Where(m => m.PortCode.Trim() == SYS_Port.PortCode.Trim()).Count() == 0)
                    {
                        SYS_Port.WhetherStart = true;
                        SYS_Port.Phone = Phones;
                        mymodels.SYS_Port.Add(SYS_Port);
                        mymodels.SaveChanges();
                        Data = "新增港口成功";
                    }
                else
                    {
                            Data = "港口代码重复";
                    }
            }
            catch (Exception)
            {

                Data = "新增港口失败";
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //新增收付
        public ActionResult InsertSF(string myExpense, string myMoney, int EtrustID,string mySettleAccounts) {
            var Data =0 ;
            try
            {
                string[] Expense = myExpense.Split(',');
                string[] Money = myMoney.Split(',');
                string[] SettleAccounts = mySettleAccounts.Split(',');
                for (int i = 0; i < Expense.Count(); i++)
                {
                    if (Convert.ToDecimal(Money[i]) > 0)
                    {
                        var ExpenseID = Convert.ToInt32(Expense[i]);
                        var N = mymodels.SYS_Charge.Where(m => m.EtrustID == EtrustID && m.ExpenseID == ExpenseID).ToList();
                        if (N.Count() > 0)
                        {
                            N[0].UnitPrice = Convert.ToDecimal(Money[i]);
                            mymodels.Entry(N[0]).State = System.Data.Entity.EntityState.Modified;
                            mymodels.SaveChanges();
                            Data++;
                        }
                        else
                        {
                            SYS_Charge C = new SYS_Charge();
                            C.EtrustID = EtrustID;
                            C.ExpenseID = ExpenseID;
                            C.UnitPrice = Convert.ToDecimal(Money[i]);
                            if (SettleAccounts[i].Trim() == "应收")
                            {
                                C.ReckoningUnits = mymodels.SYS_Etrust.Where(m => m.EtrustID == EtrustID).Select(m => m.ClientCode.Trim()).Single();
                            }
                            else if (SettleAccounts[i].Trim() == "应付") {
                                C.ReckoningUnit = (from tbE in mymodels.SYS_Etrust
                                                   join tbM in mymodels.SYS_Message on tbE.MessageID equals tbM.MessageID
                                                   where tbE.EtrustID == EtrustID
                                                   select tbM.TissueCode.Trim()).Single();
                            }
                            else if (SettleAccounts[i].Trim() == "成本")
                            {
                                C.ReckoningUnit = (from tbE in mymodels.SYS_Etrust
                                                   join tbM in mymodels.SYS_Message on tbE.MessageID equals tbM.MessageID
                                                   where tbE.EtrustID ==EtrustID
                                                   select tbM.TissueCode.Trim()).Single();
                            }
                            mymodels.SYS_Charge.Add(C);
                            mymodels.SaveChanges();
                            Data++;
                        }
                    }
                    else
                    {
                        var ExpenseID = Convert.ToInt32(Expense[i]);
                        var N = mymodels.SYS_Charge.Where(m => m.EtrustID == EtrustID && m.ExpenseID == ExpenseID).ToList();
                        if (N.Count()>0)
                        {
                            mymodels.SYS_Charge.Remove(N[0]);
                            mymodels.SaveChanges();
                            Data++;
                        }
                    }
                }
            }
            catch (Exception)
            {
                
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //审核
        public ActionResult Audit(string myEtrust) {
            var Data = 0;
            string[] Etrust = myEtrust.ToString().Split(',');
            for (int i = 0; i < Etrust.Count(); i++)
            {
                int EtrustID = Convert.ToInt32(Etrust[i]);
                var S = mymodels.SYS_Etrust.Where(m => m.EtrustID == EtrustID).Select(m => m.EtrustType.Trim()).Single();
                if (mymodels.SYS_Etrust.Where(m=>m.EtrustID==EtrustID).Select(m=>m.EtrustType.Trim()).Single()!="新单")
                {
                    if (mymodels.SYS_Etrust.Where(m => m.EtrustID == EtrustID).Select(m => m.AuditType.Trim()).Single() == "未审核")
                    {
                        var c = 0;
                        var change = mymodels.SYS_Charge.Where(m => m.EtrustID == EtrustID).Select(m => m.UnitPrice).ToList();
                        for (int n = 0; n < change.Count(); n++)
                        {
                            if (change[n] > 0)
                            {
                                c++;
                            }
                        }
                        if (change.Count() == c)
                        {
                            var E = mymodels.SYS_Etrust.Where(m => m.EtrustID == EtrustID).Single();
                            E.AuditType = "审核中";
                            mymodels.Entry(E).State = System.Data.Entity.EntityState.Modified;
                            if (mymodels.SaveChanges() > 0)
                            {
                                SYS_Commercel myCommercel = new SYS_Commercel();
                                myCommercel.EtrustID = E.EtrustID;
                                mymodels.SYS_Commercel.Add(myCommercel);
                                mymodels.SaveChanges();
                                Data++;
                            }
                        }
                        else
                        {
                            Data = -100000;
                        }
                    }
                    else {
                        Data = -200000;
                    }
                }
            }
            return Json(Data,JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region 修改
        //修改委托单
        public ActionResult UpdateEtrust(SYS_Etrust SYS_Etrust,int? ClientID) {
            var Data = "";
            try
            {
                var chagre = mymodels.SYS_Charge.Where(m => m.EtrustID == SYS_Etrust.EtrustID).ToList();
                var chagres = (from tbcharge in mymodels.SYS_Charge
                              join tbExpense in mymodels.SYS_Expense on tbcharge.ExpenseID equals tbExpense.ExpenseID
                              where tbcharge.EtrustID == SYS_Etrust.EtrustID&&tbExpense.ExpenseName.Trim()=="运费"&& tbExpense.SettleAccounts.Trim() == "应收"
                               select tbcharge ).ToList();
                var chagref = (from tbcharge in mymodels.SYS_Charge
                               join tbExpense in mymodels.SYS_Expense on tbcharge.ExpenseID equals tbExpense.ExpenseID
                               where tbcharge.EtrustID == SYS_Etrust.EtrustID && tbExpense.ExpenseName.Trim() == "运费" && tbExpense.SettleAccounts.Trim() == "应付"
                               select tbcharge).ToList();
                var chagrej = (from tbcharge in mymodels.SYS_Charge
                               join tbExpense in mymodels.SYS_Expense on tbcharge.ExpenseID equals tbExpense.ExpenseID
                               where tbcharge.EtrustID == SYS_Etrust.EtrustID && tbExpense.ExpenseName.Trim() == "司机提成" && tbExpense.SettleAccounts.Trim() == "应付"
                               select tbcharge).ToList();
                for (int i = 0; i < chagre.Count(); i++)
                {
                    mymodels.SYS_Charge.Remove(chagre[i]);
                    mymodels.SaveChanges();
                }
                SYS_Etrust.StaffID = Convert.ToInt32(Session["StaffID"].ToString());
                SYS_Etrust.ClientCode = mymodels.SYS_Client.Where(m => m.ClientID == ClientID).Select(m => m.ClientCode.Trim()).Single();
                SYS_Etrust.CheXing = "自备柜";
                SYS_Etrust.AuditType = "未审核";
                SYS_Etrust.Seal = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.Seal).Single();
                SYS_Etrust.BookingSpace = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.BookingSpace).Single();
                SYS_Etrust.Cupboard = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.Cupboard).Single();
                if (SYS_Etrust.CollectionMoney > 0)
                {
                    SYS_Etrust.WhetherShou = true;
                }
                if (SYS_Etrust.ShuiBl > 0)
                {
                    SYS_Etrust.WhetherShui = true;
                }
                if (SYS_Etrust.ManagementFee > 0)
                {
                    SYS_Etrust.WhetherBen = true;
                }
                #region
                if (mymodels.SYS_Etrust.Where(m=>m.EtrustID== SYS_Etrust.EtrustID).Select(m=>m.EtrustType.Trim()).Single()=="新单")
                {
                    
                    SYS_Etrust.EtrustType = "新单";
                    var S = false;
                    if (SYS_Etrust.ShuiBl > 0)
                    {
                        S = true;
                    }
                    var GatedotID = mymodels.SYS_ClientSite.Where(m => m.ClientSiteID == SYS_Etrust.ClientSiteID).Select(m => m.GatedotID).Single();
                    var OfferCount = (from tbOfferDetail in mymodels.SYS_OfferDetail
                                      join tbHaulWay in mymodels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                      join tbOffer in mymodels.SYS_Offer on tbOfferDetail.OfferID equals tbOffer.OfferID
                                      where tbHaulWay.MentionAreaID == SYS_Etrust.TiGuiDiDianID && tbHaulWay.AlsoTankAreaID == SYS_Etrust.HuaiGuiDiDianID && tbHaulWay.LoadingAreaID == GatedotID &&
                                      tbOffer.WhetherShui == S && tbOfferDetail.CabinetType == SYS_Etrust.CabinetType
                                      where tbOffer.OfferType.Trim() == "客户标准运费" || tbOffer.OfferType.Trim() == "客户应收费用"
                                      select new EtrustVo
                                      {
                                          money = tbOfferDetail.Money,
                                          clientID = tbOffer.ClientID,
                                          offerDetailID = tbOfferDetail.OfferDetailID,
                                          OfferType = tbOffer.OfferType.Trim()
                                      }).ToList();
                    if (OfferCount.Count() > 0)
                    {
                        var offercount = OfferCount.Where(m => m.clientID == ClientID).ToList();
                        if (offercount.Count() > 0)
                        {
                            SYS_Etrust.OfferDetailID = offercount[0].offerDetailID;
                            if (SYS_Etrust.CollectionMoney > 0)
                            {
                                SYS_Etrust.WhetherShou = true;
                            }
                            if (SYS_Etrust.ShuiBl > 0)
                            {
                                SYS_Etrust.WhetherShui = true;
                            }
                            if (SYS_Etrust.ManagementFee > 0)
                            {
                                SYS_Etrust.WhetherBen = true;
                            }
                            mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                            if (mymodels.SaveChanges() > 0)
                            {
                                SYS_Charge Charge = new SYS_Charge();
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                Charge.UnitPrice = offercount[0].money;
                                Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                Charge.ReckoningUnit = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                                if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = (offercount[0].money * SYS_Etrust.CollectionMoney) / 100;
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = (offercount[0].money * SYS_Etrust.ShuiBl) / 100;
                                    Charge.ReckoningUnits = null;
                                    Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = (offercount[0].money * SYS_Etrust.ManagementFee) / 100;
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                mymodels.SaveChanges();
                                Data = "修改委托单成功";
                            }
                        }
                        else
                        {
                            var offercounts = OfferCount.Where(m => m.OfferType == "客户标准运费").ToList();
                            SYS_Etrust.OfferDetailID = offercounts[0].offerDetailID;
                            if (SYS_Etrust.CollectionMoney > 0)
                            {
                                SYS_Etrust.WhetherShou = true;
                            }
                            if (SYS_Etrust.ShuiBl > 0)
                            {
                                SYS_Etrust.WhetherShui = true;
                            }
                            if (SYS_Etrust.ManagementFee > 0)
                            {
                                SYS_Etrust.WhetherBen = true;
                            }
                            mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                            if (mymodels.SaveChanges() > 0)
                            {
                                SYS_Charge Charge = new SYS_Charge();
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                Charge.UnitPrice = offercounts[0].money;
                                Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                Charge.ReckoningUnit = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                                if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.CollectionMoney) / 100;
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.ShuiBl) / 100;
                                    Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                    Charge.ReckoningUnits = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.ManagementFee) / 100;
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                mymodels.SaveChanges();
                                Data = "修改委托单成功";
                            }
                        }
                    }
                    else
                    {
                        SYS_Etrust.OfferDetailID = null;
                        if (SYS_Etrust.CollectionMoney > 0)
                        {
                            SYS_Etrust.WhetherShou = true;
                        }
                        if (SYS_Etrust.ShuiBl > 0)
                        {
                            SYS_Etrust.WhetherShui = true;
                        }
                        if (SYS_Etrust.ManagementFee > 0)
                        {
                            SYS_Etrust.WhetherBen = true;
                        }
                        mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                        if (mymodels.SaveChanges() > 0)
                        {
                            SYS_Charge Charge = new SYS_Charge();
                            Charge.EtrustID = SYS_Etrust.EtrustID;
                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费"&&m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                            Charge.ReckoningUnit = null;
                            mymodels.SYS_Charge.Add(Charge);
                            mymodels.SaveChanges();
                            if (SYS_Etrust.CollectionMoney > 0)
                            {
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                Charge.ReckoningUnit = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                            }
                            if (SYS_Etrust.ShuiBl > 0)
                            {
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                Charge.ReckoningUnits = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                            }
                            if (SYS_Etrust.ManagementFee > 0)
                            {
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                Charge.ReckoningUnit = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                            }
                            mymodels.SaveChanges();
                            Data = "修改委托单成功,无报价";
                        }
                    }

                }
                #endregion
                else
                {
                    var TiGuiDiDianID = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.TiGuiDiDianID).Single();
                    var HuaiGuiDiDianID = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.HuaiGuiDiDianID).Single();
                    var ClientSiteID = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.ClientSiteID).Single();
                    var PlanHandTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.PlanHandTime).Single();
                    var WhetherShou = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.WhetherShou).Single();
                    var WhetherShui = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.WhetherShui).Single();
                    var WhetherBen = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.WhetherBen).Single();
                    var CabinetType = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.CabinetType.Trim()).Single();

                    if (    TiGuiDiDianID !=SYS_Etrust.TiGuiDiDianID ||
                            HuaiGuiDiDianID != SYS_Etrust.HuaiGuiDiDianID ||
                            ClientSiteID != SYS_Etrust.ClientSiteID ||
                            PlanHandTime != SYS_Etrust.PlanHandTime ||
                            WhetherShou != SYS_Etrust.WhetherShou ||
                            WhetherShui != SYS_Etrust.WhetherShui ||
                            WhetherBen != SYS_Etrust.WhetherBen ||
                            CabinetType!=SYS_Etrust.CabinetType.Trim()
                       )
                    {
                        SYS_Etrust.EtrustType = "新单";
                        var S = false;
                        if (SYS_Etrust.ShuiBl > 0)
                        {
                            S = true;
                        }
                        var GatedotID = mymodels.SYS_ClientSite.Where(m => m.ClientSiteID == SYS_Etrust.ClientSiteID).Select(m => m.GatedotID).Single();
                        var OfferCount = (from tbOfferDetail in mymodels.SYS_OfferDetail
                                          join tbHaulWay in mymodels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                          join tbOffer in mymodels.SYS_Offer on tbOfferDetail.OfferID equals tbOffer.OfferID
                                          where tbHaulWay.MentionAreaID == SYS_Etrust.TiGuiDiDianID && tbHaulWay.AlsoTankAreaID == SYS_Etrust.HuaiGuiDiDianID && tbHaulWay.LoadingAreaID == GatedotID &&
                                          tbOffer.WhetherShui == S && tbOfferDetail.CabinetType == SYS_Etrust.CabinetType
                                          where tbOffer.OfferType.Trim() == "客户标准运费" || tbOffer.OfferType.Trim() == "客户应收费用"
                                          select new EtrustVo
                                          {
                                              money = tbOfferDetail.Money,
                                              clientID = tbOffer.ClientID,
                                              offerDetailID = tbOfferDetail.OfferDetailID,
                                              OfferType = tbOffer.OfferType.Trim()
                                          }).ToList();
                        if (OfferCount.Count() > 0)
                        {
                            var offercount = OfferCount.Where(m => m.clientID == ClientID).ToList();
                            if (offercount.Count() > 0)
                            {
                                SYS_Etrust.OfferDetailID = offercount[0].offerDetailID;
                                if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    SYS_Etrust.WhetherShou = true;
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    SYS_Etrust.WhetherShui = true;
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    SYS_Etrust.WhetherBen = true;
                                }
                                mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                                if (mymodels.SaveChanges() > 0)
                                {
                                    SYS_Charge Charge = new SYS_Charge();
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = offercount[0].money;
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                    if (SYS_Etrust.CollectionMoney > 0)
                                    {
                                        Charge.EtrustID = SYS_Etrust.EtrustID;
                                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                        Charge.UnitPrice = (offercount[0].money * SYS_Etrust.CollectionMoney) / 100;
                                        Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                        Charge.ReckoningUnit = null;
                                        mymodels.SYS_Charge.Add(Charge);
                                        mymodels.SaveChanges();
                                    }
                                    if (SYS_Etrust.ShuiBl > 0)
                                    {
                                        Charge.EtrustID = SYS_Etrust.EtrustID;
                                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                        Charge.UnitPrice = (offercount[0].money * SYS_Etrust.ShuiBl) / 100;
                                        Charge.ReckoningUnits = null;
                                        Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                        mymodels.SYS_Charge.Add(Charge);
                                        mymodels.SaveChanges();
                                    }
                                    if (SYS_Etrust.ManagementFee > 0)
                                    {
                                        Charge.EtrustID = SYS_Etrust.EtrustID;
                                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                        Charge.UnitPrice = (offercount[0].money * SYS_Etrust.ManagementFee) / 100;
                                        Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                        Charge.ReckoningUnit = null;
                                        mymodels.SYS_Charge.Add(Charge);
                                        mymodels.SaveChanges();
                                    }
                                    mymodels.SaveChanges();
                                    Data = "修改委托单成功";
                                }
                            }
                            else
                            {
                                var offercounts = OfferCount.Where(m => m.OfferType == "客户标准运费").ToList();
                                SYS_Etrust.OfferDetailID = offercounts[0].offerDetailID;
                                if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    SYS_Etrust.WhetherShou = true;
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    SYS_Etrust.WhetherShui = true;
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    SYS_Etrust.WhetherBen = true;
                                }
                                mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                                if (mymodels.SaveChanges() > 0)
                                {
                                    SYS_Charge Charge = new SYS_Charge();
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.UnitPrice = offercounts[0].money;
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                    if (SYS_Etrust.CollectionMoney > 0)
                                    {
                                        Charge.EtrustID = SYS_Etrust.EtrustID;
                                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                        Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.CollectionMoney) / 100;
                                        Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                        Charge.ReckoningUnit = null;
                                        mymodels.SYS_Charge.Add(Charge);
                                        mymodels.SaveChanges();
                                    }
                                    if (SYS_Etrust.ShuiBl > 0)
                                    {
                                        Charge.EtrustID = SYS_Etrust.EtrustID;
                                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                        Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.ShuiBl) / 100;
                                        Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                        Charge.ReckoningUnits = null;
                                        mymodels.SYS_Charge.Add(Charge);
                                        mymodels.SaveChanges();
                                    }
                                    if (SYS_Etrust.ManagementFee > 0)
                                    {
                                        Charge.EtrustID = SYS_Etrust.EtrustID;
                                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                        Charge.UnitPrice = (offercounts[0].money * SYS_Etrust.ManagementFee) / 100;
                                        Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                        Charge.ReckoningUnit = null;
                                        mymodels.SYS_Charge.Add(Charge);
                                        mymodels.SaveChanges();
                                    }
                                    mymodels.SaveChanges();
                                    Data = "修改委托单成功";
                                }
                            }
                        }
                        else
                        {
                            SYS_Etrust.OfferDetailID = null;
                            if (SYS_Etrust.CollectionMoney > 0)
                            {
                                SYS_Etrust.WhetherShou = true;
                            }
                            if (SYS_Etrust.ShuiBl > 0)
                            {
                                SYS_Etrust.WhetherShui = true;
                            }
                            if (SYS_Etrust.ManagementFee > 0)
                            {
                                SYS_Etrust.WhetherBen = true;
                            }
                            mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                            if (mymodels.SaveChanges() > 0)
                            {
                                SYS_Charge Charge = new SYS_Charge();
                                Charge.EtrustID = SYS_Etrust.EtrustID;
                                Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                Charge.ReckoningUnit = null;
                                mymodels.SYS_Charge.Add(Charge);
                                mymodels.SaveChanges();
                                if (SYS_Etrust.CollectionMoney > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ShuiBl > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                                    Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                                    Charge.ReckoningUnits = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                if (SYS_Etrust.ManagementFee > 0)
                                {
                                    Charge.EtrustID = SYS_Etrust.EtrustID;
                                    Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                                    Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                                    Charge.ReckoningUnit = null;
                                    mymodels.SYS_Charge.Add(Charge);
                                    mymodels.SaveChanges();
                                }
                                mymodels.SaveChanges();
                                Data = "修改委托单成功,无报价";
                            }
                        }
                    }
                    else
                    {
                        SYS_Etrust.EtrustType = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.EtrustType).Single();
                        SYS_Etrust.UndertakeID = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.UndertakeID).Single();
                        SYS_Etrust.CarriageNumber = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.CarriageNumber.Trim()).Single();
                        SYS_Etrust.VehicleInformationID = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.VehicleInformationID).Single();
                        SYS_Etrust.PaiCheDate = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.PaiCheDate).Single();
                        SYS_Etrust.PaiCheExplain = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.PaiCheExplain.Trim()).Single();
                        SYS_Etrust.DispatchTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.DispatchTime).Single();
                        SYS_Etrust.TiKGuiTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.TiKGuiTime).Single();
                        SYS_Etrust.ArriveFactoryTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.ArriveFactoryTime).Single();
                        SYS_Etrust.LeftFactoryTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.LeftFactoryTime).Single();
                        SYS_Etrust.HuaiWeightTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.HuaiWeightTime).Single();
                        SYS_Etrust.FollowTime = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Select(m => m.FollowTime).Single();
                        SYS_Charge Charge = new SYS_Charge();
                        Charge.EtrustID = SYS_Etrust.EtrustID;
                        Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                        if (chagres.Count()>0)
                        {
                            Charge.UnitPrice = chagres[0].UnitPrice;
                        }
                        else
                        {
                            Charge.UnitPrice = 0;
                        }
                        Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                        Charge.ReckoningUnit = null;
                        mymodels.SYS_Charge.Add(Charge);
                        mymodels.SaveChanges();
                        if (SYS_Etrust.CollectionMoney > 0)
                        {
                            Charge.EtrustID = SYS_Etrust.EtrustID;
                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "代收费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                            if (chagres.Count()>0)
                            {
                                Charge.UnitPrice = (chagres[0].UnitPrice * SYS_Etrust.CollectionMoney) / 100;
                            }
                            else
                            {
                                Charge.UnitPrice = 0;
                            }
                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                            Charge.ReckoningUnit = null;
                            mymodels.SYS_Charge.Add(Charge);
                            mymodels.SaveChanges();
                        }
                        if (SYS_Etrust.ShuiBl > 0)
                        {
                            Charge.EtrustID = SYS_Etrust.EtrustID;
                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "税金" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            if (chagres.Count()>0)
                            {
                                Charge.UnitPrice = (chagres[0].UnitPrice * SYS_Etrust.ShuiBl) / 100;
                            }
                            else
                            {
                                Charge.UnitPrice = 0;
                            }
                            Charge.ReckoningUnits = null;
                            Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                            mymodels.SYS_Charge.Add(Charge);
                            mymodels.SaveChanges();
                        }
                        if (SYS_Etrust.ManagementFee > 0)
                        {
                            Charge.EtrustID = SYS_Etrust.EtrustID;
                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "管理费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                            if (chagres.Count()>0)
                            {
                                Charge.UnitPrice = (chagres[0].UnitPrice * SYS_Etrust.ManagementFee) / 100;
                            }
                            else
                            {
                                Charge.UnitPrice = 0;
                            }
                            Charge.ReckoningUnits = SYS_Etrust.ClientCode;
                            Charge.ReckoningUnit = null;
                            mymodels.SYS_Charge.Add(Charge);
                            mymodels.SaveChanges();
                        }
                        if (SYS_Etrust.VehicleInformationID>0&&SYS_Etrust.UndertakeID>0)
                        {
                            Charge.EtrustID = SYS_Etrust.EtrustID;
                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "司机提成" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            if (chagrej.Count()>0)
                            {
                                Charge.UnitPrice = chagrej[0].UnitPrice;
                            }
                            else
                            {
                                Charge.UnitPrice = 0;
                            }
                            Charge.ReckoningUnits = null;
                            Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                            mymodels.SYS_Charge.Add(Charge);
                            mymodels.SaveChanges();
                        }
                        else if ( SYS_Etrust.UndertakeID > 0)
                        {
                            Charge.EtrustID = SYS_Etrust.EtrustID;
                            Charge.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            if (chagref.Count()>0)
                            {
                                Charge.UnitPrice = chagref[0].UnitPrice;
                            }
                            else
                            {
                                Charge.UnitPrice = 0;
                            }
                            Charge.ReckoningUnits = null;
                            Charge.ReckoningUnit = mymodels.SYS_Message.Where(m => m.MessageID == SYS_Etrust.MessageID).Select(m => m.TissueCode).Single();
                            mymodels.SYS_Charge.Add(Charge);
                            mymodels.SaveChanges();
                        }
                        if (SYS_Etrust.CollectionMoney > 0)
                        {
                            SYS_Etrust.WhetherShou = true;
                        }
                        if (SYS_Etrust.ShuiBl > 0)
                        {
                            SYS_Etrust.WhetherShui = true;
                        }
                        if (SYS_Etrust.ManagementFee > 0)
                        {
                            SYS_Etrust.WhetherBen = true;
                        }
                        mymodels.Entry(SYS_Etrust).State = System.Data.Entity.EntityState.Modified;
                    }
                }
            }
            catch (Exception)
            {
                Data = "修改委托单失败";
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //派车
        public ActionResult InsertSendCar(SYS_Etrust SYS_Etrust) {
            string Data = "";
            try
            {
                var Etrust = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Single();
                Etrust.UndertakeID = SYS_Etrust.UndertakeID;
                Etrust.VehicleInformationID = SYS_Etrust.VehicleInformationID;
                Etrust.PaiCheDate = SYS_Etrust.PaiCheDate;
                Etrust.OddNumber = SYS_Etrust.OddNumber;
                Etrust.CarriageNumber = SYS_Etrust.CarriageNumber;
                Etrust.EtrustType = "派车";
                var charge = (from tbcharge in mymodels.SYS_Charge
                              join tbExpense in mymodels.SYS_Expense on tbcharge.ExpenseID equals tbExpense.ExpenseID
                              where tbExpense.ExpenseName.Trim() == "运费" && tbExpense.SettleAccounts.Trim() == "应付"
                              where tbcharge.EtrustID == SYS_Etrust.EtrustID
                              select tbcharge).ToList();
                var charges = (from tbcharge in mymodels.SYS_Charge
                              join tbExpense in mymodels.SYS_Expense on tbcharge.ExpenseID equals tbExpense.ExpenseID
                              where tbExpense.ExpenseName.Trim() == "司机提成" && tbExpense.SettleAccounts.Trim() == "应付"
                              where tbcharge.EtrustID == SYS_Etrust.EtrustID
                              select tbcharge).ToList();
                if (charge.Count()>0)
                {
                    for (int i = 0; i < charge.Count(); i++)
                    {
                        mymodels.SYS_Charge.Remove(charge[i]);
                        mymodels.SaveChanges();
                    }
                }
                if (charges.Count() > 0)
                {
                    for (int i = 0; i < charges.Count(); i++)
                    {
                        mymodels.SYS_Charge.Remove(charges[i]);
                        mymodels.SaveChanges();
                    }
                }
                mymodels.Entry(Etrust).State = System.Data.Entity.EntityState.Modified;
                if (mymodels.SaveChanges() > 0)
                {
                    if (SYS_Etrust.UndertakeID > 0 && SYS_Etrust.VehicleInformationID > 0)
                    {
                        var GatedotID = mymodels.SYS_ClientSite.Where(m => m.ClientSiteID == Etrust.ClientSiteID).Select(m => m.GatedotID).Single();
                        var OfferCount = (from tbOfferDetail in mymodels.SYS_OfferDetail
                                          join tbHaulWay in mymodels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                          join tbOffer in mymodels.SYS_Offer on tbOfferDetail.OfferID equals tbOffer.OfferID
                                          where tbHaulWay.MentionAreaID == Etrust.TiGuiDiDianID && tbHaulWay.AlsoTankAreaID == Etrust.HuaiGuiDiDianID && tbHaulWay.LoadingAreaID == GatedotID &&
                                          tbOffer.WhetherShui == Etrust.WhetherShui && tbOfferDetail.CabinetType == Etrust.CabinetType
                                          where tbOffer.OfferType.Trim() == "司机产值"
                                          select tbOfferDetail).ToList();
                        SYS_Charge C = new SYS_Charge();
                        if (OfferCount.Count() > 0)
                        {
                            C.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "司机提成" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            C.UnitPrice = (((from tbe in mymodels.SYS_Etrust
                                             join tbv in mymodels.SYS_VehicleInformation on tbe.VehicleInformationID equals tbv.VehicleInformationID
                                             join tbc in mymodels.SYS_Chauffeur on tbv.ChauffeurID equals tbc.ChauffeurID
                                             select tbc.DeductRatio).Single() * OfferCount[0].Money) / 100);
                            C.ReckoningUnit = (from tbE in mymodels.SYS_Etrust
                                               join tbM in mymodels.SYS_Message on tbE.MessageID equals tbM.MessageID
                                               where tbE.EtrustID == Etrust.EtrustID
                                               select tbM.TissueCode.Trim()).Single();
                            C.EtrustID = Etrust.EtrustID;
                            mymodels.SYS_Charge.Add(C);
                            mymodels.SaveChanges();
                            Data = "派车成功";
                        }
                        else
                        {
                            C.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "司机提成" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            C.UnitPrice = 0;
                            C.ReckoningUnit = (from tbE in mymodels.SYS_Etrust
                                               join tbM in mymodels.SYS_Message on tbE.MessageID equals tbM.MessageID
                                               where tbE.EtrustID == Etrust.EtrustID
                                               select tbM.TissueCode.Trim()).Single();
                            C.EtrustID = Etrust.EtrustID;
                            mymodels.SYS_Charge.Add(C);
                            mymodels.SaveChanges();
                            Data = "派车成功,无报价";
                        }
                    }
                    else if (SYS_Etrust.UndertakeID > 0)
                    {
                        var ClientID = mymodels.SYS_Client.Where(m => m.ClientCode.Trim() == Etrust.ClientCode.Trim()).Select(m => m.ClientID).Single();
                        var GatedotID = mymodels.SYS_ClientSite.Where(m => m.ClientSiteID == Etrust.ClientSiteID).Select(m => m.GatedotID).Single();
                        var OfferCount = (from tbOfferDetail in mymodels.SYS_OfferDetail
                                          join tbHaulWay in mymodels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                          join tbOffer in mymodels.SYS_Offer on tbOfferDetail.OfferID equals tbOffer.OfferID
                                          where tbHaulWay.MentionAreaID == Etrust.TiGuiDiDianID && tbHaulWay.AlsoTankAreaID == Etrust.HuaiGuiDiDianID && tbHaulWay.LoadingAreaID == GatedotID &&
                                          tbOffer.WhetherShui == Etrust.WhetherShui && tbOfferDetail.CabinetType == Etrust.CabinetType
                                          where tbOffer.OfferType.Trim() == "车队标准运费" && tbOffer.ClientID == ClientID
                                          select tbOfferDetail).ToList();
                        SYS_Charge C = new SYS_Charge();
                        if (OfferCount.Count() > 0)
                        {
                            C.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            C.UnitPrice = OfferCount[0].Money;
                            C.ReckoningUnit = (from tbE in mymodels.SYS_Etrust
                                               join tbM in mymodels.SYS_Message on tbE.MessageID equals tbM.MessageID
                                               where tbE.EtrustID == Etrust.EtrustID
                                               select tbM.TissueCode.Trim()).Single();
                            C.EtrustID = Etrust.EtrustID;
                            mymodels.SYS_Charge.Add(C);
                            mymodels.SaveChanges();
                            Data = "派车成功";
                        }
                        else
                        {
                            C.ExpenseID = mymodels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();
                            C.UnitPrice = 0;
                            C.ReckoningUnit = (from tbE in mymodels.SYS_Etrust
                                               join tbM in mymodels.SYS_Message on tbE.MessageID equals tbM.MessageID
                                               where tbE.EtrustID == Etrust.EtrustID
                                               select tbM.TissueCode.Trim()).Single();
                            C.EtrustID = Etrust.EtrustID;
                            mymodels.SYS_Charge.Add(C);
                            mymodels.SaveChanges();
                            Data = "派车成功,无报价";
                        }
                    }
                }
            }
            catch (Exception)
            {
                Data = "派车失败";
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        //跟单
        public ActionResult InsertFollow(SYS_Etrust SYS_Etrust) {
            string Data = "";
            try
            {
                var Etrust = mymodels.SYS_Etrust.Where(m => m.EtrustID == SYS_Etrust.EtrustID).Single();
                Etrust.EtrustType = "运输完成";
                Etrust.DispatchTime = SYS_Etrust.DispatchTime;
                Etrust.TiKGuiTime = SYS_Etrust.TiKGuiTime;
                Etrust.ArriveFactoryTime = SYS_Etrust.ArriveFactoryTime;
                Etrust.LeftFactoryTime = SYS_Etrust.LeftFactoryTime;
                Etrust.HuaiWeightTime = SYS_Etrust.HuaiWeightTime;
                Etrust.FollowTime = SYS_Etrust.FollowTime;
                mymodels.Entry(Etrust).State = System.Data.Entity.EntityState.Modified;
                mymodels.SaveChanges();
                Data = "跟单成功";
            }
            catch (Exception)
            {
                Data = "跟单失败";
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region 删除
        public ActionResult DeleteEtrust(string myEtrust)
        {
            int Data = 0;
            try
            {
                string[] Etrust = myEtrust.ToString().Split(',');
                for (int i = 0; i < Etrust.Count(); i++)
                {
                    var EtrustID = Convert.ToInt32(Etrust[i]);
                    var AuditType = mymodels.SYS_Etrust.Where(m => m.EtrustID == EtrustID).Select(m => m.AuditType.Trim()).Single();
                    if (AuditType == "未审核"|| AuditType == "审核中")
                    {
                        int  charge = 0;
                        var Charge = mymodels.SYS_Charge.Where(m => m.EtrustID == EtrustID).ToList();
                        for (int m = 0; m < Charge.Count(); m++)
                        {
                            mymodels.SYS_Charge.Remove(Charge[m]);
                            if (mymodels.SaveChanges()>0)
                            {
                                charge++;
                            }
                        }
                        if (charge == Charge.Count())
                        {
                            var Commercel = mymodels.SYS_Commercel.Where(m => m.EtrustID == EtrustID).ToList();
                            if (Commercel.Count()>0)
                            {
                                for (int n = 0; n < Commercel.Count(); n++)
                                {
                                    mymodels.SYS_Commercel.Remove(Commercel[n]);
                                    mymodels.SaveChanges();
                                }
                            }
                            var etrust = mymodels.SYS_Etrust.Where(n => n.EtrustID == EtrustID).Single();
                            mymodels.SYS_Etrust.Remove(etrust);
                            mymodels.SaveChanges();
                            Data++;
                        }
                    }
                }
            }
            catch (Exception)
            {
                
            }
            return Json(Data, JsonRequestBehavior.AllowGet);
        }
        #endregion
        
        #region 下拉框
        //提还柜地下拉框
        public ActionResult Mention()
        {
            try
            {
                var Mention = from tbMention in mymodels.SYS_Mention
                              where tbMention.WhetherStart==true
                              select new
                              {
                                  id = tbMention.MentionID,
                                  name = tbMention.Abbreviation.Trim()
                              };
                return Json(Mention, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        //客户下拉框
        public ActionResult Client(int? ClientID,int? ClientAbbreviation)
        {
            try
            {
                if (ClientID > 0)
                {
                    var Client = from tbClient in mymodels.SYS_Client
                                 where tbClient.WhetherStart==true
                                 select new
                                 {
                                     id = tbClient.ClientID,
                                     name = tbClient.ClientCode.Trim()
                                 };
                    return Json(Client, JsonRequestBehavior.AllowGet);
                }
                else if (ClientAbbreviation > 0)
                {
                    var Client = from tbClient in mymodels.SYS_Client
                                 where tbClient.WhetherStart==true
                                 select new
                                 {
                                     id = tbClient.ClientID,
                                     name = tbClient.ClientAbbreviation.Trim()
                                 };
                    return Json(Client, JsonRequestBehavior.AllowGet);
                }
                else {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
        //员工下拉框
        public ActionResult Staff(int? MessageID,int? DepartmentID)
        {
            try
            {
                if (MessageID>0)
                {
                    var Staff = from tbStaff in mymodels.SYS_Staff
                                join tbDepartment in mymodels.SYS_Department on tbStaff.DepartmentID equals tbDepartment.DepartmentID
                                join tbMessage in mymodels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID
                                where tbDepartment.MessageID == MessageID
                                select new
                                {
                                    id = tbStaff.StaffID,
                                    name = tbStaff.StaffNumber.Trim()
                                };
                    return Json(Staff, JsonRequestBehavior.AllowGet);
                }
                else if (DepartmentID > 0)
                {
                    var Staff = from tbMention in mymodels.SYS_Staff
                                join tbDepartment in mymodels.SYS_Department on tbMention.DepartmentID equals tbDepartment.DepartmentID
                                where tbMention.DepartmentID == DepartmentID
                                select new
                                {
                                    id = tbMention.StaffID,
                                    name = tbMention.StaffNumber.Trim()
                                };
                    return Json(Staff, JsonRequestBehavior.AllowGet);
                }else
                {
                    return null;
                }
            }
            catch (Exception)
            {

                return null;
            }
        }
        //港口下拉框
        public ActionResult Port()
        {
            try
            {
                var Port = from tbPort in mymodels.SYS_Port
                              select new
                              {
                                  id = tbPort.PortID,
                                  name = tbPort.ChineseName.Trim()
                              };
                return Json(Port, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        //船舶下拉框
        public ActionResult Ship()
        {
            try
            {
                var Ship = from tbShip in mymodels.SYS_Ship
                           where tbShip.WhetherStart==true
                           select new
                           {
                               id = tbShip.ShipID,
                               name = tbShip.ChineseName.Trim()
                           };
                return Json(Ship, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        //客户联系人下拉框
        public ActionResult ClientScontacts(int? ClientID)
        {
            try
            {
                //if (ClientID > 0)
                //{
                    var ClientScontacts = from tbClientScontacts in mymodels.SYS_ClientScontacts
                                          join tbClient in mymodels.SYS_Client on tbClientScontacts.ClientID equals tbClient.ClientID
                                          where tbClientScontacts.ClientID == ClientID
                                          select new
                                          {
                                              id = tbClientScontacts.ClientScontactsID,
                                              name = tbClientScontacts.contactsName.Trim()
                                          };
                    return Json(ClientScontacts, JsonRequestBehavior.AllowGet);
                //}
                //else {
                //    var ClientScontacts = from tbClientScontacts in mymodels.SYS_ClientScontacts
                //                          select new
                //                          {
                //                              id = tbClientScontacts.ClientScontactsID,
                //                              name = tbClientScontacts.contactsName.Trim()
                //                          };
                //    return Json(ClientScontacts, JsonRequestBehavior.AllowGet);
                //}
            }
            catch (Exception)
            {

                return null;
            }
        }

        //客户地址下拉框
        public ActionResult ClientSite(int? ClientID)
        {
            try
            {
                //if (ClientID > 0)
                //{
                    var ClientSite = from tbClientSite in mymodels.SYS_ClientSite
                                     join tbClient in mymodels.SYS_Client on tbClientSite.ClientID equals tbClient.ClientID
                                     where tbClientSite.ClientID == ClientID
                                     select new
                                     {
                                         id = tbClientSite.ClientSiteID,
                                         name = tbClientSite.FactoryCode.Trim()
                                     };
                    return Json(ClientSite, JsonRequestBehavior.AllowGet);
                //}
                //else {
                //    var ClientSite = from tbClientSite in mymodels.SYS_ClientSite
                //                     select new
                //                     {
                //                         id = tbClientSite.ClientSiteID,
                //                         name = tbClientSite.FactoryCode.Trim()
                //                     };
                //    return Json(ClientSite, JsonRequestBehavior.AllowGet);
                //}
                
            }
            catch (Exception)
            {

                return null;
            }
        }
        //城市下拉框
        public ActionResult SelectCity(string Province)
        {
            var City = (from tbCity in mymodels.BS_City
                        where tbCity.Province.Trim() == Province.Trim()
                        select new
                        {
                            id = tbCity.CityID,
                            name = tbCity.CityName
                        }).ToList();
            return Json(City, JsonRequestBehavior.AllowGet);
        }
        //门点下拉框
        public ActionResult SelectGatedot(int? CityID)
        {
            var Gatedot = (from tbGatedot in mymodels.SYS_Gatedot
                        where tbGatedot.CityID == CityID
                        where tbGatedot.WhetherValid == true
                        select new
                        {
                            id = tbGatedot.GatedotID,
                            name = tbGatedot.GatedotName
                        }).ToList();
            return Json(Gatedot, JsonRequestBehavior.AllowGet);
        }
        //车辆下拉框
        public ActionResult SelectVehicleInformation(int? UndertakeID)
        {
            if (UndertakeID>0)
            {
                var VehicleInformation = (from tbVehicleInformation in mymodels.SYS_VehicleInformation
                                          join tbChauffeur in mymodels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                                          join tbClientType in mymodels.SYS_ClientType on tbChauffeur.ClientTypeID equals tbClientType.ClientTypeID
                                          where tbChauffeur.ClientTypeID == UndertakeID
                                          select new
                                          {
                                              id = tbVehicleInformation.VehicleInformationID,
                                              name = tbVehicleInformation.PlateNumbers
                                          }).ToList();
                return Json(VehicleInformation, JsonRequestBehavior.AllowGet);
            }
            else
            {
                var VehicleInformation = (from tbVehicleInformation in mymodels.SYS_VehicleInformation
                                          select new
                                          {
                                              id = tbVehicleInformation.VehicleInformationID,
                                              name = tbVehicleInformation.PlateNumbers
                                          }).ToList();
                return Json(VehicleInformation, JsonRequestBehavior.AllowGet);
            }
        }
        //承运公司下拉框 
        public ActionResult SelectUndertake() {
            var Undertake = (from tbUndertake in mymodels.SYS_ClientType
                             join tbClient in mymodels.SYS_Client on tbUndertake.ClientID equals tbClient.ClientID 
                             where tbUndertake.ClientType.Trim()=="拖车公司"
                            select new
                            {
                                id = tbUndertake.ClientTypeID,
                                name = tbClient.ClientAbbreviation
                            }).ToList();
            return Json(Undertake, JsonRequestBehavior.AllowGet);
        }
        //司机下拉框 
        public ActionResult SelectChauffeur()
        {
            var Undertake = (from tbChauffeur in mymodels.SYS_Chauffeur
                             select new
                             {
                                 id = tbChauffeur.ChauffeurID,
                                 name = tbChauffeur.ChauffeurNumber
                             }).ToList();
            return Json(Undertake, JsonRequestBehavior.AllowGet);
        }
        #endregion
    }
}